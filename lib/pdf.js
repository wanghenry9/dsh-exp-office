/**
 * PDF 适配器（阶段 5：读取、文本提取与结构识别）。
 *
 * 与前三个格式不同，PDF 不是 ZIP+XML，而是一套对象图 + 交叉引用表 + 内容流。
 * 这里用 `node:zlib` 手写解析，保持插件零运行时依赖的设计。
 *
 * 解析策略上做了一个明确的取舍：
 *   标准做法是先读 `startxref` → 交叉引用表 → 按引用遍历对象图。
 *   但现代 PDF 普遍使用**交叉引用流**（PDF 1.5+）与**对象流**（ObjStm），
 *   完整实现两者需要处理压缩对象图与增量更新链，代码量与出错面都很大。
 *   本实现对**整个文件做对象扫描**（匹配 `N G obj`），再从对象流里补充成员对象。
 *   对「读取」而言这更宽容：文件尾部被截断、或增量更新链有断裂时仍能读出内容，
 *   代价是失去了交叉引用的完整性校验能力 —— 因此 `validate()` 会把
 *   「交叉引用表未校验」如实列进 not_checked，而不是假装校验过了。
 *
 * 对齐开发要求 §八（PDF 功能规划与开发注意事项）。
 *
 * @module dsh-exp-office/pdf
 */

import zlib from 'node:zlib'
import { OfficeError } from './errors.js'
import { encodePng } from './image.js'

/** PDF 文件头。 */
const HEADER_RE = /^%PDF-(\d+\.\d+)/

/** 对象扫描：`N G obj`。 */
const OBJ_RE = /(\d+)\s+(\d+)\s+obj\b/g

/** 读取上限，防止恶意文件撑爆内存。 */
export const DEFAULT_PDF_LIMITS = Object.freeze({
  maxBytes: 512 * 1024 * 1024,
  maxObjects: 500000,
  maxStreamBytes: 256 * 1024 * 1024,
  maxPages: 20000
})

/** 一个已打开的 PDF 文档。 */
export class PdfDocument {
  #buffer
  #objects = new Map()
  #limits
  #cache = new Map()

  /**
   * @param {Buffer} buffer - 文件字节。
   * @param {Map<string, object>} objects - 对象表。
   * @param {object} limits - 上限。
   */
  constructor(buffer, objects, limits) {
    this.#buffer = buffer
    this.#objects = objects
    this.#limits = limits
  }

  /**
   * 打开一个 PDF。
   * @param {Buffer} buffer - 文件字节。
   * @param {Partial<typeof DEFAULT_PDF_LIMITS>} [limitOverrides] - 上限覆盖。
   * @returns {PdfDocument} 文档对象。
   */
  static open(buffer, limitOverrides = {}) {
    const limits = { ...DEFAULT_PDF_LIMITS, ...limitOverrides }
    if (!Buffer.isBuffer(buffer) || buffer.length < 8) {
      throw new OfficeError('CORRUPTED_DOCUMENT', '文件过小，不是有效的 PDF。')
    }
    if (buffer.length > limits.maxBytes) {
      throw new OfficeError('MEMORY_LIMIT', `PDF ${buffer.length} 字节超过上限 ${limits.maxBytes}。`)
    }
    const head = buffer.subarray(0, 1024).toString('latin1')
    const header = HEADER_RE.exec(head)
    if (!header) {
      throw new OfficeError('CORRUPTED_DOCUMENT', '缺少 %PDF- 文件头，不是有效的 PDF。')
    }

    const objects = scanObjects(buffer, limits)
    const doc = new PdfDocument(buffer, objects, limits)
    doc.info = { version: header[1] }
    return doc
  }

  /**
   * 取一个间接对象。
   * @param {number} num - 对象号。
   * @param {number} [gen] - 代号。
   * @returns {object|undefined} 解析后的对象。
   */
  resolve(num, gen = 0) {
    const direct = this.#objects.get(`${num} ${gen}`) ?? this.#objects.get(`${num} 0`)
    if (direct) return direct.value
    return undefined
  }

  /**
   * 按引用解析：接受 `{ref: 'N G R'}` 或直接值。
   * @param {unknown} value - 值或引用。
   * @param {number} [depth] - 递归深度上限。
   * @returns {unknown} 解析后的值。
   */
  deref(value, depth = 0) {
    if (depth > 32 || value === null || value === undefined) return value
    if (typeof value === 'object' && typeof value.ref === 'string') {
      const [num, gen] = value.ref.split(' ').map(Number)
      return this.deref(this.resolve(num, gen), depth + 1)
    }
    return value
  }

  /**
   * 取得（并缓存）一个流对象解码后的字节。
   * @param {object} streamObj - 带 `stream` 的解析对象。
   * @returns {Buffer} 解码后的字节。
   */
  decodeStream(streamObj) {
    if (!streamObj || streamObj.stream === undefined) {
      throw new OfficeError('INVALID_REQUEST', '该对象不是流对象。')
    }
    const cacheKey = streamObj.__id
    if (cacheKey && this.#cache.has(cacheKey)) return this.#cache.get(cacheKey)
    const dict = this.deref(streamObj.dict) ?? {}
    const filters = normalizeFilters(this.deref(dict.Filter))
    let data = streamObj.stream

    for (const filter of filters) {
      if (filter === 'FlateDecode' || filter === 'Fl') {
        try {
          data = zlib.inflateSync(data, { maxOutputLength: this.#limits.maxStreamBytes })
        } catch (err) {
          // 常见情形：流声明了 FlateDecode 但实际是原始数据（损坏或误声明）
          throw new OfficeError('CORRUPTED_DOCUMENT', `内容流解压失败：${err.message}`)
        }
      } else if (filter === 'ASCIIHexDecode' || filter === 'AHx') {
        data = Buffer.from(stripWhitespace(data.toString('latin1')).replace(/>$/, ''), 'hex')
      } else if (filter === 'ASCII85Decode' || filter === 'A85') {
        data = decodeAscii85(data.toString('latin1'))
      } else if (filter === 'DCTDecode' || filter === 'JPXDecode') {
        // 图片流：保持原始字节，不解码像素
        break
      } else {
        throw new OfficeError('UNSUPPORTED_FEATURE', `不支持的流过滤器：${filter}`, { filter })
      }
    }
    if (cacheKey && this.#cache.size < 64) this.#cache.set(cacheKey, data)
    return data
  }

  /**
   * 读取文档元数据。
   * @returns {object} 元数据。
   */
  metadata() {
    const infoDict = this.deref(this.#trailerInfoRef())
    const read = (key) => {
      const raw = infoDict ? infoDict[key] : undefined
      return raw === undefined ? null : decodePdfString(raw)
    }
    return {
      version: this.info.version,
      title: read('Title'),
      author: read('Author'),
      subject: read('Subject'),
      keywords: read('Keywords'),
      creator: read('Creator'),
      producer: read('Producer'),
      creation_date: read('CreationDate'),
      mod_date: read('ModDate'),
      trapped: read('Trapped')
    }
  }

  /**
   * 找出 trailer 里的 `/Info` 引用（取最新 trailer）。
   * @returns {unknown} Info 值或引用。
   */
  #trailerInfoRef() {
    return this.#trailerValue('Info')
  }

  /**
   * 枚举页面对象。
   *
   * 页面树可用时，**以 `/Kids` 为准**取页面的集合与顺序：这是 PDF 规范定义
   * 「文档有哪些页」的唯一依据。只按 `/Type /Page` 扫描会把两种对象也算进来：
   *   1. 被增量更新摘掉引用的旧页面对象（文件里还在，但已不是文档的页）；
   *   2. 其他工具遗留的孤儿页面对象。
   * 于是「删除一页」在重新打开后会看起来没生效 —— 所以按树过滤是必须的。
   *
   * 页面树断裂或引用不可解析时（截断文件、交叉引用受损），退回扫描全部
   * `/Type /Page` 对象并按对象号排序 —— 读得到内容比读不到更有用。
   *
   * @returns {object[]} 页面对象列表。
   */
  pages() {
    if (this.#pages) return this.#pages
    const refs = this.#kidRefs()
    let found = []
    if (refs) {
      for (const ref of refs) {
        const entry = this.#entryByRef(ref)
        if (entry && this.deref(entry.value.Type) === 'Page') {
          found.push({ key: entry.key, num: entry.num, gen: entry.gen, value: entry.value })
        }
      }
    }
    if (found.length === 0) {
      for (const [key, entry] of this.#objects) {
        const value = entry.value
        if (!value || typeof value !== 'object' || value.stream !== undefined) continue
        const type = this.deref(value.Type)
        if (type === 'Page') found.push({ key, num: entry.num, gen: entry.gen, value })
      }
      found.sort((a, b) => a.num - b.num)
    }
    if (found.length > this.#limits.maxPages) {
      throw new OfficeError('MEMORY_LIMIT', `PDF 页数 ${found.length} 超过上限 ${this.#limits.maxPages}。`)
    }
    this.#pages = found
    return found
  }

  #pages = null

  /**
   * 按 `"num gen"` 取对象表条目。
   * @param {string} ref - 引用。
   * @returns {object|undefined} 条目。
   */
  #entryByRef(ref) {
    const direct = this.#objects.get(ref)
    if (direct) return direct
    const [num] = ref.split(' ').map(Number)
    return this.#objects.get(`${num} 0`)
  }

  /**
   * 按 `/Kids` 顺序收集叶子页面引用。
   * @returns {string[]|null} `"num gen"` 列表；拿不到页面树时返回 null。
   */
  #kidRefs() {
    const root = this.#pagesRoot()
    if (!root) return null
    const out = []
    const walk = (node, depth) => {
      if (depth > 64 || !node || typeof node !== 'object') return
      const kids = this.deref(node.Kids)
      if (!Array.isArray(kids)) return
      for (const kid of kids) {
        const ref = typeof kid === 'object' && kid && typeof kid.ref === 'string' ? kid.ref : null
        if (!ref) continue
        const resolved = this.deref(kid)
        if (this.deref(resolved?.Type) === 'Pages') walk(resolved, depth + 1)
        else out.push(ref)
      }
    }
    walk(root.value, 0)
    return out.length > 0 ? out : null
  }

  /**
   * 读取每一页的尺寸、旋转与内容流概况。
   * @returns {object[]} 页面信息。
   */
  pageInfos() {
    return this.pages().map((page, index) => {
      const media = this.deref(page.value.MediaBox)
      const rotate = this.deref(page.value.Rotate)
      const contentsRef = page.value.Contents
      const streams = Array.isArray(this.deref(contentsRef)) ? this.deref(contentsRef).length : contentsRef === undefined ? 0 : 1
      return {
        index,
        object: `${page.num} ${page.gen}`,
        width_pt: Array.isArray(media) && media.length === 4 ? round2(Math.abs(Number(media[2]) - Number(media[0]))) : null,
        height_pt: Array.isArray(media) && media.length === 4 ? round2(Math.abs(Number(media[3]) - Number(media[1]))) : null,
        rotation: Number.isFinite(Number(rotate)) ? Number(rotate) : 0,
        content_streams: streams
      }
    })
  }

  /**
   * 提取文本。
   *
   * 先按 `BT`/`ET` 文本块切分，再解析 `Tj` / `TJ` / `'` / `"` 四类显示操作符。
   * 字符串编码通过字体的 `/ToUnicode` CMap 映射 —— 中文文档几乎都靠它，
   * 不做这一步中文会提取成乱码。
   *
   * @param {object} [options] - 选项。
   * @param {number} [options.maxChars] - 字符上限。
   * @param {number} [options.page] - 只提取某一页（0 基）。
   * @returns {object} `{text, length, truncated, pages, chars_per_page}`。
   */
  extractText({ maxChars = 200000, page = null } = {}) {
    const pages = this.pages()
    const targets = page === null ? pages : [pages[page]]
    if (targets.some((p) => p === undefined)) {
      throw new OfficeError('INVALID_REQUEST', `页面下标 ${page} 超出范围（共 ${pages.length} 页）。`, { page_count: pages.length })
    }

    const parts = []
    const charsPerPage = []
    let total = 0
    for (const target of targets) {
      const content = this.pageContent(target)
      const text = extractTextFromContent(this, content, target.value)
      charsPerPage.push(text.length)
      total += text.length
      parts.push(text)
    }
    const joined = parts.join('\n\f\n')
    return {
      text: joined.slice(0, maxChars),
      length: joined.length,
      truncated: joined.length > maxChars,
      pages: targets.length,
      chars_per_page: charsPerPage,
      total_chars: total
    }
  }

  /**
   * 提取带坐标的文本片段（版面分析用）。
   * @param {object} [options] - 选项。
   * @param {number|null} [options.page] - 只取某一页（0 基）。
   * @returns {object} `{page_count, runs, approximation}`。
   */
  textRuns({ page = null } = {}) {
    const pages = this.pages()
    const targets = page === null ? pages : [pages[page]]
    if (targets.some((p) => p === undefined)) {
      throw new OfficeError('INVALID_REQUEST', `页面下标 ${page} 超出范围（共 ${pages.length} 页）。`, { page_count: pages.length })
    }
    const runs = []
    targets.forEach((target, index) => {
      const pageIndex = page === null ? index : page
      const content = this.pageContent(target)
      for (const run of extractPositionedRuns(this, content, target.value, pageIndex)) runs.push(run)
    })
    return {
      page_count: targets.length,
      runs,
      approximation: [
        '坐标只取文本矩阵的平移分量，不还原旋转与缩放',
        '同一行内没有显式定位的片段，x 按「字符数 × 字号 × 0.5」估算推进'
      ]
    }
  }

  /**
   * 提取 PDF 里的表格（**纯位置推断**，不依赖框线）。
   *
   * 适用：文本可选、列大致对齐的表格（Office/LibreOffice 导出的表格、报表类 PDF）。
   * 不适用（会如实写进 `not_done`）：扫描件、只有图片的表格、跨页表格合并、合并单元格还原。
   *
   * @param {object} [options] - 选项。
   * @param {number|null} [options.page] - 只处理某一页（0 基）。
   * @param {number} [options.rowTolerance] - 同一行的 y 容差（点），默认 3。
   * @param {number} [options.columnGap] - 同一列的 x 容差（点），默认 8。
   * @param {number} [options.minRows] - 至少几行，默认 2。
   * @param {number} [options.minColumns] - 至少几列，默认 2。
   * @returns {object} 表格列表与推断说明。
   */
  extractTables({ page = null, rowTolerance = null, columnGap = 8, minRows = 2, minColumns = 2 } = {}) {
    for (const [name, value] of Object.entries({ columnGap, minRows, minColumns })) {
      if (!Number.isFinite(value) || value <= 0) throw new OfficeError('INVALID_REQUEST', `${name} 必须是正数，实际 ${value}。`)
    }
    if (rowTolerance !== null && (!Number.isFinite(rowTolerance) || rowTolerance <= 0)) {
      throw new OfficeError('INVALID_REQUEST', `rowTolerance 必须是正数或省略（省略时按行距自动推导），实际 ${rowTolerance}。`)
    }
    const { runs, approximation, page_count: pageCount } = this.textRuns({ page })
    const tables = extractTablesFromRuns(runs, { rowTolerance, columnGap, minRows, minColumns })
    return {
      page_count: pageCount,
      table_count: tables.length,
      tables,
      method: '按文本片段的位置推断：y 坐标聚类成行、x 坐标聚类成列（不依赖框线）',
      approximation,
      not_done: [
        '跨页表格自动合并（每页各自成表）',
        '合并单元格的跨行跨列还原',
        '扫描件与纯图片表格（没有文本层，需要 OCR —— 当前不支持）',
        '单元格内的换行与富文本样式'
      ]
    }
  }

  /**
   * 全文搜索。
   *
   * **位置是页内字符偏移，不是页面坐标**：要给出 x/y 必须把每个字形按文本矩阵
   * 与字体宽度算成版面坐标，那属于排版引擎的活；这里给的是可复现、可校验的
   * 「第几页 + 第几个字符 + 行号 + 上下文」，够定位与后续叠加写入，但不假装成坐标。
   *
   * 大小写不敏感走正则 `i` 标志而不是 `toLowerCase()` 比较：个别字符（如 `İ`）
   * 小写化后长度会变，用长度做偏移映射会把位置算错。
   *
   * @param {object} args - 参数。
   * @param {string} args.query - 要查找的文本。
   * @param {boolean} [args.caseSensitive] - 是否区分大小写，默认不区分。
   * @param {number} [args.maxResults] - 最多返回多少条命中，默认 100。
   * @param {number} [args.contextChars] - 命中前后各取多少字符作上下文，默认 40。
   * @param {number} [args.page] - 只搜某一页（0 基）；省略搜索全部页。
   * @returns {object} 命中列表与统计。
   */
  search({ query, caseSensitive = false, maxResults = 100, contextChars = 40, page = null }) {
    if (typeof query !== 'string' || query.length === 0) {
      throw new OfficeError('INVALID_REQUEST', 'query 不能为空。')
    }
    if (!Number.isInteger(maxResults) || maxResults < 1 || maxResults > 1000) {
      throw new OfficeError('INVALID_REQUEST', `maxResults 必须是 1–1000 的整数，实际 ${maxResults}。`)
    }
    if (!Number.isInteger(contextChars) || contextChars < 0 || contextChars > 1000) {
      throw new OfficeError('INVALID_REQUEST', `contextChars 必须是 0–1000 的整数，实际 ${contextChars}。`)
    }

    const pages = this.pages()
    const targets = page === null ? pages : [pages[page]]
    if (targets.some((p) => p === undefined)) {
      throw new OfficeError('INVALID_REQUEST', `页面下标 ${page} 超出范围（共 ${pages.length} 页）。`, { page_count: pages.length })
    }

    const pattern = new RegExp(escapeRegExp(query), caseSensitive ? 'g' : 'gi')
    const matches = []
    let total = 0
    let scannedChars = 0
    for (let i = 0; i < targets.length; i += 1) {
      const pageIndex = page === null ? i : page
      const text = extractTextFromContent(this, this.pageContent(targets[i]), targets[i].value)
      scannedChars += text.length
      pattern.lastIndex = 0
      let hit = pattern.exec(text)
      while (hit) {
        total += 1
        if (matches.length < maxResults) {
          const start = hit.index
          const end = start + hit[0].length
          matches.push({
            page: pageIndex,
            offset: start,
            line: text.slice(0, start).split('\n').length,
            match: hit[0],
            context_before: text.slice(Math.max(0, start - contextChars), start),
            context_after: text.slice(end, end + contextChars)
          })
        }
        // 零长匹配（不可能出现在转义后的字面量查询里）也要推进，否则死循环
        pattern.lastIndex = hit[0].length === 0 ? hit.index + 1 : pattern.lastIndex
        hit = pattern.exec(text)
      }
    }

    const perPage = {}
    for (const m of matches) perPage[m.page] = (perPage[m.page] ?? 0) + 1
    return {
      query,
      case_sensitive: caseSensitive,
      found: total > 0,
      total_matches: total,
      returned_matches: matches.length,
      truncated: total > matches.length,
      pages_scanned: targets.length,
      chars_scanned: scannedChars,
      matches_by_page: perPage,
      matches
    }
  }

  /**
   * 拼接并解码一页的全部内容流。
   * @param {object} page - 页面对象。
   * @returns {Buffer} 内容字节。
   */
  pageContent(page) {
    const contents = this.deref(page.value.Contents)
    const list = Array.isArray(contents) ? contents : contents === undefined ? [] : [page.value.Contents]
    const chunks = []
    for (const item of list) {
      const resolved = this.deref(item)
      if (!resolved || typeof resolved !== 'object' || resolved.stream === undefined) continue
      chunks.push(this.decodeStream(resolved))
    }
    return Buffer.concat(chunks)
  }

  /**
   * 结构识别：加密、签名、表单、批注、图片、字体与嵌入文件。
   * @returns {object} 结构统计。
   */
  structure() {
    const encryptRef = this.#trailerValue('Encrypt')
    const hasEncrypt = encryptRef !== undefined
    let annotations = 0
    let formFields = 0
    let pagesWithAnnots = 0
    for (const page of this.pages()) {
      const annots = this.deref(page.value.Annots)
      if (!Array.isArray(annots)) continue
      pagesWithAnnots += 1
      for (const annot of annots) {
        const dict = this.deref(annot)
        if (!dict || typeof dict !== 'object') continue
        annotations += 1
        if (this.deref(dict.Subtype) === 'Widget') formFields += 1
      }
    }
    const acroForm = this.deref(this.#catalogValue('AcroForm'))
    if (acroForm && typeof acroForm === 'object') {
      const fields = this.deref(acroForm.Fields)
      if (Array.isArray(fields) && fields.length > formFields) formFields = fields.length
    }

    let images = 0
    let fonts = new Set()
    let embeddedFiles = 0
    let hasSignature = false
    let objectStreams = 0
    for (const entry of this.#objects.values()) {
      const value = entry.value
      if (!value || typeof value !== 'object') continue
      // 流对象在对象表里是 `{dict, stream}` 包装，字典字段必须从 `.dict` 取 ——
      // 直接读 `value.Subtype` 永远读不到，图片与对象流的计数会**恒为 0**（真实踩过这个坑）。
      const dict = dictOf(value)
      const subtype = this.deref(dict.Subtype)
      const type = this.deref(dict.Type)
      if (subtype === 'Image') images += 1
      if (type === 'Font') {
        const base = this.deref(dict.BaseFont)
        if (typeof base === 'string') fonts.add(base)
      }
      if (type === 'ObjStm') objectStreams += 1
      if (type === 'Filespec') embeddedFiles += 1
      if (subtype === 'Widget' && this.deref(dict.FT) === 'Sig') hasSignature = true
    }
    if (this.#catalogValue('Perms') !== undefined) hasSignature = true

    return {
      page_count: this.pages().length,
      encrypted: hasEncrypt,
      encryption: hasEncrypt ? this.#encryptionInfo(encryptRef) : null,
      has_digital_signature: hasSignature,
      annotations,
      pages_with_annotations: pagesWithAnnots,
      form_fields: formFields,
      has_acroform: Boolean(acroForm && typeof acroForm === 'object'),
      images,
      font_count: fonts.size,
      fonts: [...fonts].slice(0, 50),
      embedded_files: embeddedFiles,
      object_streams: objectStreams,
      object_count: this.#objects.size
    }
  }

  /**
   * 校验 PDF 结构。
   * @returns {object} `{valid, checks, not_checked}`。
   */
  validate() {
    const checks = []
    checks.push({ name: 'PDF 文件头', ok: true, detail: `版本 ${this.info.version}` })
    checks.push({ name: '对象扫描', ok: this.#objects.size > 0, detail: `${this.#objects.size} 个间接对象` })

    let pages = []
    let pageError = null
    try {
      pages = this.pages()
    } catch (err) {
      pageError = err.message
    }
    checks.push({ name: '页面对象可解析', ok: pageError === null && pages.length > 0, detail: pageError ?? `${pages.length} 页` })

    const withSize = this.pageInfos().filter((p) => p.width_pt && p.height_pt).length
    checks.push({
      name: '页面尺寸可读',
      ok: pages.length === 0 || withSize === pages.length,
      detail: `${withSize}/${pages.length} 页含 MediaBox`
    })

    const structure = this.structure()
    checks.push({
      name: '未加密或已声明加密',
      ok: true,
      detail: structure.encrypted ? `已加密（${structure.encryption?.algorithm ?? '未知算法'}），内容流可能无法读取` : '未加密'
    })

    let textError = null
    let textLength = null
    try {
      textLength = this.extractText({ maxChars: 1 }).length
    } catch (err) {
      textError = err.message
    }
    checks.push({ name: '内容流可解码', ok: textError === null, detail: textError ?? '全部内容流解码成功' })

    return {
      valid: checks.every((c) => c.ok),
      checks,
      not_checked: PDF_NOT_CHECKED,
      structure,
      text_sample_length: textLength
    }
  }

  /**
   * 取 trailer 的某个值。
   *
   * **必须取最新的 trailer**：增量更新的文件里有多段 trailer，新的那段才带当前
   * `/Info`、`/Root`、`/Size`。取第一段会拿到「旧版本没写 /Info」之类的过期结论
   * （曾经真的因此让「新建 /Info」的写入在回读时看不见）。
   *
   * @param {string} key - 键。
   * @returns {unknown} 值。
   */
  #trailerValue(key) {
    const trailer = this.#newestTrailer()
    return trailer ? trailer[key] : undefined
  }

  /**
   * 找**当前生效**的那段 trailer。
   *
   * 权威依据是文件尾部的 `startxref`：它指向的要么是经典 `xref` 表（紧跟一段 `trailer`），
   * 要么是交叉引用流对象（trailer 信息就在那个流的字典里，而且 `/Info` 可能**直接内联**）。
   * 只按「有没有 trailer 关键字」猜会取错 —— 曾经因此把 `/Info` 丢掉，重写后元数据全空。
   *
   * 兜底顺序：startxref 读不到时，取偏移最大的经典 trailer；再没有就用交叉引用流的字典。
   *
   * @returns {object|undefined} trailer 字典。
   */
  #newestTrailer() {
    const startxref = readStartxref(this.#buffer)
    if (startxref !== null) {
      const head = this.#buffer.subarray(startxref, startxref + 32).toString('latin1')
      if (/^\s*xref\b/.test(head)) {
        const classic = this.#lastClassicTrailer()
        if (classic) return classic
      } else {
        const objMatch = /^(\d+)\s+(\d+)\s+obj/.exec(head)
        if (objMatch) {
          const entry = this.#entryByRef(`${objMatch[1]} ${objMatch[2]}`)
          const dict = entry?.value?.dict ?? entry?.value
          if (dict && typeof dict === 'object') return { ...dict, __trailer: true }
        }
      }
    }
    return this.#lastClassicTrailer() ?? this.#xrefStreamTrailer()
  }

  /**
   * 取偏移最大的经典 trailer（增量更新里最后写的那段）。
   * @returns {object|undefined} trailer 字典。
   */
  #lastClassicTrailer() {
    let best = null
    let bestOffset = -1
    for (const [key, entry] of this.#objects) {
      const value = entry.value
      if (!value || typeof value !== 'object' || !value.__trailer) continue
      if (key === 'trailer-from-xref') continue
      const offset = Number(key.split(' ')[1])
      if (Number.isFinite(offset) && offset > bestOffset) {
        bestOffset = offset
        best = value
      }
    }
    return best ?? undefined
  }

  /**
   * 取交叉引用流字典充当 trailer（PDF 1.5+ 没有 `trailer` 关键字时）。
   * @returns {object|undefined} trailer 字典。
   */
  #xrefStreamTrailer() {
    for (const [, entry] of this.#objects) {
      const value = entry.value
      if (value && typeof value === 'object' && value.__trailer) return value
    }
    return undefined
  }

  /**
   * 取文档目录（`/Root`）的某个值。
   * @param {string} key - 键。
   * @returns {unknown} 值。
   */
  #catalogValue(key) {
    for (const entry of this.#objects.values()) {
      const value = entry.value
      if (value && typeof value === 'object' && this.deref(value.Type) === 'Catalog') return value[key]
    }
    return undefined
  }

  /**
   * 读取加密字典概况。
   * @param {unknown} ref - Encrypt 引用。
   * @returns {object|null} 加密信息。
   */
  #encryptionInfo(ref) {
    const dict = this.deref(ref)
    if (!dict || typeof dict !== 'object') {
      // trailer 声明了 /Encrypt 但字典读不到（悬空引用、或被交叉引用流挡住）：
      // 这仍然是一条重要信息 —— 调用方需要知道「文件已加密且读不出参数」，
      // 返回 null 会让它误以为「没有加密信息」。
      return { filter: null, algorithm: '未知（加密字典不可读）', key_length_bits: null, readable: false }
    }
    const filter = this.deref(dict.Filter)
    const v = this.deref(dict.V)
    const lengthBits = this.deref(dict.Length)
    return {
      filter: typeof filter === 'string' ? filter : null,
      algorithm: v === 5 || v === 6 ? 'AES' : v === 4 ? 'RC4/AES（由 CF 决定）' : 'RC4',
      key_length_bits: Number.isFinite(Number(lengthBits)) ? Number(lengthBits) : 40,
      readable: true
    }
  }

  /**
   * 旋转一页。
   * @param {object} args - 参数。
   * @param {number} args.page - 页面下标（0 基）。
   * @param {number} args.degrees - 旋转角度（顺时针），归一到 0/90/180/270。
   * @returns {object} 变更信息。
   */
  rotatePage({ page, degrees }) {
    const target = this.#requirePage(page)
    const normalized = ((Math.round(degrees / 90) * 90) % 360 + 360) % 360
    const previous = Number(this.deref(target.value.Rotate) ?? 0) || 0
    target.value.Rotate = normalized
    this.#stage(`${target.num} ${target.gen}`, target.value)
    return { type: 'rotate_page', page, from: previous, to: normalized }
  }

  /**
   * 调整页面顺序。
   * @param {object} args - 参数。
   * @param {number} args.from - 原下标（0 基）。
   * @param {number} args.to - 目标下标（0 基）。
   * @returns {object} 变更信息。
   */
  movePage({ from, to }) {
    this.#requirePage(from)
    this.#requirePage(to)
    const order = this.pages().map((p) => `${p.num} ${p.gen}`)
    const [moved] = order.splice(from, 1)
    order.splice(to, 0, moved)
    this.#rewriteKids(order)
    return { type: 'move_page', from, to, order }
  }

  /**
   * 按给定顺序重排全部页面。
   *
   * `order` 描述的是**新页序**：`[2, 0, 1]` 表示新第 1 页是原第 3 页。
   * 这样调用方不需要自己算「移动一步」的下标偏移，重排结果也能被断言。
   *
   * @param {object} args - 参数。
   * @param {number[]} args.order - 原页序下标的排列。
   * @returns {object} 变更信息。
   */
  reorderPages({ order }) {
    const pages = this.pages()
    if (!Array.isArray(order) || order.length !== pages.length) {
      throw new OfficeError(
        'INVALID_REQUEST',
        `order 必须是长度等于页数（${pages.length}）的下标数组，实际 ${
          Array.isArray(order) ? order.length : typeof order
        }。`,
        { page_count: pages.length }
      )
    }
    const seen = new Set()
    for (const index of order) {
      if (!Number.isInteger(index) || index < 0 || index >= pages.length) {
        throw new OfficeError('INVALID_REQUEST', `order 含越界下标 ${index}（有效范围 0–${pages.length - 1}）。`, {
          page_count: pages.length
        })
      }
      if (seen.has(index)) {
        throw new OfficeError('INVALID_REQUEST', `order 含重复下标 ${index}；每个原页下标必须恰好出现一次。`)
      }
      seen.add(index)
    }
    const refs = order.map((index) => `${pages[index].num} ${pages[index].gen}`)
    this.#rewriteKids(refs)
    return { type: 'reorder_pages', order: [...order], page_count: refs.length }
  }

  /**
   * 在页面上叠加一行文本（水印 / 页码）。
   *
   * 这是「叠加式编辑」：**不改动原有内容流一个字节**，而是新建一个只画这一行字的内容流，
   * 追加到页面的 `/Contents` 之后。PDF 规定 `/Contents` 数组里的多个流按顺序在同一张画布上
   * 依次绘制，因此叠加是天然的 —— 这也是 §八.4 把叠加式列为稳定性最高写法的原因。
   *
   * 字体用 base-14 的 Helvetica（阅读器内置，**不需要嵌入字体**），代价是只支持
   * ASCII 可见字符：中文需要嵌入字体子集，本阶段明确不做，会直接报错而不是画出乱码。
   *
   * 透明度不做 ExtGState 软掩码，而用**浅灰填充**（`0.9 g` 之类）近似水印的淡化效果 ——
   * 结果稳定、不依赖透明度组的正确性。
   *
   * @param {object} args - 参数。
   * @param {string} args.text - 文本模板；支持 `{page}`（文档页码）、`{n}`（所选页中的序号）、`{total}`。
   * @param {number[]|null} [args.pages] - 目标页下标（0 基）；省略表示全部页。
   * @param {string} [args.position] - `center`（斜向居中）| `top` | `bottom`。
   * @param {number} [args.fontSize] - 字号；center 默认按页宽自动取值。
   * @param {number} [args.gray] - 灰度 0（黑）–1（白），默认 0.9 的浅灰。
   * @param {number} [args.margin] - top/bottom 时距页边的距离（pt），默认 24。
   * @param {(index: number, order: number) => string} [args.labelFor] - 逐页生成文本（页码场景用）。
   * @returns {object} 变更信息。
   */
  addTextOverlay({ text, pages = null, position = 'center', fontSize = null, gray = 0.9, margin = 24, labelFor = null }) {
    if (typeof text !== 'string' || text.length === 0) {
      throw new OfficeError('INVALID_REQUEST', 'text 不能为空。')
    }
    if (!['center', 'top', 'bottom'].includes(position)) {
      throw new OfficeError('INVALID_REQUEST', `position 只支持 center / top / bottom，实际 ${position}。`)
    }
    if (!(gray >= 0 && gray <= 1)) {
      throw new OfficeError('INVALID_REQUEST', `gray 必须在 0–1 之间，实际 ${gray}。`)
    }
    const all = this.pages()
    const targets = pages === null ? all.map((_, i) => i) : pages
    if (!Array.isArray(targets) || targets.length === 0) {
      throw new OfficeError('INVALID_REQUEST', 'pages 必须是非空下标数组，或省略表示全部页。')
    }
    for (const index of targets) this.#requirePage(index)

    const template = text
    const labelAt = (index, order) => {
      if (typeof labelFor === 'function') return String(labelFor(index, order))
      return template
        .replace(/\{page\}/g, String(index + 1))
        .replace(/\{n\}/g, String(order + 1))
        .replace(/\{total\}/g, String(all.length))
    }

    // 只需要一个 Helvetica 字体对象：所有页共用
    const fontRef = this.#addObject({ Type: 'Font', Subtype: 'Type1', BaseFont: 'Helvetica', Encoding: 'WinAnsiEncoding' })
    const touched = []
    for (const [order, index] of targets.entries()) {
      const label = labelAt(index, order)
      if (!/^[\x20-\x7E]+$/.test(label)) {
        throw new OfficeError(
          'UNSUPPORTED_FEATURE',
          `叠加文本只支持 ASCII 可见字符（第 ${index + 1} 页的内容为 ${JSON.stringify(label.slice(0, 20))}）：中文需要嵌入字体子集，当前阶段不做（画出来会是乱码，所以这里直接拒绝）。`,
          { solution: '改用英文/数字水印，或先在 Office 文档侧加水印再转 PDF。' }
        )
      }
      const page = this.#requirePage(index)
      const box = this.#pageBox(page)
      const width = box[2] - box[0]
      const height = box[3] - box[1]
      const size = fontSize ?? (position === 'center' ? Math.max(12, Math.min(width, height) / 8) : 12)
      const commands = buildOverlayCommands({ label, position, size, gray, margin, width, height })

      const streamRef = this.#addStream(commands)
      const resourcesRef = this.#withFontResource(page, fontRef)
      page.value.Contents = this.#appendContents(page.value.Contents, streamRef)
      page.value.Resources = { ref: resourcesRef }
      this.#stage(`${page.num} ${page.gen}`, page.value)
      touched.push(index)
    }
    return {
      type: 'text_overlay',
      position,
      pages: touched,
      page_count: all.length,
      note: '叠加式写入：原内容流未改动，新增的绘制流追加在 /Contents 之后。'
    }
  }

  /**
   * 给某一页添加一个「便签」批注（`/Subtype /Text`）。
   *
   * 与叠加文字（水印/页码）的关键区别：**批注文字由阅读器自己排版渲染**，
   * 不需要在 PDF 里嵌入字体，因此**支持中文**（内容按 UTF-16BE + BOM 写入 PDF 字符串，
   * 这是 PDF 表示非 ASCII 文本的标准做法）。水印走的是页面绘制流，才受「只能 ASCII」限制。
   *
   * 写入方式仍是增量更新：新建一个批注对象、把引用追加到该页的 `/Annots`，
   * 原页面内容流与其它对象一个字节都不动。`/Annots` 是间接引用（指向一个数组对象）时，
   * 会解引用后追加并把那个数组对象一并写回 —— 直接覆盖会让原有批注全部消失。
   *
   * 位置用 PDF 用户空间坐标（**原点在左下角**，单位 pt）。省略时放在页面左上角内侧。
   *
   * @param {object} args - 参数。
   * @param {number} args.page - 页面下标（0 基）。
   * @param {string} args.text - 批注内容（支持中文）。
   * @param {string} [args.author] - 作者（`/T`）。
   * @param {number} [args.x] - 左下角 x（pt）。
   * @param {number} [args.y] - 左下角 y（pt）。
   * @param {number} [args.width] - 图标宽（pt），默认 24。
   * @param {number} [args.height] - 图标高（pt），默认 24。
   * @param {number[]} [args.color] - RGB（0–1 三个数），默认淡黄。
   * @param {boolean} [args.open] - 是否默认展开便签，默认 false。
   * @returns {object} 变更信息。
   */
  addAnnotation({ page, text, author = null, x = null, y = null, width = 24, height = 24, color = [1, 0.85, 0.2], open = false }) {
    if (typeof text !== 'string' || text.length === 0) {
      throw new OfficeError('INVALID_REQUEST', 'text 不能为空。')
    }
    const target = this.#requirePage(page)
    for (const [name, value] of [['x', x], ['y', y], ['width', width], ['height', height]]) {
      if (value !== null && !Number.isFinite(value)) throw new OfficeError('INVALID_REQUEST', `${name} 必须是有限数字（pt）。`)
    }
    if (!(width > 0) || !(height > 0)) throw new OfficeError('INVALID_REQUEST', `width / height 必须为正数，实际 ${width} / ${height}。`)
    if (!Array.isArray(color) || color.length !== 3 || color.some((c) => !Number.isFinite(c) || c < 0 || c > 1)) {
      throw new OfficeError('INVALID_REQUEST', 'color 必须是 0–1 之间的三个数（RGB）。')
    }

    const box = this.#pageBox(target)
    const left = x === null ? Math.round(box[0] + 8) : x
    const bottom = y === null ? Math.round(box[3] - height - 8) : y
    const rect = [left, bottom, left + width, bottom + height]

    const dict = {
      Type: 'Annot',
      Subtype: 'Text',
      Rect: rect,
      Contents: encodePdfText(text),
      Name: 'Comment',
      C: color,
      // F=4 是 Print 标志：不加的话批注打印时会消失
      F: 4,
      ...(open ? { Open: true } : {}),
      ...(author ? { T: encodePdfText(String(author)) } : {})
    }
    const ref = this.#addObject(dict)

    const existing = target.value.Annots
    if (existing === undefined) {
      target.value.Annots = [{ ref }]
    } else if (Array.isArray(existing)) {
      existing.push({ ref })
    } else if (existing && typeof existing === 'object' && typeof existing.ref === 'string') {
      const list = this.deref(existing)
      if (!Array.isArray(list)) {
        throw new OfficeError('UNSUPPORTED_FEATURE', '/Annots 指向的不是数组，本工具不处理这种结构（拒绝写入而不是覆盖它）。')
      }
      list.push({ ref })
      this.#stage(existing.ref, list)
    } else {
      throw new OfficeError('UNSUPPORTED_FEATURE', `/Annots 的类型无法识别（${typeof existing}），拒绝覆盖。`)
    }
    this.#stage(`${target.num} ${target.gen}`, target.value)

    return {
      type: 'add_annotation',
      page,
      subtype: 'Text',
      rect,
      author,
      has_author: Boolean(author),
      open,
      text_length: text.length,
      non_ascii: !/^[\x00-\x7F]*$/.test(text),
      note: '增量更新：页面内容流未改动，只新增批注对象并追加到 /Annots。批注文字支持中文（由阅读器渲染，不需要嵌入字体）。'
    }
  }

  /**
   * 列出各页的批注（读回核对用）。
   * @returns {object} 批注列表与统计。
   */
  annotations() {
    const rows = []
    this.pages().forEach((page, index) => {
      const list = this.deref(page.value.Annots)
      if (!Array.isArray(list)) return
      for (const item of list) {
        const dict = this.deref(item)
        if (!dict || typeof dict !== 'object') continue
        rows.push({
          page: index,
          subtype: this.deref(dict.Subtype) ?? null,
          contents: decodePdfString(this.#pendingValueOf(dict.Contents)),
          author: decodePdfString(this.#pendingValueOf(dict.T)),
          rect: this.deref(dict.Rect) ?? null,
          flags: Number(this.deref(dict.F) ?? 0) || 0,
          open: this.deref(dict.Open) === true,
          color: this.deref(dict.C) ?? null,
          is_form_field: this.deref(dict.Subtype) === 'Widget'
        })
      }
    })
    return {
      count: rows.length,
      pages_with_annotations: new Set(rows.map((r) => r.page)).size,
      annotations: rows
    }
  }

  /**
   * 读取 AcroForm 表单字段（支持嵌套 `/Kids` 与被继承的 `/FT`）。
   *
   * 覆盖三类终端字段：文本（`Tx`）、按钮（`Btn`：复选框 / 单选组）、选择（`Ch`：列表 / 下拉）。
   * 按钮的「开状态名」从 `/AP /N` 的子键里取（排除 `Off`）——这个名字**不是固定的**，
   * 不同 PDF 生成器会写成 `Yes` / `1` / `开` 等，所以要读出来而不是猜。
   *
   * @returns {object} 表单结构与字段列表。
   */
  forms() {
    const walker = this.#collectFormFields()
    if (!walker) {
      return { has_form: false, field_count: 0, fields: [], need_appearances: null, note: '文档没有 AcroForm 表单。' }
    }
    const needAppearances = this.#pendingValueOf(walker.acro.NeedAppearances)
    return {
      has_form: true,
      field_count: walker.fields.length,
      need_appearances: needAppearances === true,
      signature_fields: walker.fields.filter((f) => f.ft === 'Sig').length,
      fields: walker.fields.map((f) => this.#describeField(f)),
      note: '值取自字段字典的 /V；复选框的 /AS 与 /V 可能不一致（/NeedAppearances 为真时由阅读器重绘）。'
    }
  }

  /**
   * 填写 AcroForm 表单字段。
   *
   * 只改字段字典的 `/V`（以及按钮的 `/AS`），并把 AcroForm 的 `/NeedAppearances` 置为 `true`
   * 让阅读器重绘外观。**不生成外观流**：那需要把文字按字体度量渲染成 XObject，
   * 属于排版引擎的活（做不到时宁可让阅读器重绘，也不画一个错的）。
   *
   * 明确不支持的形态会**报错而不是猜**：内联字典字段（没有独立对象号）、
   * 签名与按钮字段、只读字段（除非显式 `force`）。
   *
   * @param {object} args - 参数。
   * @param {object} args.fields - 字段名 → 值（文本字段用字符串，复选框用布尔或开状态名，选择字段用选项值）。
   * @param {boolean} [args.force] - 允许改写标记为只读的字段，默认 false。
   * @returns {object} 填写结果。
   */
  fillForm({ fields, force = false }) {
    if (!fields || typeof fields !== 'object' || Array.isArray(fields)) {
      throw new OfficeError('INVALID_REQUEST', 'fields 必须是「字段名 → 值」的对象。')
    }
    const names = Object.keys(fields)
    if (names.length === 0) throw new OfficeError('INVALID_REQUEST', 'fields 不能为空。')

    const walker = this.#collectFormFields()
    if (!walker) throw new OfficeError('UNSUPPORTED_FEATURE', '文档没有 AcroForm 表单，无法填写。')
    const byName = new Map(walker.fields.map((f) => [f.name, f]))
    const warnings = []

    // 第一遍：**只校验、不改动**。全部通过后才开始写，避免中途报错留下「填了一半」的文档
    // （调用方若忽略错误继续 save，半填状态就会被写进文件）。
    const plan = []
    for (const name of names) {
      const field = byName.get(name)
      if (!field) {
        throw new OfficeError('INVALID_REQUEST', `找不到表单字段「${name}」。`, { field_names: [...byName.keys()] })
      }
      if (!field.ref) {
        throw new OfficeError('UNSUPPORTED_FEATURE', `字段「${name}」是内联字典（没有独立对象号），本工具不改写这种结构。`)
      }
      if (field.ft === 'Sig') throw new OfficeError('UNSUPPORTED_FEATURE', `字段「${name}」是数字签名字段，不能填写。`)
      if (field.ft === 'Btn' && (field.flags & 65536) !== 0) {
        throw new OfficeError('UNSUPPORTED_FEATURE', `字段「${name}」是按钮（Pushbutton）字段，没有可填的值。`)
      }
      if (field.readOnly && !force) {
        throw new OfficeError('PERMISSION_DENIED', `字段「${name}」标记为只读（/Ff 第 1 位）。如确需改写请显式传 force=true。`, {
          needsConfirmation: true
        })
      }
      const value = fields[name]
      if (field.ft === 'Tx') {
        plan.push({ field, name, kind: 'Tx', value: value === null || value === undefined ? '' : String(value) })
      } else if (field.ft === 'Btn') {
        const on = field.onStates
        let target
        if (typeof value === 'boolean') target = value ? (on[0] ?? 'On') : 'Off'
        else if (value === null || value === undefined) target = 'Off'
        else target = String(value)
        if (target !== 'Off' && !on.includes(target)) {
          throw new OfficeError(
            'INVALID_REQUEST',
            `字段「${name}」的开状态只能是 ${on.map((s) => `「${s}」`).join(' / ')} 或 Off，实际「${target}」。`,
            { on_states: on }
          )
        }
        plan.push({ field, name, kind: 'Btn', target, on })
      } else if (field.ft === 'Ch') {
        const text = String(value)
        if (field.options.length > 0 && !field.options.includes(text) && !field.editable) {
          throw new OfficeError('INVALID_REQUEST', `字段「${name}」的选项里没有「${text}」。`, { options: field.options })
        }
        plan.push({ field, name, kind: 'Ch', value: text, index: field.options.indexOf(text) })
      } else {
        throw new OfficeError('UNSUPPORTED_FEATURE', `字段「${name}」的类型 ${field.ft ?? '未知'} 暂不支持填写。`)
      }
    }
    if (!walker.acroRef) {
      throw new OfficeError('UNSUPPORTED_FEATURE', 'AcroForm 是内联字典（没有独立对象号），本工具不改写这种结构。')
    }

    // 第二遍：全部校验通过，开始写
    const filled = []
    for (const item of plan) {
      const { field } = item
      const dict = field.dict
      if (item.kind === 'Tx') {
        dict.V = encodePdfText(item.value)
        filled.push({ name: item.name, field_type: 'Tx', value: item.value })
      } else if (item.kind === 'Btn') {
        dict.V = item.target
        // 值写在字段上，外观状态要写到每个 widget 的 /AS（合并式字段则写在自己身上）。
        // **单选组**不能把所有 widget 都设成开：每个 widget 有自己的开状态名，
        // 只有与目标值同名的那个设成开、其余设 Off，否则整组都会显示选中。
        if (field.hasOwnAppearance) dict.AS = item.target
        for (const kid of field.kids) {
          if (!kid.ref) {
            warnings.push({ code: 'SKIPPED_INLINE_WIDGET', message: `字段「${item.name}」有一个内联 widget，未改写它的 /AS。` })
            continue
          }
          const kidOn = collectOnStates(this, kid.dict, [])
          const as = field.radio ? (kidOn.includes(item.target) ? item.target : 'Off') : item.target
          this.#stage(kid.ref, { ...kid.dict, AS: as })
        }
        filled.push({ name: item.name, field_type: 'Btn', value: item.target, on_states: item.on, radio: field.radio })
      } else {
        dict.V = encodePdfText(item.value)
        if (item.index >= 0) dict.I = [item.index]
        else delete dict.I
        filled.push({ name: item.name, field_type: 'Ch', value: item.value, option_index: item.index >= 0 ? item.index : null })
      }
      this.#stage(field.ref, dict)
    }

    walker.acro.NeedAppearances = true
    this.#stage(walker.acroRef, walker.acro)

    return {
      type: 'fill_form',
      filled_count: filled.length,
      filled,
      need_appearances: true,
      warnings,
      note: '只改 /V 与 /AS 并要求阅读器重绘外观（不生成外观流）。'
    }
  }

  /**
   * 汇总表单字段的内部结构（供 `forms()` 与 `fillForm()` 共用）。
   *
   * 只支持**间接对象**形式的字段与 AcroForm：内联字典没有独立对象号，改写它要连带重写父对象，
   * 那种结构在本工具里会被明确拒绝，而不是猜着改。
   *
   * @returns {{acro: object, acroRef: string, fields: object[]}|null} 表单结构；没有表单时返回 null。
   */
  #collectFormFields() {
    const acroRaw = this.#catalogValue('AcroForm')
    // `/AcroForm 10 0 R` 取出来是 `{ref: '10 0'}`，必须归一成字符串再交给 `#stage`
    // （把对象当 key 塞进 #pending 会在 save() 时才炸，且报错完全指不到现场）
    const acroRef = acroRaw && typeof acroRaw === 'object' && typeof acroRaw.ref === 'string' ? acroRaw.ref : null
    const acro = this.#pendingValueOf(this.deref(acroRaw))
    if (!acro || typeof acro !== 'object' || acro.stream) return null
    const refOf = (value) => (value && typeof value === 'object' && typeof value.ref === 'string' ? value.ref : null)
    const fields = []
    const walk = (value, prefix) => {
      const dict = this.#pendingValueOf(this.deref(value))
      if (!dict || typeof dict !== 'object' || dict.stream) return
      const partial = decodePdfString(this.#pendingValueOf(dict.T))
      const name = partial ? (prefix ? `${prefix}.${partial}` : partial) : prefix
      const ft = this.#pendingValueOf(dict.FT)
      const kidsRaw = this.#pendingValueOf(this.deref(dict.Kids))
      const kids = Array.isArray(kidsRaw) ? kidsRaw : []
      if (kids.length > 0 && !ft) {
        for (const kid of kids) walk(kid, name)
        return
      }
      const flags = Number(this.#pendingValueOf(dict.Ff) ?? 0) || 0
      const kidDicts = kids
        .map((kid) => ({ ref: refOf(kid), dict: this.#pendingValueOf(this.deref(kid)) }))
        .filter((k) => k.dict && typeof k.dict === 'object' && !k.dict.stream)
      const onStates = collectOnStates(this, dict, kidDicts)
      const options = []
      const rawOptions = this.#pendingValueOf(this.deref(dict.Opt))
      if (Array.isArray(rawOptions)) {
        for (const option of rawOptions) {
          const resolved = this.#pendingValueOf(this.deref(option))
          if (Array.isArray(resolved)) options.push(decodePdfString(this.#pendingValueOf(this.deref(resolved[0]))) ?? '')
          else if (resolved !== undefined) options.push(decodePdfString(resolved) ?? '')
        }
      }
      fields.push({
        name,
        ref: refOf(value),
        dict,
        ft: typeof ft === 'string' ? ft : null,
        flags,
        readOnly: (flags & 1) !== 0,
        required: (flags & 2) !== 0,
        editable: (flags & 262144) !== 0,
        multiline: (flags & 4096) !== 0,
        password: (flags & 8192) !== 0,
        combo: (flags & 131072) !== 0,
        radio: (flags & 32768) !== 0,
        kids: kidDicts,
        onStates,
        options,
        hasOwnAppearance: Boolean(this.deref(dict.AP))
      })
    }
    const rootFields = this.#pendingValueOf(this.deref(acro.Fields))
    if (Array.isArray(rootFields)) for (const field of rootFields) walk(field, '')
    return { acro, acroRef, fields }
  }

  /**
   * 把内部字段结构投影成对外可读的形状。
   * @param {object} field - `#collectFormFields()` 的字段项。
   * @returns {object} 对外字段信息。
   */
  #describeField(field) {
    const value = this.#pendingValueOf(field.dict.V)
    let current = null
    if (field.ft === 'Btn') {
      if (typeof value === 'boolean') current = value ? (field.onStates[0] ?? 'On') : 'Off'
      else if (typeof value === 'string') current = value
      else current = 'Off'
    } else {
      current = decodePdfString(value)
    }
    return {
      name: field.name,
      field_type: field.ft,
      kind: fieldKindLabel(field),
      value: current,
      options: field.options.length > 0 ? field.options : null,
      on_states: field.ft === 'Btn' ? field.onStates : null,
      // 每个 widget 的外观状态：单选组靠它区分「哪一个被选中」
      widget_states:
        field.ft === 'Btn'
          ? [
              ...(field.hasOwnAppearance ? [{ on_states: field.onStates, appearance_state: this.#pendingValueOf(field.dict.AS) ?? 'Off' }] : []),
              ...field.kids.map((kid) => {
                const own = collectOnStates(this, kid.dict, [])
                return { on_states: own, appearance_state: this.#pendingValueOf(kid.dict.AS) ?? 'Off' }
              })
            ]
          : null,
      read_only: field.readOnly,
      required: field.required,
      checked: field.ft === 'Btn' && !field.radio ? current !== 'Off' && current !== null : null
    }
  }

  /**
   * 取某个值的「待写回版本优先」的值。
   *
   * 刚写入的批注对象在 `#pending` 里，对象表里还是旧值 —— 与 `rewrite()` 的处理保持一致，
   * 否则「写完立刻读回」会读到空。
   *
   * @param {unknown} value - 解析出来的值。
   * @returns {unknown} 最新值。
   */
  #pendingValueOf(value) {
    if (value && typeof value === 'object' && typeof value.ref === 'string') {
      return this.#pending.get(value.ref) ?? value
    }
    return value
  }

  /**
   * 新建一个对象并登记待写回。
   * @param {unknown} value - 对象值。
   * @returns {string} `"num gen"` 引用。
   */
  #addObject(value) {
    const key = this.#nextNum()
    this.#objects.set(key, { num: Number(key.split(' ')[0]), gen: 0, value })
    this.#stage(key, value)
    return key
  }

  /**
   * 新建一个内容流对象。
   * @param {string} commands - 内容流指令（ASCII）。
   * @returns {string} 引用。
   */
  #addStream(commands) {
    const stream = Buffer.from(commands, 'latin1')
    const wrapper = { dict: { Length: stream.length }, stream }
    const key = this.#nextNum()
    this.#objects.set(key, { num: Number(key.split(' ')[0]), gen: 0, value: wrapper })
    this.#stage(key, wrapper)
    return key
  }

  /**
   * 分配一个新的对象号（取已用最大号 +1，包括本次待写回的对象）。
   * @returns {string} `"num 0"`。
   */
  #nextNum() {
    let maxNum = 0
    for (const entry of this.#objects.values()) if (entry.num > maxNum) maxNum = entry.num
    for (const key of this.#pending.keys()) {
      const num = Number(key.split(' ')[0])
      if (num > maxNum) maxNum = num
    }
    return `${maxNum + 1} 0`
  }

  /**
   * 为一个页面准备「带上叠加字体」的资源字典。
   *
   * 两个坑必须同时避开：
   *   1. 不能只写一个含 Wm 的页面级 `/Resources`：页面级资源会**整体覆盖**从页面树
   *      继承来的资源，原有字体全丢。所以要复制「有效资源」（页面自身或继承来的）。
   *   2. 更不能把 `/Font` 换成只含 Wm 的字典：那会把页面原有的 F1/F2… 全部丢掉 ——
   *      页面还能打开、还能解析，但原有文字会失去字体（中文直接退化成乱码/控制字符）。
   *      必须在原有 `/Font` 条目**之上**加一个 Wm。
   *
   * 还有一个**只有第三方阅读器才暴露得出来**的坑：字体字典与外层资源字典都必须**内联**。
   * 写成 `/Resources → 对象 → /Font → 对象 → /Wm → 对象` 这种多级间接引用在规范里合法、
   * 我们自己的解析器也读得出来，但 Word 的 PDF 解析器会**静默丢弃**这些文字
   * （用最小样本逐项对照确认：同一绘制指令，内联能读到、多级间接读不到）。
   *
   * @param {object} page - 页面条目。
   * @param {string} fontRef - Helvetica 字体对象引用。
   * @returns {string} 新资源字典的引用。
   */
  #withFontResource(page, fontRef) {
    const effective = this.deref(this.#inherited(page.value, 'Resources'))
    const copy = {}
    if (effective && typeof effective === 'object') {
      for (const [key, value] of Object.entries(effective)) {
        if (key === 'Font' || key.startsWith('__')) continue
        copy[key] = value
      }
    }
    // 合并原有字体条目（可能是直接字典，也可能是引用），并把 Wm 内联在同一层
    const existingFonts = this.deref(effective?.Font)
    const fonts = {}
    if (existingFonts && typeof existingFonts === 'object') {
      for (const [name, ref] of Object.entries(existingFonts)) {
        if (name.startsWith('__')) continue
        fonts[name] = ref
      }
    }
    if (fonts.Wm === undefined) fonts.Wm = { ref: fontRef }
    copy.Font = fonts
    return this.#addObject(copy)
  }

  /**
   * 把新内容流追加到页面内容之后。
   * @param {unknown} contents - 原 `/Contents` 值。
   * @param {string} streamRef - 新流引用。
   * @returns {object[]} 新的内容流数组。
   */
  #appendContents(contents, streamRef) {
    if (contents === undefined || contents === null) return [{ ref: streamRef }]
    const list = Array.isArray(contents) ? [...contents] : [contents]
    list.push({ ref: streamRef })
    return list
  }

  /**
   * 读一个可能被页面树继承的值。
   * @param {object} pageValue - 页面字典。
   * @param {string} key - 键名。
   * @returns {unknown} 值。
   */
  #inherited(pageValue, key) {
    let node = pageValue
    for (let depth = 0; depth < 32 && node && typeof node === 'object'; depth += 1) {
      const direct = node[key]
      if (direct !== undefined) return direct
      const parent = this.deref(node.Parent)
      node = parent && typeof parent === 'object' ? parent : null
    }
    return undefined
  }

  /**
   * 页面尺寸（含继承的 MediaBox），缺省 A4。
   * @param {object} page - 页面条目。
   * @returns {number[]} `[x0, y0, x1, y1]`。
   */
  #pageBox(page) {
    const box = this.deref(this.#inherited(page.value, 'MediaBox'))
    if (Array.isArray(box) && box.length === 4 && box.every((n) => Number.isFinite(Number(n)))) {
      return box.map(Number)
    }
    return [0, 0, 595.28, 841.89]
  }

  /**
   * 写入文档元数据（`/Info` 字典）。
   *
   * 只改 `/Info` 对象本身，其余内容一律不动。原本没有 `/Info` 时会新建一个，
   * 并在新 trailer 里补上 `/Info` 引用 —— 这一步容易被漏掉：对象写进了文件，
   * 但 trailer 不引用它，任何阅读器都看不到。
   *
   * 文本编码：纯 ASCII 用字面串（PDFDocEncoding），含非 ASCII（如中文）用
   * **UTF-16BE + BOM 的十六进制串** —— 这是 PDF 里表示非 ASCII 文本的标准做法。
   *
   * @param {object} args - 参数。
   * @param {string} [args.title] - 标题。
   * @param {string} [args.author] - 作者。
   * @param {string} [args.subject] - 主题。
   * @param {string} [args.keywords] - 关键词。
   * @param {string} [args.creator] - 创建程序。
   * @param {string} [args.producer] - 生成程序。
   * @returns {object} 变更信息。
   */
  updateMetadata({ title, author, subject, keywords, creator, producer } = {}) {
    const provided = { title, author, subject, keywords, creator, producer }
    const names = {
      title: 'Title',
      author: 'Author',
      subject: 'Subject',
      keywords: 'Keywords',
      creator: 'Creator',
      producer: 'Producer'
    }
    const infoRef = this.#trailerValue('Info')
    const hasInfo = infoRef && typeof infoRef === 'object' && typeof infoRef.ref === 'string'
    const info = hasInfo ? { ...(this.deref(infoRef) ?? {}) } : {}
    for (const key of Object.keys(info)) if (key.startsWith('__')) delete info[key]

    const fields = []
    for (const [key, name] of Object.entries(names)) {
      const value = provided[key]
      if (value === undefined || value === null) continue
      fields.push({ field: name, from: decodePdfString(info[name]) ?? null, to: String(value) })
      info[name] = encodePdfText(String(value))
    }
    if (fields.length === 0) {
      throw new OfficeError('INVALID_REQUEST', '至少需要给出一个要写入的元数据字段（title/author/subject/keywords/creator/producer）。')
    }
    if (hasInfo) {
      this.#stage(infoRef.ref, info)
    } else {
      this.#infoOverride = this.#addObject(info)
    }
    return { type: 'update_metadata', fields, created_info: !hasInfo }
  }

  /** 原本没有 /Info 时新建对象的引用，写 trailer 时要用。 */
  #infoOverride = null

  /**
   * 把文档重写成一份**全新的紧凑 PDF**（对象图重写 / 字节回收）。
   *
   * 与增量更新相反：这里只写出从 trailer 出发**可达**的对象，重新编号、重新建 xref，
   * 于是「被删掉的页面对象、旧版本的页面、对象流容器、旧交叉引用段」全部消失。
   * 三个用途：
   *   1. **回收字节** —— 反复增量更新的文件会一直变大，重写后回到基线；
   *   2. **拆分/导出** —— `keepPages` 只保留指定页面，未选页面的内容**不再留在文件字节里**
   *      （增量式删页做不到这一点：对象还在文件里，只是没人引用）；
   *   3. 消除对象流与交叉引用流的复杂度，产出一份最朴素的 PDF。
   *
   * 保真取舍（明确写出来）：
   *   - 内容流**保持原压缩字节**（`/Filter` 与数据一起搬），不做解压再压缩；
   *   - 对象流（ObjStm）里的成员对象会被**展平**成普通顶层对象：语义等价，布局不同；
   *   - 交叉引用流被换成经典 xref 表：任何阅读器都读得懂；
   *   - 有引用指向**不可达或不可解析**的对象时直接拒绝重写，而不是产出一份悄悄缺内容的文件。
   *
   * @param {object} [args] - 参数。
   * @param {number[]|null} [args.keepPages] - 只保留这些页（0 基，按给定顺序）；省略表示全部页。
   * @returns {object} `{buffer, objects, bytes, pages}`。
   */
  rewrite({ keepPages = null } = {}) {
    const structure = this.structure()
    if (structure.encrypted) {
      throw new OfficeError('PASSWORD_REQUIRED', 'PDF 已加密，无法在不知口令的情况下重写。', { needsConfirmation: true })
    }
    const pages = this.pages()
    let keep = pages.map((_, index) => index)
    if (keepPages !== null) {
      if (!Array.isArray(keepPages) || keepPages.length === 0) {
        throw new OfficeError('INVALID_REQUEST', 'keepPages 必须是非空下标数组，或省略表示全部页。')
      }
      for (const index of keepPages) this.#requirePage(index)
      keep = [...keepPages]
    }

    const graph = this.#collectGraph({ keep })
    const file = serializePdfFile({
      objects: graph.objects,
      rootNum: graph.catalogNum,
      infoNum: graph.infoNum,
      inlineInfo: graph.inlineInfo,
      version: this.info.version
    })
    return {
      buffer: file.buffer,
      objects: graph.objects.length,
      bytes: file.bytes,
      pages: keep,
      dropped: pages.length - keep.length,
      removed_page_refs: graph.removedRefs
    }
  }

  /**
   * 收集「从 trailer 出发可达」的对象图，并重编号。
   *
   * 重写与合并共用这一步：两者都是「按可达性挑对象 + 重新编号」，
   * 区别只在最后怎么组装（重写直接写文件；合并还要新建目录与页面树）。
   *
   * @param {object} args - 参数。
   * @param {number[]} args.keep - 保留的页下标（顺序即页序）。
   * @param {number} [args.startNum] - 起始对象号；合并多份文档时用来错开编号。
   * @returns {object} 图：`{objects, remap, catalogNum, pagesRootNum, infoNum, inlineInfo, removedRefs, nextNum}`。
   */
  #collectGraph({ keep, startNum = 1 }) {
    const pages = this.pages()
    const trailer = this.#newestTrailer() ?? {}
    const roots = []
    // /Info 既可能是间接引用，也可能**直接内联**在 trailer 里（交叉引用流的字典就常这样写）。
    const inlineInfo =
      trailer.Info && typeof trailer.Info === 'object' && typeof trailer.Info.ref !== 'string' ? trailer.Info : null
    const infoEntry = trailer.Info && typeof trailer.Info.ref === 'string' ? this.#entryByRef(trailer.Info.ref) : null
    const infoKey = infoEntry ? `${infoEntry.num} ${infoEntry.gen}` : null
    if (trailer.Root && typeof trailer.Root.ref === 'string') roots.push(trailer.Root.ref)
    if (infoKey) roots.push(infoKey)
    if (roots.length === 0) {
      throw new OfficeError('CORRUPTED_DOCUMENT', 'trailer 里没有 /Root，无法重写。')
    }

    const order = [] // {key, entry, value}，value 已按目标页序过滤并排序 /Kids
    const seen = new Set()
    const queue = [...roots]
    const missing = []
    // 页面引用归一化：文件里写的是 `"N G"`，对象表里的键也以它为准；
    // 但生成号可能对不上（悬空 gen 会退回 0），所以用对象表归一化后再比对。
    const rankOf = new Map()
    keep.forEach((index, position) => rankOf.set(`${pages[index].num} ${pages[index].gen}`, position))
    const rank = (ref) => {
      const entry = this.#entryByRef(ref)
      if (!entry) return -1
      const type = this.deref(entry.value?.Type)
      if (type === 'Pages') return Number.MAX_SAFE_INTEGER // 子页面树整棵保留，其 /Kids 会另行过滤排序
      const found = rankOf.get(`${entry.num} ${entry.gen}`)
      return found === undefined ? -1 : found
    }
    const sanitizeStats = { removed: 0 }
    const isDroppedPage = (item) => {
      if (!item || typeof item !== 'object' || typeof item.ref !== 'string') return false
      const entry = this.#entryByRef(item.ref)
      if (!entry) return false
      if (this.deref(entry.value?.Type) !== 'Page') return false
      return !rankOf.has(`${entry.num} ${entry.gen}`)
    }
    while (queue.length > 0) {
      const ref = queue.shift()
      const entry = this.#entryByRef(ref)
      if (!entry) {
        missing.push(ref)
        continue
      }
      const key = `${entry.num} ${entry.gen}`
      if (seen.has(key)) continue
      seen.add(key)
      // 待写回的版本优先：`updateMetadata` 这类操作是把新字典放进 #pending，
      // 对象表里仍是旧值。忽略它会让「先改元数据再重写」静默丢掉修改。
      const current = this.#pending.get(key) ?? entry.value
      const value = sanitizePageRefs(filterPageNode(current, rank), isDroppedPage, sanitizeStats)
      order.push({ key, entry, value })
      queue.push(...collectRefs(value))
    }
    if (missing.length > 0) {
      throw new OfficeError('CORRUPTED_DOCUMENT', `有 ${missing.length} 个被引用的对象解析不出来，拒绝重写出不完整的文件。`, {
        missing: missing.slice(0, 10)
      })
    }

    const remap = new Map()
    order.forEach((item, index) => remap.set(item.key, startNum + index))
    // 对象值里的引用**在这里一次性改成新编号**：这样序列化阶段不必再关心映射，
    // 新建的合成对象（合并时的目录与页面树）与新编号也就天然一致。
    const objects = order.map((item) => ({ num: remap.get(item.key), value: relabelRefs(item.value, remap) }))
    const catalogKey = this.#rootKeyOf(order)
    return {
      objects,
      remap,
      catalogNum: remap.get(catalogKey),
      pagesRootNum: remap.get(this.#pagesRootKey(order)),
      infoNum: infoKey && remap.has(infoKey) ? remap.get(infoKey) : null,
      inlineInfo,
      removedRefs: sanitizeStats.removed,
      nextNum: startNum + order.length
    }
  }

  /**
   * 找可达对象里页面树根（`/Type /Pages` 且没有 `/Parent`）的键。
   * @param {object[]} order - 可达对象列表。
   * @returns {string} 对象键。
   */
  #pagesRootKey(order) {
    for (const item of order) {
      const value = item.value
      if (!value || typeof value !== 'object') continue
      if (this.deref(value.Type) === 'Pages' && this.deref(value.Parent) === undefined) return item.key
    }
    throw new OfficeError('CORRUPTED_DOCUMENT', '可达对象里找不到页面树根（/Type /Pages）。')
  }

  /**
   * 合并多份 PDF：把所有文档的可达对象按序重编号后放进一个文件，并新建目录与页面树。
   *
   * 合并的本质是**跨文件搬运对象图**：两份文档的同号对象毫无关系，必须整体错开编号，
   * 所有内部引用跟着改写（这一步复用重写的可达性收集）。页面树用嵌套写法：
   * 新建一个 `/Pages` 根，把每份文档原来的 `/Pages` 根作为子节点挂上去 ——
   * 这样每份文档自己的页面树与 `/Parent` 关系都保持原样。
   *
   * **明确不做的事**：只搬运页面内容与资源，**不带**各份文档的目录（`/Outlines`）、
   * 命名目标、表单（`/AcroForm`）、结构树与页面标签 —— 它们跨文档合并需要重映射，
   * 静默带进来只会得到互相打架的引用。合并后只保留第一份文档的 `/Info` 元数据。
   *
   * @param {Buffer[]} buffers - 各份 PDF 字节（顺序即页序）。
   * @param {object} [options] - 选项。
   * @param {string} [options.version] - 输出 PDF 版本；默认用第一份的版本。
   * @returns {object} `{buffer, objects, bytes, pages, sources}`。
   */
  static merge(buffers, { version = null } = {}) {
    if (!Array.isArray(buffers) || buffers.length < 2) {
      throw new OfficeError('INVALID_REQUEST', '合并至少需要两份 PDF。')
    }
    const docs = buffers.map((buffer) => PdfDocument.open(buffer))
    docs.forEach((doc, index) => {
      if (doc.structure().encrypted) {
        throw new OfficeError('PASSWORD_REQUIRED', `第 ${index + 1} 份 PDF 已加密，无法合并。`, { needsConfirmation: true })
      }
    })

    const objects = []
    const childRootNums = []
    let next = 1
    let pageCount = 0
    let infoNum = null
    let inlineInfo = null
    let removedRefs = 0
    docs.forEach((doc, index) => {
      const keep = doc.pages().map((_, i) => i)
      const graph = doc.#collectGraph({ keep, startNum: next })
      // 丢掉各自的目录对象：合并后只留我们新建的那一个
      for (const item of graph.objects) {
        if (item.num === graph.catalogNum) continue
        objects.push(item)
      }
      childRootNums.push(graph.pagesRootNum)
      pageCount += keep.length
      removedRefs += graph.removedRefs
      next = graph.nextNum
      if (index === 0) {
        infoNum = graph.infoNum
        inlineInfo = graph.inlineInfo
      }
    })

    const newRootNum = next
    const newCatalogNum = next + 1
    // 各份文档原来的页面树根挂到新根下：补上 /Parent，保持树的双向关系完整
    const childRoots = new Set(childRootNums)
    for (const item of objects) {
      if (childRoots.has(item.num)) item.value = { ...item.value, Parent: { ref: `${newRootNum} 0` } }
    }
    objects.push({
      num: newRootNum,
      value: { Type: 'Pages', Kids: childRootNums.map((num) => ({ ref: `${num} 0` })), Count: pageCount }
    })
    objects.push({ num: newCatalogNum, value: { Type: 'Catalog', Pages: { ref: `${newRootNum} 0` } } })
    objects.sort((a, b) => a.num - b.num)

    const file = serializePdfFile({
      objects,
      rootNum: newCatalogNum,
      infoNum,
      inlineInfo,
      version: version ?? docs[0].info.version
    })
    return {
      buffer: file.buffer,
      objects: objects.length,
      bytes: file.bytes,
      pages: pageCount,
      sources: docs.length,
      removed_page_refs: removedRefs
    }
  }

  /**
   * 找可达对象里目录（Catalog）的键，作为新 trailer 的 /Root。
   * @param {object[]} order - 可达对象列表。
   * @returns {string} 对象键。
   */
  #rootKeyOf(order) {
    for (const item of order) {
      const value = item.value
      if (value && typeof value === 'object' && this.deref(value.Type) === 'Catalog') return item.key
    }
    throw new OfficeError('CORRUPTED_DOCUMENT', '可达对象里找不到 /Type /Catalog，无法重写。')
  }
  /**
   * 提取页面上的图片。
   *
   * 只做「把 PDF 里已经存在的图像流解出来」这一件事，不做光栅化（不渲染页面）：
   *   - `/DCTDecode`（JPEG）与 `/JPXDecode`（JPEG 2000）**原样导出**，不重新编码；
   *   - `/FlateDecode` + 8 位分量 → 还原预测器后包成 PNG；
   *   - 索引色会展开成 RGB，`/SMask` 会合成成带透明通道的 PNG；
   *   - 1/2/4 位分量、CMYK、CCITT/JBIG2 等当前不支持的形态**明确跳过并给出原因**，
   *     而不是导出打不开的文件。
   *
   * 同一张图片被多页共用时只导出一次，并在 `pages` 里列出用到它的页码。
   *
   * @param {object} [args] - 参数。
   * @param {number[]|null} [args.pages] - 只提取这些页（0 基）；省略表示全部页。
   * @returns {object} `{images, skipped}`：可导出的图片与跳过的对象（含原因）。
   */
  extractImages({ pages = null } = {}) {
    const all = this.pages()
    const targets = pages === null ? all.map((_, index) => index) : pages
    if (!Array.isArray(targets) || targets.length === 0) {
      throw new OfficeError('INVALID_REQUEST', 'pages 必须是非空下标数组，或省略表示全部页。')
    }
    for (const index of targets) this.#requirePage(index)

    // 私有字段只能在类内访问，所以把「按引用取对象（并让待写回版本优先）」作为回调传下去
    const resolveRef = (ref) => {
      if (ref && typeof ref === 'object' && typeof ref.ref === 'string') {
        const entry = this.#entryByRef(ref.ref)
        return entry ? (this.#pending.get(`${entry.num} ${entry.gen}`) ?? entry.value) : undefined
      }
      return this.deref(ref)
    }

    const found = new Map()
    for (const index of targets) {
      const page = this.#requirePage(index)
      const resources = this.deref(this.#inherited(page.value, 'Resources'))
      const xobjects = this.deref(resources?.XObject)
      if (!xobjects || typeof xobjects !== 'object') continue
      for (const [name, ref] of Object.entries(xobjects)) {
        if (name.startsWith('__')) continue
        const entry = ref && typeof ref === 'object' && typeof ref.ref === 'string' ? this.#entryByRef(ref.ref) : null
        const value = resolveRef(ref)
        if (this.deref(dictOf(value).Subtype) !== 'Image') continue
        const key = entry ? `${entry.num} ${entry.gen}` : name
        let item = found.get(key)
        if (!item) {
          item = decodeImageXObject(this, value, key, name, resolveRef)
          found.set(key, item)
        }
        if (!item.pages.includes(index + 1)) item.pages.push(index + 1)
      }
    }
    const items = [...found.values()].map((item) => ({ ...item, pages: [...item.pages].sort((a, b) => a - b) }))
    return {
      images: items.filter((item) => !item.skipped),
      skipped: items.filter((item) => item.skipped)
    }
  }

  /**
   * 取一个页面上的图片对象清单（供工具层做命名与页内序号）。
   * @param {number[]|null} [pages] - 目标页。
   * @returns {object[]} `{objectKey, name, pages, width, height}`。
   */
  listImages({ pages = null } = {}) {
    const { images, skipped } = this.extractImages({ pages })
    return [...images, ...skipped].map((item) => ({
      objectKey: item.objectKey,
      name: item.name,
      pages: item.pages,
      width: item.width,
      height: item.height,
      skipped: item.skipped ?? null
    }))
  }

  /**
   * 删除一页。
   *
   * 只从页面树的 `/Kids` 摘掉引用；页面对象本身仍留在文件里
   * （增量更新不回收字节）。这一点在返回值里说明，避免调用方误以为文件变小了。
   *
   * @param {object} args - 参数。
   * @param {number} args.page - 页面下标（0 基）。
   * @returns {object} 变更信息。
   */
  deletePage({ page }) {
    const target = this.#requirePage(page)
    if (this.pages().length <= 1) {
      throw new OfficeError('INVALID_REQUEST', 'PDF 只剩一页，删除后将不再是有效文档。')
    }
    const key = `${target.num} ${target.gen}`
    const order = this.pages()
      .map((p) => `${p.num} ${p.gen}`)
      .filter((k) => k !== key)
    this.#rewriteKids(order)
    return {
      type: 'delete_page',
      page,
      removed_object: key,
      page_count: order.length,
      note: '页面对象仍保留在文件中（增量更新不回收字节），但已不被页面树引用。'
    }
  }

  /**
   * 取一个页面，越界时给出总数。
   * @param {number} page - 页面下标。
   * @returns {object} 页面条目。
   */
  #requirePage(page) {
    const target = this.pages()[page]
    if (!target) {
      throw new OfficeError('INVALID_REQUEST', `页面下标 ${page} 超出范围（共 ${this.pages().length} 页）。`, {
        page_count: this.pages().length
      })
    }
    return target
  }

  /**
   * 重写页面树的 `/Kids` 与 `/Count`。
   *
   * 走页面树而不是改页面对象上的 `/Parent`：渲染顺序与页面存在性由 `/Kids` 决定，
   * 只改页面对象不会改变任何可见结果。
   *
   * @param {string[]} order - 页面对象键（`"num gen"`）的新顺序。
   * @returns {void}
   */
  #rewriteKids(order) {
    const root = this.#pagesRoot()
    if (!root) throw new OfficeError('CORRUPTED_DOCUMENT', '找不到页面树根（/Type /Pages），无法调整页面。')
    const kids = this.deref(root.value.Kids)
    if (!Array.isArray(kids) || kids.length === 0) {
      throw new OfficeError('CORRUPTED_DOCUMENT', '页面树的 /Kids 为空，无法调整页面。')
    }
    const directPages = kids.filter((k) => {
      const resolved = this.deref(k)
      return resolved && typeof resolved === 'object' && this.deref(resolved.Type) === 'Page'
    })
    if (directPages.length !== kids.length) {
      throw new OfficeError('UNSUPPORTED_FEATURE', '页面树含嵌套节点，当前只支持扁平页面树的重排与删除。')
    }
    root.value.Kids = order.map((key) => ({ ref: key }))
    root.value.Count = order.length
    this.#stage(`${root.num} ${root.gen}`, root.value)
    this.#pages = null // 顺序变了，缓存作废
  }

  /**
   * 找页面树根。
   * @returns {object|undefined} 页面树根条目。
   */
  #pagesRoot() {
    for (const entry of this.#objects.values()) {
      const value = entry.value
      if (!value || typeof value !== 'object' || value.stream !== undefined) continue
      if (this.deref(value.Type) === 'Pages' && this.deref(value.Parent) === undefined && value.Kids !== undefined) {
        return entry
      }
    }
    return undefined
  }

  /**
   * 登记一个被修改的对象，等待写回。
   * @param {string} key - `"num gen"`。
   * @param {unknown} value - 新的对象值。
   * @returns {void}
   */
  #stage(key, value) {
    // key 必须是 `"num gen"` 字符串：塞进别的类型不会立刻报错，而是在 save() 里以
    // `k.split is not a function` 这种指不到现场的形态炸掉（踩过一次，这里加护栏）
    if (typeof key !== 'string') {
      throw new OfficeError('INTERNAL_ERROR', `#stage 的 key 必须是 "num gen" 字符串，实际是 ${typeof key}。`)
    }
    this.#pending.set(key, value)
  }

  #pending = new Map()

  /** @returns {boolean} 是否有待写回的修改。 */
  get dirty() {
    return this.#pending.size > 0
  }

  /**
   * 以**增量更新**方式写回修改。
   *
   * 原始字节一个都不动，只在文件尾部追加被修改对象的新版本、新的交叉引用段，
   * 以及带 `/Prev` 指回旧 xref 的 trailer。这与 OOXML 侧的「最小修改」是同一思路：
   * 未被触碰的内容不可能被改坏。
   *
   * @returns {Buffer} 新的 PDF 字节。
   */
  save() {
    if (this.#pending.size === 0) return Buffer.from(this.#buffer)

    const prevStartxref = readStartxref(this.#buffer)
    const maxNum = Math.max(
      ...[...this.#objects.values()].map((e) => e.num),
      ...[...this.#pending.keys()].map((k) => Number(k.split(' ')[0]))
    )
    const size = maxNum + 1

    const chunks = []
    let offset = this.#buffer.length
    const entries = []
    for (const [key, value] of this.#pending) {
      const [num, gen] = key.split(' ').map(Number)
      const text = `${num} ${gen} obj\n${serializePdfObject(value)}\nendobj\n`
      entries.push({ num, gen, offset })
      chunks.push(Buffer.from(text, 'latin1'))
      offset += Buffer.byteLength(text, 'latin1')
    }
    entries.sort((a, b) => a.num - b.num)

    let xref = 'xref\n0 1\n0000000000 65535 f \n'
    let i = 0
    while (i < entries.length) {
      let j = i
      while (j + 1 < entries.length && entries[j + 1].num === entries[j].num + 1) j += 1
      xref += `${entries[i].num} ${j - i + 1}\n`
      for (let k = i; k <= j; k += 1) {
        xref += `${String(entries[k].offset).padStart(10, '0')} ${String(entries[k].gen).padStart(5, '0')} n \n`
      }
      i = j + 1
    }

    const trailerParts = [`/Size ${size}`]
    const rootRef = this.#rootRef()
    if (rootRef) trailerParts.push(`/Root ${rootRef} R`)
    const infoRef = this.#infoOverride ?? this.#trailerValue('Info')
    if (typeof infoRef === 'string') trailerParts.push(`/Info ${infoRef} R`)
    else if (infoRef && typeof infoRef === 'object' && typeof infoRef.ref === 'string') trailerParts.push(`/Info ${infoRef.ref} R`)
    if (Number.isFinite(prevStartxref)) trailerParts.push(`/Prev ${prevStartxref}`)

    const tail = Buffer.from(`${xref}trailer\n<< ${trailerParts.join(' ')} >>\nstartxref\n`, 'latin1')
    // startxref 的值必须是 **xref 关键字本身的偏移**（不是值自身的偏移，也不是 trailer 的偏移）。
    // 差一个字节，第三方阅读器就整份文件都读不出来。
    const startxrefValue = this.#buffer.length + chunks.reduce((n, c) => n + c.length, 0)
    const end = Buffer.from(`\n%%EOF\n`, 'latin1')
    const result = Buffer.concat([this.#buffer, ...chunks, tail, Buffer.from(String(startxrefValue), 'latin1'), end])
    this.#pending.clear()
    return result
  }

  /**
   * 取文档目录的引用字符串。
   * @returns {string|null} `"num gen"` 或 null。
   */
  #rootRef() {
    const trailer = this.#newestTrailer()
    const root = trailer ? trailer.Root : undefined
    if (root && typeof root === 'object' && typeof root.ref === 'string') return root.ref
    for (const entry of this.#objects.values()) {
      const value = entry.value
      if (value && typeof value === 'object' && this.deref(value.Type) === 'Catalog') return `${entry.num} ${entry.gen}`
    }
    return null
  }
}

/**
 * 从文件尾部读出最后一个 `startxref` 的偏移。
 *
 * 增量更新要写 `/Prev` 指回旧交叉引用段，否则其他阅读器会把旧对象当成空闲对象。
 * 只扫描尾部 2 KB：`startxref` 按规范必然出现在离文件尾很近的位置。
 *
 * @param {Buffer} buf - 整个文件的字节。
 * @returns {number|null} 偏移；读不到时返回 null。
 */
function readStartxref(buf) {
  const window = buf.subarray(Math.max(0, buf.length - 2048)).toString('latin1')
  let last = null
  for (const match of window.matchAll(/startxref\s+(\d+)/g)) last = match[1]
  if (last === null) return null
  const value = Number(last)
  return Number.isSafeInteger(value) && value >= 0 ? value : null
}

/**
 * 序列化一个 PDF 对象。
 *
 * 只支持解析器产出的值类型，正好覆盖页面对象与页面树的全部内容：
 * 数字、名称、字符串、`{ref}` 间接引用、数组、字典、布尔与 null。
 * 引用按 `num gen R` 三段式写出 —— 缺生成号会被阅读器当成数字（目录/元数据静默丢失）。
 *
 * @param {unknown} value - 值。
 * @returns {string} PDF 语法文本。
 * @throws {OfficeError} 遇到无法表达的值时抛出。
 */
function serializePdfObject(value) {
  if (value === null || value === undefined) return 'null'
  if (typeof value === 'boolean') return value ? 'true' : 'false'
  if (typeof value === 'number') return formatPdfNumber(value)
  // 名称在解析后与「未知关键字」同为字符串，此处按名称写回：
  // 页面对象与页面树里的字符串值只可能是名称（Type/Contents/MediaBox 等）。
  if (typeof value === 'string') return `/${serializePdfName(value)}`
  if (Buffer.isBuffer(value)) return serializePdfBuffer(value)
  if (Array.isArray(value)) return `[${value.map((item) => serializePdfObject(item)).join(' ')}]`
  if (typeof value === 'object') {
    // 流对象：解析后的形态是 `{dict, stream}`，新造的流也是这个形态。
    if (value.stream !== undefined && value.dict && typeof value.dict === 'object') {
      if (!Buffer.isBuffer(value.stream)) {
        throw new OfficeError('UNSUPPORTED_FEATURE', '只能写回 Buffer 形式的流字节。')
      }
      const dict = { ...value.dict, Length: value.stream.length }
      return `${serializePdfObject(dict)}\nstream\n${value.stream.toString('latin1')}\nendstream`
    }
    const keys = Object.keys(value)
    if (keys.length === 1 && keys[0] === 'ref' && typeof value.ref === 'string') {
      const [num, gen] = value.ref.split(' ').map(Number)
      if (!Number.isSafeInteger(num) || !Number.isSafeInteger(gen)) {
        throw new OfficeError('CORRUPTED_DOCUMENT', `无法写回非法引用：${value.ref}`)
      }
      return `${num} ${gen} R`
    }
    if (value.__streamStart !== undefined) {
      // 带流的对象不能在增量更新里重写：流字节不在内存模型里，写出去会变成空流。
      throw new OfficeError('UNSUPPORTED_FEATURE', '当前不支持重写带内容流的 PDF 对象。')
    }
    const parts = []
    for (const [key, item] of Object.entries(value)) {
      if (key.startsWith('__')) continue // 解析器内部标记，不写回
      parts.push(`/${serializePdfName(key)} ${serializePdfObject(item)}`)
    }
    return `<< ${parts.join(' ')} >>`
  }
  throw new OfficeError('UNSUPPORTED_FEATURE', `无法序列化的 PDF 值：${typeof value}`)
}

/**
 * 解出一个图像 XObject：返回可写盘的格式与字节，或给出跳过原因。
 *
 * @param {object} doc - 文档（用于 deref / decodeStream）。
 * @param {unknown} wrapper - 图像对象（`{dict, stream}`）。
 * @param {string} key - 对象键。
 * @param {string} name - 资源名。
 * @param {(ref: unknown) => unknown} resolveRef - 按引用取对象（待写回版本优先）。
 * @returns {object} 结果项。
 */
function decodeImageXObject(doc, wrapper, key, name, resolveRef) {
  const dict = dictOf(wrapper)
  const width = Number(doc.deref(dict.Width))
  const height = Number(doc.deref(dict.Height))
  const bits = Number(doc.deref(dict.BitsPerComponent) ?? 8)
  const base = { objectKey: key, name, width, height, bits_per_component: bits, pages: [] }
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    return { ...base, skipped: '缺少有效的 /Width 或 /Height' }
  }
  if (doc.deref(dict.ImageMask) === true) {
    return { ...base, skipped: '图像蒙版（/ImageMask true）本身不是可查看的图片' }
  }

  const filters = normalizeFilters(doc.deref(dict.Filter))
  if (filters.includes('JPXDecode')) {
    return { ...base, format: 'jp2', bytes: wrapper.stream, extension: 'jp2', note: 'JPEG 2000 原样导出（未解码像素）' }
  }
  if (filters.length === 1 && filters[0] === 'DCTDecode') {
    return { ...base, format: 'jpeg', bytes: wrapper.stream, extension: 'jpg', color_space: colorSpaceName(doc, dict.ColorSpace) }
  }
  const decodable = new Set(['FlateDecode', 'Fl', 'ASCIIHexDecode', 'AHx', 'ASCII85Decode', 'A85'])
  const unsupported = filters.filter((filter) => !decodable.has(filter))
  if (unsupported.length > 0) {
    return { ...base, skipped: `不支持的过滤器：${unsupported.join('+')}` }
  }
  if (bits !== 8) {
    return { ...base, skipped: `当前只支持 8 位分量，实际 ${bits} 位` }
  }

  const space = resolveColorSpace(doc, dict.ColorSpace)
  if (!space) {
    return { ...base, skipped: `不支持的色彩空间：${JSON.stringify(colorSpaceName(doc, dict.ColorSpace))}` }
  }

  let data
  try {
    data = doc.decodeStream(wrapper)
  } catch (err) {
    return { ...base, skipped: `解码失败：${err.message}` }
  }
  const params = firstDecodeParms(doc, dict.DecodeParms)
  const predictor = Number(doc.deref(params?.Predictor) ?? 1)
  if (predictor > 1) {
    try {
      data = undoPredictor(data, {
        predictor,
        colors: Number(doc.deref(params?.Colors) ?? space.channels),
        columns: Number(doc.deref(params?.Columns) ?? width)
      })
    } catch (err) {
      return { ...base, skipped: `预测器还原失败：${err.message}` }
    }
  }

  const expected = width * height * space.channels
  if (data.length < expected) {
    return { ...base, skipped: `像素数据不足：期望 ${expected} 字节，实际 ${data.length}` }
  }
  if (data.length > expected) data = data.subarray(0, expected)

  let pixels = data
  let colorType = space.colorType
  let palette = null
  if (space.kind === 'indexed') {
    // 索引色展开成 RGB：PNG 的索引色要求调色板块，展开更省事也更通用
    const expanded = Buffer.alloc(width * height * 3)
    for (let i = 0; i < width * height; i += 1) {
      const at = data[i] * 3
      expanded[i * 3] = space.palette[at] ?? 0
      expanded[i * 3 + 1] = space.palette[at + 1] ?? 0
      expanded[i * 3 + 2] = space.palette[at + 2] ?? 0
    }
    pixels = expanded
    colorType = 2
  }

  let alpha = null
  const smaskRef = dict.SMask
  if (smaskRef !== undefined) {
    const mask = resolveRef(smaskRef)
    const maskDict = dictOf(mask)
    try {
      if (Number(doc.deref(maskDict.Width)) !== width || Number(doc.deref(maskDict.Height)) !== height) {
        return { ...base, skipped: '软掩码尺寸与图像不一致，拒绝导出可能错位的图片' }
      }
      const maskData = doc.decodeStream(mask)
      if (maskData.length < width * height) {
        return { ...base, skipped: '软掩码数据不足' }
      }
      alpha = maskData.subarray(0, width * height)
    } catch (err) {
      return { ...base, skipped: `软掩码解码失败：${err.message}` }
    }
  }

  try {
    const png = encodePng({ width, height, colorType, data: pixels, palette, alpha })
    return {
      ...base,
      format: 'png',
      extension: 'png',
      bytes: png,
      color_space: space.label,
      has_alpha: Boolean(alpha)
    }
  } catch (err) {
    return { ...base, skipped: `PNG 编码失败：${err.message}` }
  }
}

/**
 * 解析色彩空间，给出通道数与 PNG 颜色类型。
 * @param {object} doc - 文档。
 * @param {unknown} raw - `/ColorSpace` 值。
 * @returns {object|null} `{kind, channels, colorType, label, palette?}`。
 */
function resolveColorSpace(doc, raw) {
  const value = doc.deref(raw)
  if (typeof value === 'string') {
    if (value === 'DeviceRGB') return { kind: 'rgb', channels: 3, colorType: 2, label: 'DeviceRGB' }
    if (value === 'DeviceGray' || value === 'G') return { kind: 'gray', channels: 1, colorType: 0, label: 'DeviceGray' }
    return null
  }
  if (Array.isArray(value)) {
    const family = doc.deref(value[0])
    if (family === 'Indexed' || family === 'I') {
      const baseSpace = resolveColorSpace(doc, value[1])
      const lookup = doc.deref(value[3])
      const palette = Buffer.isBuffer(lookup) ? lookup : null
      if (!baseSpace || !palette) return null
      if (baseSpace.channels !== 3) return null // 只展开 RGB 调色板，灰度/CMYK 调色板较少见
      return { kind: 'indexed', channels: 1, colorType: 3, label: 'Indexed', palette }
    }
    if (family === 'ICCBased') {
      const stream = doc.deref(value[1])
      const n = Number(doc.deref(dictOf(stream).N) ?? 0)
      if (n === 3) return { kind: 'rgb', channels: 3, colorType: 2, label: 'ICCBased(3)' }
      if (n === 1) return { kind: 'gray', channels: 1, colorType: 0, label: 'ICCBased(1)' }
      return null
    }
    if (family === 'CalRGB') return { kind: 'rgb', channels: 3, colorType: 2, label: 'CalRGB' }
    if (family === 'CalGray') return { kind: 'gray', channels: 1, colorType: 0, label: 'CalGray' }
    return null
  }
  return null
}

/**
 * 色彩空间的显示名（用于响应里报告，不做解析）。
 * @param {object} doc - 文档。
 * @param {unknown} raw - `/ColorSpace` 值。
 * @returns {unknown} 名称或原始结构。
 */
function colorSpaceName(doc, raw) {
  const value = doc.deref(raw)
  if (Array.isArray(value)) return doc.deref(value[0])
  return value
}

/**
 * 取 `/DecodeParms` 里与 Flate 对应的那一个（可能是数组，按过滤器逐个对应）。
 * @param {object} doc - 文档。
 * @param {unknown} raw - `/DecodeParms` 值。
 * @returns {object|null} 参数字典。
 */
function firstDecodeParms(doc, raw) {
  const value = doc.deref(raw)
  if (!value) return null
  if (Array.isArray(value)) {
    for (const item of value) {
      const resolved = doc.deref(item)
      if (resolved && typeof resolved === 'object') return resolved
    }
    return null
  }
  return typeof value === 'object' ? value : null
}

/**
 * 还原 PNG/TIFF 预测器（PDF 图像流里非常常见）。
 *
 * 支持 `/Predictor 1`（无）、`2`（TIFF 水平差分）、`10–15`（PNG 逐行过滤，15 表示逐行自带过滤字节）。
 *
 * @param {Buffer} data - 已解压的字节。
 * @param {object} args - 参数。
 * @param {number} args.predictor - 预测器编号。
 * @param {number} args.colors - 每像素分量数。
 * @param {number} args.columns - 每行像素数。
 * @returns {Buffer} 还原后的原始像素。
 */
function undoPredictor(data, { predictor, colors, columns }) {
  const bpp = colors // 8 位分量下，每像素字节数 = 分量数
  const rowLength = columns * bpp
  if (predictor === 2) {
    const rows = Math.floor(data.length / rowLength)
    const out = Buffer.from(data.subarray(0, rows * rowLength))
    for (let y = 0; y < rows; y += 1) {
      const start = y * rowLength
      for (let x = bpp; x < rowLength; x += 1) {
        out[start + x] = (out[start + x] + out[start + x - bpp]) & 0xff
      }
    }
    return out
  }
  if (predictor >= 10) {
    const fixedFilter = predictor === 15 ? null : predictor - 10
    const stride = rowLength + 1
    const rows = Math.floor(data.length / stride)
    const out = Buffer.alloc(rows * rowLength)
    let previous = Buffer.alloc(rowLength)
    for (let y = 0; y < rows; y += 1) {
      const filter = fixedFilter ?? data[y * stride]
      const row = data.subarray(y * stride + 1, y * stride + 1 + rowLength)
      const current = Buffer.alloc(rowLength)
      for (let x = 0; x < rowLength; x += 1) {
        const left = x >= bpp ? current[x - bpp] : 0
        const up = previous[x]
        const upLeft = x >= bpp ? previous[x - bpp] : 0
        let value = row[x]
        if (filter === 1) value += left
        else if (filter === 2) value += up
        else if (filter === 3) value += Math.floor((left + up) / 2)
        else if (filter === 4) value += paeth(left, up, upLeft)
        current[x] = value & 0xff
      }
      current.copy(out, y * rowLength)
      previous = current
    }
    return out
  }
  return data
}

/**
 * PNG Paeth 预测器。
 * @param {number} a - 左。
 * @param {number} b - 上。
 * @param {number} c - 左上。
 * @returns {number} 预测值。
 */
function paeth(a, b, c) {
  const p = a + b - c
  const pa = Math.abs(p - a)
  const pb = Math.abs(p - b)
  const pc = Math.abs(p - c)
  if (pa <= pb && pa <= pc) return a
  return pb <= pc ? b : c
}

/**
 * 取一个对象值的字典：流对象在对象表里是 `{dict, stream}` 包装，其余就是字典本身。
 *
 * 这个区分很容易漏：漏了就会「结构合法但字段读不到」（图片计数恒为 0）。
 *
 * @param {unknown} value - 对象值。
 * @returns {object} 字典（拿不到时返回空对象）。
 */
function dictOf(value) {
  if (!value || typeof value !== 'object') return {}
  if (value.dict && typeof value.dict === 'object') return value.dict
  return value
}

/**
 * 按重编号表改写信封里的引用，返回深拷贝（数组/字典/流都走到）。
 *
 * 重写与合并都必须在**序列化之前**完成这一步：对象值里既有原始编号的引用，
 * 也有新建合成对象的引用，混在一起序列化就没法判断哪个该映射。
 *
 * @param {unknown} value - 对象值。
 * @param {Map<string, number>} remap - 旧引用 → 新对象号。
 * @returns {unknown} 改写后的值。
 */
function relabelRefs(value, remap) {
  if (value === null || typeof value !== 'object') return value
  if (Buffer.isBuffer(value)) return value
  if (typeof value.ref === 'string') {
    const mapped = remap.get(value.ref)
    if (mapped === undefined) {
      throw new OfficeError('CORRUPTED_DOCUMENT', `引用 ${value.ref} 不在重编号表里，拒绝产出缺内容的文件。`)
    }
    return { ref: `${mapped} 0` }
  }
  if (Array.isArray(value)) return value.map((item) => relabelRefs(item, remap))
  if (value.stream !== undefined && value.dict && typeof value.dict === 'object') {
    return { dict: relabelRefs(value.dict, remap), stream: value.stream }
  }
  const out = {}
  for (const [key, item] of Object.entries(value)) {
    if (key.startsWith('__')) continue
    out[key] = relabelRefs(item, remap)
  }
  return out
}

/**
 * 把一组已经编号好的对象写成一份完整 PDF（经典 xref 表 + trailer）。
 *
 * 重写与合并共用这一步。对象值里的引用必须**已经是最终编号**（见 `relabelRefs`）：
 * 漏掉生成号（写成 `/Root 1 R`）这类错误会让阅读器把引用当成数字，目录与元数据静默丢失。
 *
 * @param {object} args - 参数。
 * @param {{num: number, value: unknown}[]} args.objects - 对象列表（编号已确定）。
 * @param {number} args.rootNum - 目录对象号。
 * @param {number|null} [args.infoNum] - 元数据对象号（`/Info` 为间接引用时）。
 * @param {object|null} [args.inlineInfo] - 内联的元数据字典（trailer 里直接写字典的形态）。
 * @param {string} args.version - PDF 版本，如 `1.7`。
 * @returns {{buffer: Buffer, bytes: number}} 文件字节。
 */
export function serializePdfFile({ objects, rootNum, infoNum = null, inlineInfo = null, version }) {
  const header = Buffer.from(`%PDF-${version}\n%\xE2\xE3\xCF\xD3\n`, 'latin1')
  const chunks = [header]
  let offset = header.length
  const offsets = new Map()
  for (const item of objects) {
    const body = `${item.num} 0 obj\n${serializePdfObject(item.value)}\nendobj\n`
    const buf = Buffer.from(body, 'latin1')
    offsets.set(item.num, offset)
    chunks.push(buf)
    offset += buf.length
  }

  const maxNum = objects.reduce((max, item) => Math.max(max, item.num), 0)
  const xrefOffset = offset
  let xref = `xref\n0 ${maxNum + 1}\n0000000000 65535 f \n`
  for (let i = 1; i <= maxNum; i += 1) {
    const at = offsets.get(i)
    xref += at === undefined ? '0000000000 65535 f \n' : `${String(at).padStart(10, '0')} 00000 n \n`
  }
  const trailerParts = [`/Size ${maxNum + 1}`, `/Root ${rootNum} 0 R`]
  if (infoNum !== null) trailerParts.push(`/Info ${infoNum} 0 R`)
  else if (inlineInfo) trailerParts.push(`/Info ${serializePdfObject(inlineInfo)}`)
  const tail = `${xref}trailer\n<< ${trailerParts.join(' ')} >>\nstartxref\n${xrefOffset}\n%%EOF\n`
  chunks.push(Buffer.from(tail, 'latin1'))
  const buffer = Buffer.concat(chunks)
  return { buffer, bytes: buffer.length }
}

/**
 * 页面树节点：把 `/Kids` 过滤成只保留选中页，按给定顺序排序，并同步 `/Count`。
 *
 * 重写时不能直接改解析出来的树（它共享、且后续还可能被读），所以返回浅拷贝。
 * 嵌套的 `/Pages` 子节点在这里整体保留 —— 它们会作为独立对象被访问到，各自过滤一次；
 * 排序在**每个节点内部**按 `rank` 进行，因此扁平页面树能得到与 keepPages 完全一致的页序。
 *
 * @param {unknown} value - 对象值。
 * @param {(ref: string) => number} rank - 该引用在目标页序里的位置；不在其中返回 -1。
 * @returns {unknown} 过滤后的值。
 */
function filterPageNode(value, rank) {
  if (!value || typeof value !== 'object') return value
  const dict = value.dict ?? value
  if (!Array.isArray(dict.Kids)) return value
  const kids = dict.Kids.filter((kid) => {
    const ref = kid && typeof kid === 'object' && typeof kid.ref === 'string' ? kid.ref : null
    if (!ref) return true // 非引用（异常结构）原样保留，由可达性检查兜底
    return rank(ref) >= 0
  })
  kids.sort((a, b) => rank(a.ref) - rank(b.ref))
  const merged = { ...dict, Kids: kids, Count: kids.length }
  return value.dict ? { dict: merged, stream: value.stream } : merged
}

/**
 * 清掉指向**被丢弃页面**的引用。
 *
 * 为什么必须做：删页/拆分时把页面从 `/Kids` 摘掉还不够 —— 书签（`/Outlines`）、
 * 命名目标（`/Dests`）、链接批注的 `/Dest`、结构树的 `/Pg` 都可能直接指向那个页面对象。
 * 只要还有一处引用，可达性遍历就会把整页内容（含内容流）写回新文件，
 * 「未选页内容不再留在字节里」这条保证就破了 —— 实测确实如此。
 *
 * 处理方式是按容器删除引用，而不是把目标对象一并带走：
 *   - 数组项直接引用被丢弃页 → 删掉该项（`/Annots`、`/Names` 等）；
 *   - 数组项是目标数组（`/Dest [页 /XYZ …]`）→ 删掉整个目标项；
 *   - 字典里某个键的值引用被丢弃页 → 删掉该键（`/Pg`、`/P` 等可选键）。
 *
 * @param {unknown} value - 对象值。
 * @param {(value: unknown) => boolean} isDroppedPage - 该值是否指向被丢弃的页面。
 * @param {{removed: number}} stats - 统计（原地累加）。
 * @returns {unknown} 清理后的值（浅拷贝）。
 */
function sanitizePageRefs(value, isDroppedPage, stats) {
  if (!value || typeof value !== 'object') return value
  if (Buffer.isBuffer(value)) return value // 字符串/二进制原样返回：Buffer 也是对象，递归会把它拆成数字键
  if (typeof value.ref === 'string') return value // 引用本身不在这里删，由上层容器决定
  const dict = value.dict ?? value
  if (Array.isArray(dict)) {
    const out = []
    for (const item of dict) {
      if (isDroppedPage(item)) {
        stats.removed += 1
        continue
      }
      if (Array.isArray(item) && isDroppedPage(item[0])) {
        stats.removed += 1
        continue
      }
      out.push(sanitizePageRefs(item, isDroppedPage, stats))
    }
    return value.dict ? { dict: out, stream: value.stream } : out
  }
  const copy = {}
  for (const [key, item] of Object.entries(dict)) {
    if (key.startsWith('__')) continue
    if (isDroppedPage(item)) {
      stats.removed += 1
      continue
    }
    if (Array.isArray(item) && isDroppedPage(item[0])) {
      stats.removed += 1
      continue
    }
    copy[key] = sanitizePageRefs(item, isDroppedPage, stats)
  }
  return value.dict ? { dict: copy, stream: value.stream } : copy
}

/**
 * 收集一个对象值里出现的全部间接引用（含流对象的字典）。
 * @param {unknown} value - 对象值。
 * @returns {string[]} `"num gen"` 列表。
 */
function collectRefs(value) {
  const out = []
  const visit = (node, depth) => {
    if (depth > 64 || node === null || node === undefined) return
    if (typeof node !== 'object') return
    if (typeof node.ref === 'string') {
      out.push(node.ref)
      return
    }
    if (Buffer.isBuffer(node)) return
    if (Array.isArray(node)) {
      for (const item of node) visit(item, depth + 1)
      return
    }
    for (const [key, item] of Object.entries(node)) {
      if (key.startsWith('__')) continue
      visit(item, depth + 1)
    }
  }
  visit(value, 0)
  return out
}

/**
 * 格式化数字。
 * @param {number} value - 数字。
 * @returns {string} 文本。
 */
function formatPdfNumber(value) {
  if (!Number.isFinite(value)) throw new OfficeError('INVALID_REQUEST', `PDF 数字不是有限值：${value}`)
  if (Number.isInteger(value)) return String(value)
  return value.toFixed(4).replace(/0+$/, '').replace(/\.$/, '')
}

/**
 * 转义名称里的分隔符与不可打印字节。
 * @param {string} name - 名称（latin1 字节序列）。
 * @returns {string} 转义后的名称。
 */
function serializePdfName(name) {
  let out = ''
  for (const byte of Buffer.from(String(name), 'latin1')) {
    out += byte < 0x21 || byte > 0x7e || isDelimiterOrWhitespace(byte) || byte === 0x23
      ? `#${byte.toString(16).toUpperCase().padStart(2, '0')}`
      : String.fromCharCode(byte)
  }
  return out
}

/**
 * 写字符串。可打印 ASCII 用字面串，其余用十六进制串避免编码歧义。
 * @param {Buffer} buf - 字节。
 * @returns {string} PDF 字符串。
 */
function serializePdfBuffer(buf) {
  const printable = buf.every((byte) => byte >= 0x20 && byte <= 0x7e && byte !== 0x28 && byte !== 0x29 && byte !== 0x5c)
  if (!printable) return `<${buf.toString('hex').toUpperCase()}>`
  return `(${buf.toString('latin1').replace(/[\\()]/g, (ch) => `\\${ch}`)})`
}

/**
 * 把文本编码成 PDF 字符串字节。
 *
 * 纯 ASCII 用 PDFDocEncoding 的字节（写出来是字面串）；含非 ASCII 时用
 * UTF-16BE + BOM（`FE FF`），序列化器会把它写成十六进制串 —— 这是 PDF 里
 * 表示中文等非 ASCII 文本的标准做法，阅读器靠 BOM 识别编码。
 *
 * @param {string} text - 文本。
 * @returns {Buffer} 字节。
 */
function encodePdfText(text) {
  if (/^[\x00-\x7F]*$/.test(text)) return Buffer.from(text, 'latin1')
  const body = Buffer.alloc(text.length * 2)
  for (let i = 0; i < text.length; i += 1) body.writeUInt16BE(text.charCodeAt(i), i * 2)
  return Buffer.concat([Buffer.from([0xfe, 0xff]), body])
}

/**
 * 生成叠加文本的内容流指令。
 *
 * 只用 `q` / `Q` 保存恢复图形状态、`g` 设灰度、`BT`/`ET` 包住文本、
 * `Tm` 定位、`Tj` 出字 —— 全部是 PDF 1.0 就有的操作符，不依赖透明度组或可选内容组。
 *
 * 文本宽度按 Helvetica 的经验平均字宽（0.5 em）估算，用于水平居中；
 * 水印对居中精度不敏感，但估算法会在注释里说明，而不是假装是精确排版。
 *
 * @param {object} args - 参数。
 * @param {string} args.label - 实际文本。
 * @param {string} args.position - `center` | `top` | `bottom`。
 * @param {number} args.size - 字号。
 * @param {number} args.gray - 灰度。
 * @param {number} args.margin - 页边距（pt）。
 * @param {number} args.width - 页宽。
 * @param {number} args.height - 页高。
 * @returns {string} 内容流文本。
 */
function buildOverlayCommands({ label, position, size, gray, margin, width, height }) {
  const literal = serializePdfBuffer(Buffer.from(label, 'latin1'))
  const estimated = label.length * size * 0.5
  const grayText = formatPdfNumber(Math.min(1, Math.max(0, gray)))
  if (position === 'center') {
    // 45° 斜向：cos45 = sin45 = 0.7071
    const x = (width - estimated * 0.7071) / 2
    const y = (height - estimated * 0.7071) / 2
    return `q\n${grayText} g\nBT\n/Wm ${formatPdfNumber(size)} Tf\n0.7071 0.7071 -0.7071 0.7071 ${formatPdfNumber(x)} ${formatPdfNumber(y)} Tm\n${literal} Tj\nET\nQ\n`
  }
  const x = Math.max(0, (width - estimated) / 2)
  const y = position === 'top' ? height - margin - size : margin
  return `q\n${grayText} g\nBT\n/Wm ${formatPdfNumber(size)} Tf\n1 0 0 1 ${formatPdfNumber(x)} ${formatPdfNumber(y)} Tm\n${literal} Tj\nET\nQ\n`
}

/** 本阶段不做的检查项，如实列出而不是假装通过。 */
const PDF_NOT_CHECKED = Object.freeze([
  { name: '交叉引用表一致性', reason: '本实现按对象扫描而非交叉引用遍历，因此不做 xref 校验（换取对截断/增量更新文件的容错）' },
  { name: '数字签名有效性', reason: '只检测签名是否存在，不验证签名链与是否被篡改' },
  { name: '字体嵌入完整性', reason: '只统计字体，不校验嵌入子集是否完整' },
  { name: '版面结构还原', reason: '文本按内容流顺序提取，不还原栏位、表格与坐标' },
  { name: '视觉外观', reason: '需要渲染成图片后比对' }
])

/**
 * 扫描整个文件，建立 `"num gen" → {value, num, gen}` 的对象表。
 *
 * @param {Buffer} buffer - 文件字节。
 * @param {object} limits - 上限。
 * @returns {Map<string, object>} 对象表。
 */
function scanObjects(buffer, limits) {
  const objects = new Map()
  const text = buffer.toString('latin1')
  const matches = [...text.matchAll(OBJ_RE)]
  if (matches.length > limits.maxObjects) {
    throw new OfficeError('MEMORY_LIMIT', `对象数 ${matches.length} 超过上限 ${limits.maxObjects}。`)
  }

  for (const match of matches) {
    const num = Number(match[1])
    const gen = Number(match[2])
    const bodyStart = match.index + match[0].length
    // 同名对象重复出现时保留最后一个（增量更新会追加新版本）
    try {
      const parsed = parseValue(buffer, bodyStart, limits)
      const value = parsed.value
      if (value && typeof value === 'object' && typeof value.__streamStart === 'number') {
        const stream = readStreamBytes(buffer, value.__streamStart, parsed.end)
        delete value.__streamStart
        objects.set(`${num} ${gen}`, { num, gen, value: { dict: value, stream, __id: `${num} ${gen}` } })
      } else {
        objects.set(`${num} ${gen}`, { num, gen, value })
      }
    } catch {
      // 单个对象解析失败不应让整个文件打不开：跳过，
      // 由 validate() 的检查项与 warnings 反映出来
    }
  }

  // 收集 trailer 字典（可能有多段，取最后一个）
  const trailerRe = /trailer\b/g
  const trailerMatches = [...text.matchAll(trailerRe)]
  for (const match of trailerMatches) {
    try {
      const parsed = parseValue(buffer, match.index + match[0].length, limits)
      if (parsed.value && typeof parsed.value === 'object') {
        objects.set(`trailer ${match.index}`, { num: -1, gen: 0, value: { ...parsed.value, __trailer: true } })
      }
    } catch {
      // 忽略
    }
  }
  if (trailerMatches.length === 0) {
    // PDF 1.5+ 的 trailer 在交叉引用流里，没有 `trailer` 关键字
    for (const entry of objects.values()) {
      const value = entry.value?.dict ?? entry.value
      if (value && typeof value === 'object' && value.Type === 'XRef') {
        objects.set('trailer-from-xref', { num: -2, gen: 0, value: { ...value, __trailer: true } })
        break
      }
    }
  }

  // 展开对象流（ObjStm）里的成员对象：现代 PDF 的目录与页面常在里面
  expandObjectStreams(buffer, objects, limits)
  return objects
}

/**
 * 展开全部对象流，把成员对象补进对象表（不覆盖已存在的直接对象）。
 * @param {Buffer} buffer - 文件字节。
 * @param {Map<string, object>} objects - 对象表。
 * @param {object} limits - 上限。
 * @returns {void}
 */
function expandObjectStreams(buffer, objects, limits) {
  for (const entry of [...objects.values()]) {
    const wrapper = entry.value
    if (!wrapper || typeof wrapper !== 'object' || wrapper.stream === undefined) continue
    const dict = wrapper.dict
    if (dict.Type !== 'ObjStm') continue
    let data
    try {
      const filters = normalizeFilters(dict.Filter)
      data = wrapper.stream
      for (const filter of filters) {
        if (filter === 'FlateDecode' || filter === 'Fl') {
          data = zlib.inflateSync(data, { maxOutputLength: limits.maxStreamBytes })
        } else {
          data = null
          break
        }
      }
    } catch {
      continue
    }
    if (!data) continue
    const n = Number(dict.N)
    const first = Number(dict.First)
    if (!Number.isFinite(n) || !Number.isFinite(first)) continue
    const header = data.subarray(0, first).toString('latin1')
    const pairs = header.trim().split(/\s+/).map(Number).filter((x) => Number.isFinite(x))
    for (let i = 0; i < n; i += 1) {
      const objNum = pairs[i * 2]
      const offset = pairs[i * 2 + 1]
      if (!Number.isFinite(objNum) || !Number.isFinite(offset)) continue
      if (objects.has(`${objNum} 0`)) continue // 直接对象优先
      try {
        const parsed = parseValue(data, first + offset, limits)
        objects.set(`${objNum} 0`, { num: objNum, gen: 0, value: parsed.value })
      } catch {
        // 忽略单个成员
      }
    }
  }
}

/**
 * 解析一个 PDF 对象值。
 *
 * 支持：字典、数组、名称、字面字符串、十六进制字符串、数字、布尔、null、
 * 间接引用（`N G R`），以及字典后紧跟的 `stream ... endstream`。
 *
 * @param {Buffer} buf - 数据。
 * @param {number} start - 起始偏移。
 * @param {object} limits - 上限。
 * @returns {{value: unknown, end: number}} 解析结果。
 */
function parseValue(buf, start, limits) {
  let pos = skipWhitespaceAndComments(buf, start)
  if (pos >= buf.length) return { value: null, end: pos }

  const byte = buf[pos]

  if (byte === 0x3c && buf[pos + 1] === 0x3c) {
    const { dict, end } = parseDict(buf, pos, limits)
    const after = skipWhitespaceAndComments(buf, end)
    if (buf.subarray(after, after + 6).toString('latin1') === 'stream') {
      let dataStart = after + 6
      if (buf[dataStart] === 0x0d) dataStart += 1
      if (buf[dataStart] === 0x0a) dataStart += 1
      return { value: { ...dict, __streamStart: dataStart }, end }
    }
    return { value: dict, end }
  }
  if (byte === 0x5b) {
    const arr = []
    pos += 1
    for (;;) {
      pos = skipWhitespaceAndComments(buf, pos)
      if (pos >= buf.length) break
      if (buf[pos] === 0x5d) {
        pos += 1
        break
      }
      const parsed = parseValue(buf, pos, limits)
      arr.push(parsed.value)
      pos = parsed.end
    }
    return { value: arr, end: pos }
  }
  if (byte === 0x2f) {
    pos += 1
    const startName = pos
    while (pos < buf.length && !isDelimiterOrWhitespace(buf[pos])) pos += 1
    return { value: decodeName(buf.subarray(startName, pos).toString('latin1')), end: pos }
  }
  if (byte === 0x28) {
    const { text, end } = parseLiteralString(buf, pos)
    return { value: text, end }
  }
  if (byte === 0x3c) {
    const end = buf.indexOf(0x3e, pos + 1)
    if (end === -1) throw new OfficeError('CORRUPTED_DOCUMENT', '十六进制字符串未闭合。')
    const hex = stripWhitespace(buf.subarray(pos + 1, end).toString('latin1'))
    return { value: Buffer.from(hex.length % 2 === 0 ? hex : `${hex}0`, 'hex'), end: end + 1 }
  }
  if (byte === 0x5d || byte === 0x3e) return { value: null, end: pos + 1 }

  const wordStart = pos
  while (pos < buf.length && !isDelimiterOrWhitespace(buf[pos])) pos += 1
  const word = buf.subarray(wordStart, pos).toString('latin1')
  if (word === 'true') return { value: true, end: pos }
  if (word === 'false') return { value: false, end: pos }
  if (word === 'null') return { value: null, end: pos }
  if (/^[+-]?\d+$/.test(word)) {
    // 可能是间接引用 `N G R`
    const save = pos
    const p1 = skipWhitespaceAndComments(buf, pos)
    let p2 = p1
    while (p2 < buf.length && !isDelimiterOrWhitespace(buf[p2])) p2 += 1
    const genWord = buf.subarray(p1, p2).toString('latin1')
    if (/^\d+$/.test(genWord)) {
      const p3 = skipWhitespaceAndComments(buf, p2)
      if (buf[p3] === 0x52 /* R */ && isDelimiterOrWhitespace(buf[p3 + 1] ?? 0x20)) {
        return { value: { ref: `${word} ${genWord}` }, end: p3 + 1 }
      }
    }
    pos = save
    return { value: Number(word), end: pos }
  }
  if (/^[+-]?(\d*\.\d*|\d+)$/.test(word)) return { value: Number(word), end: pos }
  return { value: word, end: pos }
}

/**
 * 解析字典。
 * @param {Buffer} buf - 数据。
 * @param {number} start - `<<` 的偏移。
 * @param {object} limits - 上限。
 * @returns {{dict: object, end: number}} 字典与结束偏移。
 */
function parseDict(buf, start, limits) {
  const dict = {}
  let pos = start + 2
  for (;;) {
    pos = skipWhitespaceAndComments(buf, pos)
    if (pos >= buf.length) throw new OfficeError('CORRUPTED_DOCUMENT', '字典未闭合。')
    if (buf[pos] === 0x3e && buf[pos + 1] === 0x3e) {
      pos += 2
      break
    }
    if (buf[pos] !== 0x2f) {
      // 容忍非名称键：跳过以恢复
      const skipped = parseValue(buf, pos, limits)
      pos = skipped.end
      continue
    }
    const keyParsed = parseValue(buf, pos, limits)
    const valueParsed = parseValue(buf, keyParsed.end, limits)
    dict[String(keyParsed.value)] = valueParsed.value
    pos = valueParsed.end
  }
  return { dict, end: pos }
}

/**
 * 解析字面字符串 `(...)`，处理转义与嵌套括号。
 * @param {Buffer} buf - 数据。
 * @param {number} start - `(` 的偏移。
 * @returns {{text: Buffer, end: number}} 字节内容与结束偏移。
 */
function parseLiteralString(buf, start) {
  const out = []
  let depth = 1
  let pos = start + 1
  while (pos < buf.length) {
    const b = buf[pos]
    if (b === 0x5c /* \ */) {
      const next = buf[pos + 1]
      const map = { 0x6e: 0x0a, 0x72: 0x0d, 0x74: 0x09, 0x62: 0x08, 0x66: 0x0c }
      if (map[next] !== undefined) {
        out.push(map[next])
        pos += 2
        continue
      }
      if (next >= 0x30 && next <= 0x37) {
        let oct = ''
        let p = pos + 1
        while (p < buf.length && oct.length < 3 && buf[p] >= 0x30 && buf[p] <= 0x37) {
          oct += String.fromCharCode(buf[p])
          p += 1
        }
        out.push(Number.parseInt(oct, 8) & 0xff)
        pos = p
        continue
      }
      out.push(next)
      pos += 2
      continue
    }
    if (b === 0x28) depth += 1
    if (b === 0x29) {
      depth -= 1
      if (depth === 0) {
        pos += 1
        break
      }
    }
    out.push(b)
    pos += 1
  }
  return { text: Buffer.from(out), end: pos }
}

/**
 * 读取 `stream ... endstream` 之间的字节。
 * @param {Buffer} buf - 数据。
 * @param {number} start - 流数据起点。
 * @param {number} [dictEnd] - 字典结束偏移（保留参数）。
 * @returns {Buffer} 流字节。
 */
function readStreamBytes(buf, start, dictEnd) {
  void dictEnd
  const end = buf.indexOf('endstream', start, 'latin1')
  if (end === -1) return buf.subarray(start)
  // 去掉 endstream 之前的一个换行
  let stop = end
  if (buf[stop - 1] === 0x0a) stop -= 1
  if (buf[stop - 1] === 0x0d) stop -= 1
  return buf.subarray(start, stop)
}

/**
 * 判断是否 PDF 分隔符或空白。
 * @param {number} byte - 字节。
 * @returns {boolean} 是否分隔。
 */
function isDelimiterOrWhitespace(byte) {
  return (
    byte === 0x20 ||
    byte === 0x0a ||
    byte === 0x0d ||
    byte === 0x09 ||
    byte === 0x0c ||
    byte === 0x00 ||
    byte === 0x28 ||
    byte === 0x29 ||
    byte === 0x3c ||
    byte === 0x3e ||
    byte === 0x5b ||
    byte === 0x5d ||
    byte === 0x7b ||
    byte === 0x7d ||
    byte === 0x2f ||
    byte === 0x25
  )
}

/**
 * 跳过空白与注释。
 * @param {Buffer} buf - 数据。
 * @param {number} pos - 起始位置。
 * @returns {number} 新位置。
 */
function skipWhitespaceAndComments(buf, pos) {
  let p = pos
  for (;;) {
    while (p < buf.length && (buf[p] === 0x20 || buf[p] === 0x0a || buf[p] === 0x0d || buf[p] === 0x09 || buf[p] === 0x0c || buf[p] === 0x00)) p += 1
    if (buf[p] === 0x25 /* % */) {
      while (p < buf.length && buf[p] !== 0x0a && buf[p] !== 0x0d) p += 1
      continue
    }
    return p
  }
}

/**
 * 解码 PDF 名称里的 `#XX` 转义。
 * @param {string} raw - 原始名称。
 * @returns {string} 解码后的名称。
 */
function decodeName(raw) {
  return raw.replace(/#([0-9A-Fa-f]{2})/g, (_, hex) => String.fromCharCode(Number.parseInt(hex, 16)))
}

/**
 * 解码 PDF 字符串（字面或十六进制字节）为 JS 字符串。
 *
 * 优先按 UTF-16BE（有 BOM 时），否则按 PDFDocEncoding 的常见子集处理。
 *
 * @param {unknown} value - 字符串值。
 * @returns {string} 文本。
 */
function decodePdfString(value) {
  if (value === null || value === undefined) return null
  if (typeof value === 'string') return value
  if (!Buffer.isBuffer(value)) return String(value)
  if (value.length >= 2 && value[0] === 0xfe && value[1] === 0xff) {
    let out = ''
    for (let i = 2; i + 1 < value.length; i += 2) out += String.fromCharCode(value.readUInt16BE(i))
    return out
  }
  return value.toString('latin1')
}

/**
 * 归一化 `/Filter` 为数组形式。
 * @param {unknown} filter - Filter 值。
 * @returns {string[]} 过滤器名列表。
 */
function normalizeFilters(filter) {
  if (filter === undefined || filter === null) return []
  return Array.isArray(filter) ? filter.map(String) : [String(filter)]
}

/**
 * 去掉所有空白字符。
 * @param {string} text - 文本。
 * @returns {string} 结果。
 */
function stripWhitespace(text) {
  return text.replace(/\s+/g, '')
}

/**
 * 解码 ASCII85。
 * @param {string} text - 编码文本。
 * @returns {Buffer} 解码字节。
 */
function decodeAscii85(text) {
  const clean = text.replace(/\s+/g, '').replace(/^<~/, '').replace(/~>$/, '')
  const out = []
  let tuple = 0
  let count = 0
  for (const ch of clean) {
    if (ch === 'z' && count === 0) {
      out.push(0, 0, 0, 0)
      continue
    }
    const code = ch.charCodeAt(0) - 33
    if (code < 0 || code > 84) continue
    tuple = tuple * 85 + code
    count += 1
    if (count === 5) {
      out.push((tuple >>> 24) & 0xff, (tuple >>> 16) & 0xff, (tuple >>> 8) & 0xff, tuple & 0xff)
      tuple = 0
      count = 0
    }
  }
  if (count > 0) {
    for (let i = count; i < 5; i += 1) tuple = tuple * 85 + 84
    const bytes = [(tuple >>> 24) & 0xff, (tuple >>> 16) & 0xff, (tuple >>> 8) & 0xff, tuple & 0xff]
    out.push(...bytes.slice(0, count - 1))
  }
  return Buffer.from(out)
}

/**
 * 从内容流中提取文本。
 *
 * @param {PdfDocument} doc - 文档对象。
 * @param {Buffer} content - 内容字节。
 * @param {object} pageDict - 页面字典（用于取字体资源）。
 * @returns {string} 文本。
 */
function extractTextFromContent(doc, content, pageDict) {
  const text = content.toString('latin1')
  const cmaps = loadToUnicodeMaps(doc, pageDict)
  // 状态必须是本函数局部的：放模块级会被并发调用与跨页调用污染
  let currentMap = null
  let pendingFont = null
  const pendingStrings = []
  const out = []
  let i = 0
  let inText = false
  let lastWasText = false

  while (i < text.length) {
    const ch = text[i]

    if (ch === '%') {
      while (i < text.length && text[i] !== '\n' && text[i] !== '\r') i += 1
      continue
    }
    if (/\s/.test(ch)) {
      i += 1
      continue
    }

    // 操作符或操作数
    if (ch === '/' || ch === '(' || ch === '[' || ch === '<' || /[0-9+\-.]/.test(ch)) {
      const parsed = readOperand(text, i)
      if (parsed) {
        // 记录最近一个名称操作数：`/F1 10.5 Tf` 里 `/F1` 是字体名、`Tf` 才是操作符，
        // 把「名称等于 Tf」当作条件会让字体名永远记录不上（曾导致中文全成乱码）。
        if (parsed.kind === 'name') pendingFont = parsed
        else if (parsed.kind === 'string') pendingStrings.push(parsed.bytes)
        else if (parsed.kind === 'array') pendingStrings.push(...parsed.strings)
        i = parsed.end
        continue
      }
    }

    const startOp = i
    while (i < text.length && !/[\s/[\]<>()]/.test(text[i])) i += 1
    const op = text.slice(startOp, i)
    if (op === '') {
      i += 1
      continue
    }

    if (op === 'BT') inText = true
    else if (op === 'ET') {
      inText = false
      if (lastWasText) out.push('\n')
      lastWasText = false
    } else if (op === 'Tf') {
      if (pendingFont) currentMap = cmaps.get(pendingFont.name) ?? null
      pendingFont = null
    } else if (op === 'Tj' || op === 'TJ' || op === "'" || op === '"') {
      if (inText && pendingStrings.length > 0) {
        let chunk = ''
        for (const bytes of pendingStrings) chunk += decodeWithMap(bytes, currentMap)
        if (chunk !== '') {
          out.push(chunk)
          lastWasText = true
        }
      }
      if (op === "'" || op === '"') {
        out.push('\n')
        lastWasText = false
      }
    } else if (op === 'Td' || op === 'TD' || op === 'T*' || op === 'Tm') {
      // `Tm` 同样意味着「另起一行」：从零生成的 PDF 与很多库都用 `Tm` 定位每一行，
      // 不把它当换行会让整页文本粘成一行。
      if (lastWasText) {
        out.push('\n')
        lastWasText = false
      }
    }
    pendingStrings.length = 0
  }
  return out.join('').replace(/\n{3,}/g, '\n\n').trim()
}

/**
 * 读一个操作数（名称、字符串、数组、数字）。
 * @param {string} text - 内容流文本。
 * @param {number} start - 起始位置。
 * @returns {object|null} 操作数描述。
 */
function readOperand(text, start) {
  const ch = text[start]
  if (ch === '/') {
    let i = start + 1
    let name = ''
    while (i < text.length && !/[\s/[\]<>()]/.test(text[i])) {
      name += text[i]
      i += 1
    }
    return { kind: 'name', name: decodeName(name), end: i }
  }
  if (ch === '(') {
    const { bytes, end } = readLiteralStringFromText(text, start)
    return { kind: 'string', bytes, end }
  }
  if (ch === '<' && text[start + 1] !== '<') {
    const end = text.indexOf('>', start + 1)
    if (end === -1) return null
    const hex = stripWhitespace(text.slice(start + 1, end))
    return { kind: 'string', bytes: Buffer.from(hex.length % 2 === 0 ? hex : `${hex}0`, 'hex'), end: end + 1 }
  }
  if (ch === '[') {
    const strings = []
    let i = start + 1
    while (i < text.length && text[i] !== ']') {
      if (text[i] === '(') {
        const parsed = readLiteralStringFromText(text, i)
        strings.push(parsed.bytes)
        i = parsed.end
        continue
      }
      if (text[i] === '<' && text[i + 1] !== '<') {
        const end = text.indexOf('>', i + 1)
        if (end === -1) break
        const hex = stripWhitespace(text.slice(i + 1, end))
        strings.push(Buffer.from(hex.length % 2 === 0 ? hex : `${hex}0`, 'hex'))
        i = end + 1
        continue
      }
      i += 1
    }
    return { kind: 'array', strings, end: i + 1 }
  }
  if (/[0-9+\-.]/.test(ch)) {
    let i = start
    while (i < text.length && /[0-9+\-.eE]/.test(text[i])) i += 1
    return { kind: 'number', value: Number(text.slice(start, i)), end: i }
  }
  return null
}

/**
 * 提取**带坐标**的文本片段。
 *
 * 与 {@link PdfDocument#extractText} 的区别：这里跟踪文本矩阵（`BT`/`Tm`/`Td`/`TD`/`T*`），
 * 因此每个片段都带页内坐标，可用于版面判断（表格还原、按位置检索）。
 *
 * 如实说明的近似：
 *   - 坐标只取平移分量（`Tm` 的 e/f），**不还原旋转与缩放**；旋转文本的坐标会偏。
 *   - 同一行里多个片段若没有显式定位，x 只能按「字符数 × 字号 × 0.5」估算推进
 *     （嵌入字体没有可靠的宽度表），因此**同格内的先后顺序可靠、精确间距不可靠**。
 *
 * @param {object} [options] - 选项。
 * @param {number|null} [options.page] - 只取某一页（0 基）；省略取全部页。
 * @returns {object} `{page_count, runs, approximation}`。
 */
function extractPositionedRuns(doc, content, pageDict, pageIndex) {
  const text = content.toString('latin1')
  const cmaps = loadToUnicodeMaps(doc, pageDict)
  const runs = []
  let currentMap = null
  let pendingFontName = null
  let fontSize = 0
  let tm = [1, 0, 0, 1, 0, 0]
  let tlm = [1, 0, 0, 1, 0, 0]
  let leading = 0
  const numbers = []
  const strings = []
  let i = 0
  let inText = false

  const applyTd = (tx, ty) => {
    tlm = [tlm[0], tlm[1], tlm[2], tlm[3], tlm[4] + tx * tlm[0] + ty * tlm[2], tlm[5] + tx * tlm[1] + ty * tlm[3]]
    tm = [...tlm]
  }

  while (i < text.length) {
    const ch = text[i]
    if (ch === '%') {
      while (i < text.length && text[i] !== '\n' && text[i] !== '\r') i += 1
      continue
    }
    if (/\s/.test(ch)) {
      i += 1
      continue
    }
    if (ch === '/' || ch === '(' || ch === '[' || ch === '<' || /[0-9+\-.]/.test(ch)) {
      const parsed = readOperand(text, i)
      if (parsed) {
        if (parsed.kind === 'name') pendingFontName = parsed.name
        else if (parsed.kind === 'string') strings.push(parsed.bytes)
        else if (parsed.kind === 'array') strings.push(...parsed.strings)
        else if (parsed.kind === 'number') numbers.push(parsed.value)
        i = parsed.end
        continue
      }
    }
    const startOp = i
    while (i < text.length && !/[\s/[\]<>()]/.test(text[i])) i += 1
    const op = text.slice(startOp, i)
    if (op === '') {
      i += 1
      continue
    }

    if (op === 'BT') {
      inText = true
      tm = [1, 0, 0, 1, 0, 0]
      tlm = [...tm]
    } else if (op === 'ET') {
      inText = false
    } else if (op === 'Tf') {
      if (pendingFontName) currentMap = cmaps.get(pendingFontName) ?? null
      if (numbers.length > 0) fontSize = numbers[numbers.length - 1]
      pendingFontName = null
    } else if (op === 'Tm' && numbers.length >= 6) {
      const six = numbers.slice(-6)
      tm = six
      tlm = [...six]
    } else if ((op === 'Td' || op === 'TD') && numbers.length >= 2) {
      const [tx, ty] = numbers.slice(-2)
      if (op === 'TD') leading = -ty
      applyTd(tx, ty)
    } else if (op === 'TL' && numbers.length >= 1) {
      leading = numbers[numbers.length - 1]
    } else if (op === 'T*') {
      applyTd(0, -leading)
    } else if (op === 'Tj' || op === 'TJ' || op === "'" || op === '"') {
      if (op === "'" || op === '"') applyTd(0, -leading)
      if (inText && strings.length > 0) {
        let chunk = ''
        for (const bytes of strings) chunk += decodeWithMap(bytes, currentMap)
        if (chunk.trim() !== '') {
          runs.push({
            page: pageIndex,
            text: chunk,
            x: Math.round(tm[4] * 100) / 100,
            y: Math.round(tm[5] * 100) / 100,
            font_size: fontSize,
            font: pendingFontOrNull(cmaps, currentMap)
          })
        }
      }
    }
    numbers.length = 0
    strings.length = 0
  }
  return runs
}

/**
 * 反查某个 ToUnicode 映射对应的字体名（仅用于报告，拿不到就返回 null）。
 * @param {Map} cmaps - 字体名 → 映射。
 * @param {Map|null} map - 当前映射。
 * @returns {string|null} 字体名。
 */
function pendingFontOrNull(cmaps, map) {
  if (!map) return null
  for (const [name, candidate] of cmaps) {
    if (candidate === map) return name
  }
  return null
}

/**
 * 提取文本片段并按版面聚成表格。
 *
 * 做法（**纯位置推断**，不依赖框线）：
 *   1. 先从 y 坐标的间隔推出**行距**，据此得到行容差（同一行的基线可能相差几个点：
 *      同一行的数字与汉字基线常不完全一致，写死一个容差会把一行拆成两行）；
 *   2. 按 y 聚类成行，按行距倍数切出多张表；
 *   3. 把所有行的片段起点按 x 聚成列，行 × 列得到网格，裁掉全空行与全空列。
 *
 * 明确不做（会写进返回值的 `not_done`）：跨页表格合并、合并单元格的跨行跨列还原、
 * 扫描件（没有文本层）、只有图片的表格。
 *
 * @param {object[]} runs - 文本片段（带 x/y）。
 * @param {object} [options] - 选项。
 * @param {number|null} [options.rowTolerance] - 同一行的 y 容差；省略则按行距自动推导。
 * @param {number} [options.columnGap] - 同一列的 x 容差（点），默认 8。
 * @param {number} [options.minRows] - 至少几行才算表格，默认 2。
 * @param {number} [options.minColumns] - 至少几列才算表格，默认 2。
 * @returns {object[]} 表格列表。
 */
function extractTablesFromRuns(runs, { rowTolerance = null, columnGap = 8, minRows = 2, minColumns = 2 } = {}) {
  const tables = []
  const pages = [...new Set(runs.map((r) => r.page))].sort((a, b) => a - b)
  for (const page of pages) {
    const pageRuns = runs.filter((r) => r.page === page && r.text.trim() !== '')
    // 注意：这里不能用「片段数 < 行数×列数」提前退出 —— 靠空格对齐的行只有一个片段，
    // 展开成列之前片段数会比表格的单元格数少得多。
    if (pageRuns.length === 0) continue

    // 1) 行距与行容差
    const ys = [...new Set(pageRuns.map((r) => Math.round(r.y * 100) / 100))].sort((a, b) => b - a)
    const gaps = []
    for (let i = 1; i < ys.length; i += 1) {
      const gap = ys[i - 1] - ys[i]
      if (gap > 0.5) gaps.push(gap)
    }
    gaps.sort((a, b) => a - b)
    const medianGap = gaps.length > 0 ? gaps[Math.floor(gaps.length / 2)] : 12
    const sortedGaps = [...gaps].sort((a, b) => a - b)
    const pitch = sortedGaps.length > 0 ? sortedGaps[0] : medianGap
    const tolerance = rowTolerance ?? Math.max(2, Math.min(medianGap * 0.45, 12))

    // 2) 行聚类（y 从大到小：PDF 坐标原点在左下）
    const sorted = [...pageRuns].sort((a, b) => b.y - a.y || a.x - b.x)
    const rows = []
    for (const run of sorted) {
      const last = rows[rows.length - 1]
      if (last && Math.abs(last.y - run.y) <= tolerance) {
        last.runs.push(run)
      } else {
        rows.push({ y: run.y, runs: [run] })
      }
    }
    for (const row of rows) {
      row.runs.sort((a, b) => a.x - b.x)
      // 有些生成器（包括本插件自己的 PDF 生成器）会把一整行的多列写成**一个**字符串，
      // 靠连续空格对齐。这类行按「两个以上连续空格」切成虚拟列，否则等宽对齐的表格
      // 会被当成单列文本而漏掉。
      row.runs = expandWhitespaceColumns(row.runs)
      // 这一行里是否存在「真正的列间隔」：下一个片段起点位于上一个片段估算结尾之后足够远。
      // 正文里的词间距（几个点）不算列间隔，否则每一行散文都会被当成表格。
      let widest = 0
      for (const run of row.runs) widest = Math.max(widest, run.font_size ?? 0)
      const threshold = Math.max(columnGap, widest * 1.2)
      row.gaps = 0
      for (let i = 1; i < row.runs.length; i += 1) {
        const previous = row.runs[i - 1]
        const end = previous.x + estimateTextWidth(previous.text, previous.font_size)
        if (row.runs[i].x - end > threshold) row.gaps += 1
      }
    }

    // 3) 连续的「表格行」组成一张表；完全对齐已知列的稀疏行（空单元格/合并单元格）也留在表内
    const blocks = []
    let block = []
    let columnStarts = []
    for (const row of rows) {
      const alignsWithBlock =
        block.length > 0 &&
        row.runs.length > 0 &&
        // 「对齐」必须近乎精确（2 pt）：空的/合并的单元格与列起点是同一条基线，
        // 而正文段落只是碰巧落在列起点附近（实测差 5 pt 就会被误并进表格）。
        row.runs.every((run) => columnStarts.some((start) => Math.abs(start - run.x) <= Math.min(columnGap, 2)))
      if (row.gaps > 0 || alignsWithBlock) {
        block.push(row)
        for (const run of row.runs) {
          if (!columnStarts.some((start) => Math.abs(start - run.x) <= columnGap)) columnStarts.push(run.x)
        }
        columnStarts.sort((a, b) => a - b)
        continue
      }
      if (block.length > 0) blocks.push(block)
      block = []
      columnStarts = []
    }
    if (block.length > 0) blocks.push(block)

    for (const candidate of blocks) {
      if (candidate.length < minRows) continue
      const starts = []
      for (const row of candidate) {
        for (const run of row.runs) {
          if (!starts.some((start) => Math.abs(start - run.x) <= columnGap)) starts.push(run.x)
        }
      }
      starts.sort((a, b) => a - b)
      if (starts.length < minColumns) continue

      const grid = candidate.map((row) => {
        const cells = new Array(starts.length).fill('')
        for (const run of row.runs) {
          let best = 0
          let bestDistance = Infinity
          starts.forEach((start, index) => {
            const distance = Math.abs(start - run.x)
            if (distance < bestDistance) {
              bestDistance = distance
              best = index
            }
          })
          cells[best] = joinRuns(cells[best], run.text)
        }
        return cells
      })
      const keep = starts.map((_, index) => grid.some((row) => row[index].trim() !== ''))
      const trimmed = grid.map((row) => row.filter((_, index) => keep[index]))
      const columns = keep.filter(Boolean).length
      if (columns < minColumns) continue
      const nonEmptyRows = trimmed.filter((row) => row.some((cell) => cell.trim() !== ''))
      if (nonEmptyRows.length < minRows) continue
      tables.push({
        page,
        row_count: nonEmptyRows.length,
        column_count: columns,
        rows: nonEmptyRows,
        y_top: Math.round(candidate[0].y * 100) / 100,
        y_bottom: Math.round(candidate[candidate.length - 1].y * 100) / 100,
        row_tolerance: Math.round(tolerance * 100) / 100
      })
    }
  }
  return tables
}

/**
 * 把一个片段按「两个以上连续空格」切成多个虚拟片段（列）。
 *
 * 位置用估算宽度累加得出：只用于把「靠空格对齐」的表格认出来，
 * 精度不影响单元格内容（内容是按空格切出来的原文）。
 *
 * @param {object[]} runs - 同一行的片段（已按 x 排序）。
 * @returns {object[]} 展开后的片段。
 */
function expandWhitespaceColumns(runs) {
  const out = []
  for (const run of runs) {
    if (!/ {2,}/.test(run.text)) {
      out.push(run)
      continue
    }
    const parts = run.text.split(/ {2,}/).filter((part) => part !== '')
    let offset = 0
    for (const part of parts) {
      const at = run.text.indexOf(part, offset)
      const prefix = run.text.slice(0, at)
      out.push({ ...run, text: part, x: Math.round((run.x + estimateTextWidth(prefix, run.font_size)) * 100) / 100 })
      offset = at + part.length
    }
  }
  return out
}

/**
 * 估算一段文本的排版宽度（点）。
 *
 * 中日韩字符按 1 em、其余按 0.5 em —— 只用来判断「两个片段之间是词间距还是列间隔」，
 * 不追求与阅读器逐点一致（PDF 的精确宽度在字体宽度表里，嵌入子集字体常常拿不到）。
 *
 * @param {string} text - 文本。
 * @param {number} fontSize - 字号（点）。
 * @returns {number} 估算宽度。
 */
function estimateTextWidth(text, fontSize) {
  const size = Number.isFinite(fontSize) && fontSize > 0 ? fontSize : 10
  let em = 0
  // eslint-disable-next-line no-control-regex
  const cjk = /[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffef]/
  for (const ch of String(text)) em += cjk.test(ch) ? 1 : 0.5
  return em * size
}

/**
 * 拼接同一单元格里的相邻片段。
 *
 * 中日韩文字之间**不加空格**（PDF 常把一个词拆成多个片段，加空格会把「1月」变成「1 月」）；
 * 拉丁字母之间保留一个空格。
 *
 * @param {string} left - 已有文本。
 * @param {string} right - 新片段。
 * @returns {string} 拼接结果。
 */
function joinRuns(left, right) {
  if (left === '') return right
  if (right === '') return left
  // eslint-disable-next-line no-control-regex
  const cjk = /[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffef]/
  const lastChar = left[left.length - 1]
  const firstChar = right[0]
  if (cjk.test(lastChar) || cjk.test(firstChar)) return left + right
  return `${left} ${right}`
}

/**
 * 从内容流文本里读一个字面字符串。
 * @param {string} text - 内容流文本。
 * @param {number} start - `(` 的位置。
 * @returns {{bytes: Buffer, end: number}} 字节与结束位置。
 */
function readLiteralStringFromText(text, start) {
  const out = []
  let depth = 1
  let i = start + 1
  while (i < text.length) {
    const c = text[i]
    if (c === '\\') {
      const next = text[i + 1]
      const map = { n: 10, r: 13, t: 9, b: 8, f: 12 }
      if (map[next] !== undefined) {
        out.push(map[next])
        i += 2
        continue
      }
      if (/[0-7]/.test(next ?? '')) {
        let oct = ''
        let p = i + 1
        while (p < text.length && oct.length < 3 && /[0-7]/.test(text[p])) {
          oct += text[p]
          p += 1
        }
        out.push(Number.parseInt(oct, 8) & 0xff)
        i = p
        continue
      }
      out.push(next.charCodeAt(0))
      i += 2
      continue
    }
    if (c === '(') depth += 1
    if (c === ')') {
      depth -= 1
      if (depth === 0) {
        i += 1
        break
      }
    }
    out.push(c.charCodeAt(0) & 0xff)
    i += 1
  }
  return { bytes: Buffer.from(out), end: i }
}

/**
 * 载入页面字体资源的 ToUnicode 映射。
 *
 * 中文 PDF 的字符串是字形码而非 Unicode，没有这一步中文会提取成乱码。
 *
 * @param {PdfDocument} doc - 文档。
 * @param {object} pageDict - 页面字典。
 * @returns {Map<string, Map<number, string>>} `字体名 → 码点映射`。
 */
function loadToUnicodeMaps(doc, pageDict) {
  const maps = new Map()
  const fonts = doc.deref(doc.deref(pageDict.Resources)?.Font)
  if (!fonts || typeof fonts !== 'object') return maps
  for (const [name, ref] of Object.entries(fonts)) {
    const font = doc.deref(ref)
    if (!font || typeof font !== 'object') continue
    const toUnicode = doc.deref(font.ToUnicode)
    if (!toUnicode || typeof toUnicode !== 'object' || toUnicode.stream === undefined) continue
    try {
      maps.set(name, parseToUnicodeCMap(doc.decodeStream(toUnicode).toString('latin1')))
    } catch {
      // 单个字体映射失败不影响其它字体
    }
  }
  return maps
}

/**
 * 解析 ToUnicode CMap 的 `beginbfchar` / `beginbfrange` 段。
 * @param {string} cmap - CMap 文本。
 * @returns {Map<number, string>} 码点映射。
 */
function parseToUnicodeCMap(cmap) {
  const map = new Map()
  const toCode = (hex) => Number.parseInt(hex, 16)
  const toText = (hex) => {
    let out = ''
    const clean = hex.trim()
    for (let i = 0; i + 3 < clean.length + 1; i += 4) {
      const unit = clean.slice(i, i + 4)
      if (unit.length < 4) break
      out += String.fromCharCode(Number.parseInt(unit, 16))
    }
    return out
  }

  for (const block of cmap.matchAll(/beginbfchar([\s\S]*?)endbfchar/g)) {
    for (const pair of block[1].matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g)) {
      map.set(toCode(pair[1]), toText(pair[2]))
    }
  }
  for (const block of cmap.matchAll(/beginbfrange([\s\S]*?)endbfrange/g)) {
    const body = block[1]
    for (const triple of body.matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g)) {
      const lo = toCode(triple[1])
      const hi = toCode(triple[2])
      const dst = toCode(triple[3])
      for (let code = lo; code <= hi && code - lo < 65536; code += 1) {
        map.set(code, String.fromCharCode(dst + (code - lo)))
      }
    }
    for (const arrayForm of body.matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[([\s\S]*?)\]/g)) {
      const lo = toCode(arrayForm[1])
      const items = [...arrayForm[3].matchAll(/<([0-9A-Fa-f]+)>/g)]
      items.forEach((item, index) => map.set(lo + index, toText(item[1])))
    }
  }
  return map
}

/**
 * 用 ToUnicode 映射解码一段字符串字节。
 *
 * 无映射时按单字节 / 双字节（CID 字体常见）启发式解码，
 * 并在结果中保留可读的 ASCII 部分 —— 总比整段乱码有用。
 *
 * @param {Buffer} bytes - 字符串字节。
 * @param {Map<number, string>|null} map - ToUnicode 映射。
 * @returns {string} 文本。
 */
function decodeWithMap(bytes, map) {
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    let out = ''
    for (let i = 2; i + 1 < bytes.length; i += 2) out += String.fromCharCode(bytes.readUInt16BE(i))
    return out
  }
  if (map && map.size > 0) {
    // 双字节码优先（CID 字体），不行再退回单字节
    const twoByte = []
    let ok = true
    for (let i = 0; i + 1 < bytes.length; i += 2) {
      const code = bytes.readUInt16BE(i)
      if (!map.has(code)) {
        ok = false
        break
      }
      twoByte.push(map.get(code))
    }
    if (ok && twoByte.length > 0 && bytes.length % 2 === 0) return twoByte.join('')
    let out = ''
    for (const b of bytes) out += map.get(b) ?? (b >= 32 && b < 127 ? String.fromCharCode(b) : '')
    return out
  }
  return bytes.toString('latin1')
}

/**
 * 转义正则元字符，让用户查询按**字面量**匹配。
 *
 * 不转义的话，查 `(a)` 或 `1.5` 会被当成正则：轻则匹配到别的内容，
 * 重则像 `[` 这样直接抛 `SyntaxError`。
 *
 * @param {string} text - 原始查询。
 * @returns {string} 可安全嵌入正则的字符串。
 */
function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/**
 * 收集按钮字段的「开状态名」。
 *
 * 这些名字来自 `/AP /N` 的子键（排除 `Off`），**不是固定值**：不同生成器会写 `Yes` / `1` / `开`。
 * 合并式字段的 `/AP` 在自己身上；分离式（如单选组）则在各个 widget 的子字典上，因此要一起看。
 *
 * @param {object} doc - PDF 文档。
 * @param {object} dict - 字段字典。
 * @param {object[]} kidDicts - 子 widget 字典列表。
 * @returns {string[]} 开状态名（去重，保持出现顺序）。
 */
function collectOnStates(doc, dict, kidDicts) {
  const out = []
  const readFrom = (target) => {
    const ap = doc.deref(target.AP)
    if (!ap || typeof ap !== 'object') return
    const normal = doc.deref(ap.N)
    if (!normal || typeof normal !== 'object' || normal.stream) return
    for (const key of Object.keys(normal)) {
      if (key === 'Off') continue
      if (!out.includes(key)) out.push(key)
    }
  }
  readFrom(dict)
  for (const kid of kidDicts) readFrom(kid.dict)
  return out
}

/**
 * 字段类型的中文说明。
 * @param {object} field - 内部字段项。
 * @returns {string} 说明。
 */
function fieldKindLabel(field) {
  if (field.ft === 'Tx') return field.multiline ? '多行文本' : field.password ? '密码文本' : '文本'
  if (field.ft === 'Btn') {
    if ((field.flags & 65536) !== 0) return '按钮（不可填）'
    return field.radio ? '单选组' : '复选框'
  }
  if (field.ft === 'Ch') return field.combo ? (field.editable ? '可编辑下拉' : '下拉') : '列表'
  if (field.ft === 'Sig') return '数字签名（不可填）'
  return '未知类型'
}

/**
 * 四舍五入到两位小数。
 * @param {number} value - 数值。
 * @returns {number} 结果。
 */
function round2(value) {
  return Math.round(value * 100) / 100
}

export { PDF_NOT_CHECKED, decodePdfString, parseValue }
