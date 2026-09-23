/**
 * DOCX 适配器：读取、结构识别与段落级写入。
 *
 * 与 XLSX 共用同一套底座：`ooxml.js` 的 ZIP 容器与最小修改 XML 引擎，
 * 因此本文件只负责 WordprocessingML 语义，不重复任何容器逻辑。
 *
 * 写入遵循「局部 XML 节点修改」原则（§六.4）：插入/改写/删除段落都只触碰
 * 目标 `<w:p>` 的字节区间，样式定义、编号、页眉页脚、图片关系、批注、
 * 修订与域代码一律原样保留，绝不重建整个文档。
 *
 * 对齐开发要求 §六（DOCX 功能规划与开发注意事项）与 §十六.1（validate_docx）。
 *
 * @module dsh-exp-office/docx
 */

import {
  ZipPackage,
  XmlDoc,
  findAll,
  find,
  attr,
  nodeText,
  escapeXmlText,
  escapeXmlAttr,
  ensureContentTypeDefault,
  ensureContentTypeOverride,
  addRelationshipTo,
  EMU_PER_PX,
  detectImageType,
  imageContentType,
  readImageSize,
  resolveImageExtent,
  DEFAULT_ZIP_LIMITS
} from './ooxml.js'

// 保持对外接口不变：这两个函数原在 docx.js 导出，工具层与测试依赖它
export { detectImageType, readImageSize }
import { OfficeError } from './errors.js'

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const CT_DOCUMENT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
const CT_STYLES = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'
const CT_HEADER = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
const CT_FOOTER = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'

/** 一个已打开的 DOCX 文档。 */
export class DocxDocument {
  #pkg

  /**
   * @param {ZipPackage} pkg - 底层 OPC 包。
   * @param {object} info - 文档结构信息。
   */
  constructor(pkg, info) {
    this.#pkg = pkg
    this.info = info
  }

  /**
   * 打开一个 .docx。
   * @param {Buffer} buffer - 文件字节。
   * @param {Partial<typeof DEFAULT_ZIP_LIMITS>} [limits] - ZIP 安全上限。
   * @returns {DocxDocument} 文档对象。
   */
  static open(buffer, limits) {
    if (buffer.length >= 8 && buffer.readUInt32LE(0) === 0xe011cfd0) {
      throw new OfficeError('PASSWORD_REQUIRED', '文件是加密的 OOXML（OLE 复合文档容器），需要密码或解密副本。')
    }
    const pkg = ZipPackage.open(buffer, limits)
    if (!pkg.has('word/document.xml')) {
      throw new OfficeError('CORRUPTED_DOCUMENT', '缺少 word/document.xml，不是有效的 .docx 文件。')
    }
    const names = pkg.names()
    const info = {
      hasMacro: names.some((n) => n.includes('vbaProject')),
      hasSignatures: names.some((n) => n.includes('_xmlsignatures/')),
      headers: names.filter((n) => /^word\/header\d*\.xml$/.test(n)).sort(),
      footers: names.filter((n) => /^word\/footer\d*\.xml$/.test(n)).sort(),
      commentsPart: names.find((n) => /^word\/comments\.xml$/.test(n)) ?? null,
      footnotesPart: names.find((n) => /^word\/footnotes\.xml$/.test(n)) ?? null,
      endnotesPart: names.find((n) => /^word\/endnotes\.xml$/.test(n)) ?? null,
      media: names.filter((n) => n.startsWith('word/media/')),
      embedded: names.filter((n) => n.startsWith('word/embeddings/'))
    }
    return new DocxDocument(pkg, info)
  }

  /**
   * 读取文档元数据（docProps/core.xml 与 app.xml）。
   * @returns {object} 元数据。
   */
  metadata() {
    const result = { title: null, subject: null, creator: null, last_modified_by: null, created: null, modified: null, revision: null }
    if (this.#pkg.has('docProps/core.xml')) {
      const doc = XmlDoc.parse(this.#pkg.readText('docProps/core.xml'))
      const root = doc.root.children.find((c) => c.type === 'element')
      const pick = (name) => {
        const node = find(root, name)
        return node ? nodeText(doc, node) : null
      }
      result.title = pick('title')
      result.subject = pick('subject')
      result.creator = pick('creator')
      result.last_modified_by = pick('lastModifiedBy')
      result.created = pick('created')
      result.modified = pick('modified')
      result.revision = pick('revision')
    }
    if (this.#pkg.has('docProps/app.xml')) {
      const doc = XmlDoc.parse(this.#pkg.readText('docProps/app.xml'))
      const root = doc.root.children.find((c) => c.type === 'element')
      const pick = (name) => {
        const node = find(root, name)
        return node ? nodeText(doc, node) : null
      }
      result.application = pick('Application')
      result.pages = Number(pick('Pages') ?? 0) || null
      result.words = Number(pick('Words') ?? 0) || null
      result.paragraphs_reported = Number(pick('Paragraphs') ?? 0) || null
    }
    return result
  }

  /**
   * 读取正文段落。
   *
   * `w:del` 子树（修订中已删除的文字）不计入正文，`w:ins` 计入 ——
   * 这是 Word 显示修订后的实际可见文本。
   *
   * @param {object} [options] - 选项。
   * @param {number} [options.offset] - 起始下标。
   * @param {number} [options.limit] - 返回数量上限。
   * @returns {object} `{total, offset, returned, paragraphs}`。
   */
  paragraphs({ offset = 0, limit = 500 } = {}) {
    const doc = this.#documentDoc()
    const body = find(doc.root, 'body')
    if (!body) throw new OfficeError('CORRUPTED_DOCUMENT', 'document.xml 缺少 body 节点。')
    const nodes = (body.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'p')
    const all = nodes.map((p, index) => paragraphInfo(doc, p, index))
    const page = all.slice(offset, offset + limit)
    return { total: all.length, offset, returned: page.length, paragraphs: page }
  }

  /**
   * 读取正文表格。
   * @returns {object} `{count, tables}`。
   */
  tables() {
    const doc = this.#documentDoc()
    const body = find(doc.root, 'body')
    const tables = findAll(body ?? doc.root, 'tbl').map((tbl, index) => {
      const rows = findAll(tbl, 'tr').map((tr) =>
        findAll(tr, 'tc').map((tc) => {
          const paragraphs = findAll(tc, 'p').map((p) => paragraphText(doc, p))
          return { text: paragraphs.join('\n'), paragraphs: paragraphs.length }
        })
      )
      const grid = find(tbl, 'tblGrid')
      const columns = grid ? findAll(grid, 'gridCol').length : (rows[0]?.length ?? 0)
      return { index, rows: rows.length, columns, cells: rows }
    })
    return { count: tables.length, tables }
  }

  /**
   * 读取页眉或页脚文本。
   * @param {'header'|'footer'} kind - 类型。
   * @returns {object[]} 每个部件的文本与页码域信息。
   */
  headersOrFooters(kind) {
    const parts = kind === 'header' ? this.info.headers : this.info.footers
    return parts.map((part) => {
      const doc = XmlDoc.parse(this.#pkg.readText(part))
      const texts = findAll(doc.root, 'p').map((p) => paragraphText(doc, p))
      const hasPageField = /PAGE/.test(this.#pkg.readText(part))
      return { part, paragraphs: texts.filter((t) => t !== ''), has_page_number_field: hasPageField }
    })
  }

  /**
   * 结构识别：批注、修订、域代码、书签、内容控件、目录。
   * @returns {object} 结构统计。
   */
  structure() {
    const doc = this.#documentDoc()
    const raw = this.#pkg.readText('word/document.xml')
    // 域代码与页码域常出现在页眉/页脚中，不能只扫正文。
    const peripheral = [...this.info.headers, ...this.info.footers].map((p) => this.#pkg.readText(p)).join('')
    const anywhere = raw + peripheral
    return {
      paragraphs: findAll(doc.root, 'p').length,
      tables: findAll(doc.root, 'tbl').length,
      images: this.info.media.length,
      headers: this.info.headers.length,
      footers: this.info.footers.length,
      has_comments: this.info.commentsPart !== null,
      comments: this.info.commentsPart ? findAll(XmlDoc.parse(this.#pkg.readText(this.info.commentsPart)).root, 'comment').length : 0,
      has_revisions: /<w:(ins|del)[ >]/.test(raw),
      has_fields: /<w:fldChar|<w:instrText|<w:fldSimple/.test(anywhere),
      has_page_number_field: /PAGE/.test(anywhere),
      has_toc: /TOC\s/.test(anywhere),
      has_bookmarks: /<w:bookmarkStart/.test(raw),
      has_content_controls: /<w:sdt[ >]/.test(raw),
      has_hyperlink_field: /HYPERLINK/.test(anywhere),
      footnotes: this.info.footnotesPart ? findAll(XmlDoc.parse(this.#pkg.readText(this.info.footnotesPart)).root, 'footnote').length : 0,
      endnotes: this.info.endnotesPart ? findAll(XmlDoc.parse(this.#pkg.readText(this.info.endnotesPart)).root, 'endnote').length : 0,
      embedded_objects: this.info.embedded.length,
      has_macro: this.info.hasMacro,
      has_digital_signature: this.info.hasSignatures,
      section_properties: findAll(doc.root, 'sectPr').map((s) => ({
        page_width_twips: Number(attr(find(s, 'pgSz') ?? { attrs: new Map() }, 'w') ?? 0) || null,
        page_height_twips: Number(attr(find(s, 'pgSz') ?? { attrs: new Map() }, 'h') ?? 0) || null
      }))
    }
  }

  /**
   * 提取全文纯文本。
   * @param {number} [maxChars] - 字符上限，防止超长文档撑爆上下文。
   * @returns {object} `{text, truncated, length}`。
   */
  text(maxChars = 200000) {
    const doc = this.#documentDoc()
    const body = find(doc.root, 'body')
    if (!body) throw new OfficeError('CORRUPTED_DOCUMENT', 'document.xml 缺少 body 节点。')
    const lines = []
    for (const node of body.children ?? []) {
      if (node.type !== 'element') continue
      const name = localName(node)
      if (name === 'p') lines.push(paragraphText(doc, node))
      else if (name === 'tbl') {
        for (const tr of findAll(node, 'tr')) {
          lines.push(findAll(tr, 'tc').map((tc) => findAll(tc, 'p').map((p) => paragraphText(doc, p)).join(' ')).join('\t'))
        }
      }
    }
    const text = lines.join('\n')
    return { text: text.slice(0, maxChars), truncated: text.length > maxChars, length: text.length }
  }

  /**
   * 读取并缓存 document.xml。
   *
   * 两个标记是**不同**的事，必须分开：
   *   - `#treeStale`：内存里的树落后于补丁，读取前需 rescan；
   *   - `#needsFlush`：包里的 document.xml 落后于树，save() 时必须写回。
   * 早先把两者合成一个标记，结果「写-读-写」序列中一次读取就清掉了脏标记，
   * save() 直接跳过写回，修改静默丢失。
   *
   * @returns {XmlDoc} 文档对象。
   */
  #documentDoc() {
    if (!this.#docCache) {
      this.#docCache = XmlDoc.parse(this.#pkg.readText('word/document.xml'))
      return this.#docCache
    }
    if (this.#treeStale) {
      this.#docCache.rescan()
      this.#treeStale = false
    }
    return this.#docCache
  }

  #docCache = null
  #treeStale = false
  #needsFlush = false

  /**
   * 标记文档已被修改。
   * @returns {void}
   */
  #markDirty() {
    this.#treeStale = true
    this.#needsFlush = true
  }

  /**
   * 取 `<w:body>` 下直接的 `<w:p>` 段落节点列表。
   *
   * 只取 body 的直接子段落：表格单元格里的段落属于表格，不参与段落级编辑，
   * 否则「第 3 段」在含表格的文档里会指向错误的目标。
   *
   * @returns {object[]} 段落节点。
   */
  bodyParagraphs() {
    const doc = this.#documentDoc()
    const body = find(doc.root, 'body')
    if (!body) throw new OfficeError('CORRUPTED_DOCUMENT', 'document.xml 缺少 body 节点。')
    return (body.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'p')
  }

  /**
   * 在指定位置之后插入一个新段落。
   *
   * 只插入一个新的 `<w:p>` 节点，不重建文档，因此样式定义、编号、页眉页脚、
   * 图片关系、批注与域代码全部原样保留。
   *
   * @param {object} args - 参数。
   * @param {number} args.after - 插入到第几段之后（0 基）；传 -1 表示插到最前。
   * @param {string} args.text - 段落文本，`\n` 会转成 `<w:br/>` 换行。
   * @param {string} [args.style] - 段落样式 ID，如 `Heading1`。
   * @param {string} [args.alignment] - 水平对齐：left/center/right/both。
   * @returns {object} 变更信息。
   */
  insertParagraph({ after, text, style, alignment }) {
    const doc = this.#documentDoc()
    const paragraphs = this.bodyParagraphs()
    const index = after + 1
    if (index < 0 || index > paragraphs.length) {
      throw new OfficeError('INVALID_REQUEST', `插入位置 ${index} 超出范围（文档共 ${paragraphs.length} 段）。`, {
        paragraph_count: paragraphs.length
      })
    }
    const xml = buildParagraphXml({ text, style, alignment })
    const anchor = paragraphs[index]
    if (anchor) {
      doc.insertBefore(anchor, xml)
    } else {
      const body = find(doc.root, 'body')
      doc.appendChild(body, xml)
    }
    this.#markDirty()
    return { type: 'insert_paragraph', index, style: style ?? null, text_preview: preview(text) }
  }

  /**
   * 查询一个段落含有哪些会被整段替换破坏的结构化标记。
   *
   * 供调用方在动手前判断哪些段落可以安全改写；返回空数组表示可安全改写。
   *
   * @param {number} index - 段落下标（0 基，仅计 body 直接子段落）。
   * @returns {{index: number, safe: boolean, markup: string[], labels: string[]}} 段落安全性。
   */
  paragraphSafety(index) {
    const paragraphs = this.bodyParagraphs()
    const p = paragraphs[index]
    if (!p) {
      throw new OfficeError('FILE_NOT_FOUND', `不存在第 ${index} 段（文档共 ${paragraphs.length} 段）。`, { paragraph_count: paragraphs.length })
    }
    const markup = this.#protectedMarkup(p)
    return {
      index,
      safe: markup.length === 0,
      markup,
      labels: markup.map((m) => MARKUP_LABELS[m] ?? m)
    }
  }

  /**
   * 改写一个段落的文本，保留其段落属性（样式、对齐、编号等）。
   *
   * **保护性检查**：段落里若含书签、批注锚点、修订标记、域代码或超链接，
   * 替换 run 内容会把它们一并抹掉 —— 部件本身还在，但引用消失，Word 会
   * 静默丢弃对应功能（实测：改一段会让批注从 1 条变成 0 条）。
   * 因此默认拒绝，必须显式传 `allowMarkupLoss` 才继续。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 段落下标（0 基，仅计 body 直接子段落）。
   * @param {string} args.text - 新文本；空字符串表示清空该段内容。
   * @param {boolean} [args.allowMarkupLoss] - 明知会丢失结构化标记仍继续。
   * @returns {object} 变更信息。
   */
  updateParagraph({ index, text, allowMarkupLoss = false }) {
    const doc = this.#documentDoc()
    const paragraphs = this.bodyParagraphs()
    const p = paragraphs[index]
    if (!p) {
      throw new OfficeError('FILE_NOT_FOUND', `不存在第 ${index} 段（文档共 ${paragraphs.length} 段）。`, { paragraph_count: paragraphs.length })
    }
    this.#assertSafeToReplace(p, index, '改写', allowMarkupLoss)

    const previous = paragraphText(doc, p)
    const pPr = find(p, 'pPr')
    // 保留 <w:pPr>（段落属性，含样式与编号），只替换其后的 run 内容。
    const from = pPr ? pPr.end : p.startTagEnd
    doc.patch(from, p.contentEnd, runsXml(text))
    this.#markDirty()
    return { type: 'update_paragraph', index, from: preview(previous), to: preview(text) }
  }

  /**
   * 段落内一旦被整段替换就会失效的结构化标记。
   *
   * 这些不是「内容」，而是被别处引用的锚点：批注靠 commentReference 关联，
   * 书签靠 bookmarkStart/End 定界，目录靠域代码。删掉锚点，
   * 对应的部件会变成无人引用的孤儿，Word 会直接不显示它。
   */
  static PROTECTED_MARKUP = Object.freeze([
    'bookmarkStart',
    'bookmarkEnd',
    'commentRangeStart',
    'commentRangeEnd',
    'commentReference',
    'ins',
    'del',
    'fldChar',
    'instrText',
    'fldSimple',
    'hyperlink',
    'sdt',
    'footnoteReference',
    'endnoteReference'
  ])

  /**
   * 检查段落是否含有会被整段替换破坏的结构化标记。
   * @param {object} p - 段落节点。
   * @returns {string[]} 命中的标记本地名。
   */
  #protectedMarkup(p) {
    const found = new Set()
    const walk = (node) => {
      for (const child of node.children ?? []) {
        if (child.type !== 'element') continue
        const name = localName(child)
        if (DocxDocument.PROTECTED_MARKUP.includes(name)) found.add(name)
        walk(child)
      }
    }
    walk(p)
    return [...found]
  }

  /**
   * 在整段替换/删除前断言安全，否则抛出可操作的错误。
   * @param {object} p - 段落节点。
   * @param {number} index - 段落下标。
   * @param {string} action - 动作名（改写/删除）。
   * @param {boolean} allowed - 调用方是否已显式接受损失。
   * @returns {void}
   */
  #assertSafeToReplace(p, index, action, allowed) {
    if (allowed) return
    const markup = this.#protectedMarkup(p)
    if (markup.length === 0) return
    const readable = markup.map((m) => MARKUP_LABELS[m] ?? m).join('、')
    throw new OfficeError(
      'UNSUPPORTED_FEATURE',
      `第 ${index} 段含${readable}，${action}该段会使其失效（Word 会静默丢弃对应功能）。`,
      { index, markup },
      {
        solution:
          `如确实要${action}，请显式传 allow_markup_loss=true 接受该损失；` +
          '或改为插入新段落、或选择不含这些标记的段落。',
        needsConfirmation: true
      }
    )
  }

  /**
   * 删除一个段落。
   *
   * 与改写同样做保护性检查：删除含批注锚点或书签的段落会孤立它们。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 段落下标（0 基）。
   * @param {boolean} [args.allowMarkupLoss] - 明知会丢失结构化标记仍继续。
   * @returns {object} 变更信息。
   */
  deleteParagraph({ index, allowMarkupLoss = false }) {
    const doc = this.#documentDoc()
    const paragraphs = this.bodyParagraphs()
    const p = paragraphs[index]
    if (!p) {
      throw new OfficeError('FILE_NOT_FOUND', `不存在第 ${index} 段（文档共 ${paragraphs.length} 段）。`, { paragraph_count: paragraphs.length })
    }
    if (paragraphs.length === 1) {
      throw new OfficeError('INVALID_REQUEST', '文档只剩一个段落，删除后将不再是有效文档。')
    }
    this.#assertSafeToReplace(p, index, '删除', allowMarkupLoss)
    const removed = paragraphText(doc, p)
    doc.remove(p)
    this.#markDirty()
    return { type: 'delete_paragraph', index, removed_preview: preview(removed) }
  }

  /**
   * 查找替换正文中的文本。
   *
   * 替换发生在 `<w:t>` 节点内，因此**跨 run 断开的文本不会被命中** —— 这是刻意的：
   * 跨 run 替换要重建 run 结构，会破坏原有的字符格式边界（例如「**加粗**的部分」
   * 被拆成三个 run 时无法安全合并）。当某段合并后的文本含目标却没有单个
   * `<w:t>` 命中时，会在 `split_run_warnings` 中明确报告而不静默跳过。
   *
   * @param {object} args - 参数。
   * @param {string} args.find - 查找内容。
   * @param {string} args.replace - 替换内容。
   * @param {boolean} [args.matchCase] - 是否区分大小写。
   * @param {boolean} [args.includeHeaders] - 是否同时替换页眉页脚。
   * @returns {object} 变更统计。
   */
  findAndReplace({ find: needle, replace: replacement, matchCase = false, includeHeaders = false }) {
    if (typeof needle !== 'string' || needle === '') {
      throw new OfficeError('INVALID_REQUEST', '查找内容不能为空。')
    }
    const parts = ['word/document.xml']
    if (includeHeaders) parts.push(...this.info.headers, ...this.info.footers)

    let replacements = 0
    const changes = []
    const splitRunWarnings = []
    for (const part of parts) {
      const isMain = part === 'word/document.xml'
      const doc = isMain ? this.#documentDoc() : XmlDoc.parse(this.#pkg.readText(part))
      for (const p of findAll(doc.root, 'p')) {
        const paragraphBefore = paragraphText(doc, p)
        if (!contains(paragraphBefore, needle, matchCase)) continue
        let hitInParagraph = 0
        for (const t of findAll(p, 't')) {
          const original = nodeText(doc, t)
          const updated = replaceAll(original, needle, replacement, matchCase)
          if (updated === original) continue
          doc.setText(t, updated)
          hitInParagraph += 1
          replacements += 1
          changes.push({ part, from: preview(original), to: preview(updated) })
        }
        if (hitInParagraph === 0) {
          splitRunWarnings.push({
            code: 'TEXT_SPLIT_ACROSS_RUNS',
            message: `段落「${preview(paragraphBefore, 40)}」中匹配到的文本被拆散在多个格式片段里，未替换；请人工处理或调整模板。`
          })
        }
      }
      if (isMain) this.#markDirty()
      else this.#pkg.write(part, doc.toString())
    }
    return { replacements, changes, split_run_warnings: splitRunWarnings }
  }

  /**
   * 序列化文档。
   * @returns {Buffer} .docx 字节。
   */
  save() {
    if (this.#docCache && this.#needsFlush) {
      this.#pkg.write('word/document.xml', this.#docCache.toString())
      this.#docCache.rescan()
      this.#treeStale = false
      this.#needsFlush = false
    }
    return this.#pkg.toBuffer()
  }

  /**
   * 对文档做结构校验（开发要求 §十六.1 的 validate_docx 检查项）。
   *
   * 只做**可静态判定**的检查：包结构、XML 合法性、关系有效性、样式/编号引用完整性、
   * 页眉页脚引用、目录与域代码、批注与修订是否仍在。
   *
   * 明确不做、并如实列在 `not_checked` 里的项：空白页面、元素重叠、表格越界、
   * 页眉页脚重叠、字体替换 —— 这些必须渲染成图片才能判定，属于视觉回归范围，
   * 不能靠读 XML 猜测（§十八.6：不得为了让测试通过而忽略真实的格式损失）。
   *
   * @returns {object} `{valid, checks, not_checked, fonts, rel_targets}`。
   */
  validate() {
    const checks = []
    const pkg = this.#pkg
    const names = new Set(pkg.names())

    checks.push({ name: 'ZIP 包结构与中央目录', ok: true, detail: `${names.size} 个部件` })

    const required = ['[Content_Types].xml', '_rels/.rels', 'word/document.xml']
    const missing = required.filter((p) => !names.has(p))
    checks.push({ name: '必需部件齐全', ok: missing.length === 0, detail: missing.length ? `缺少 ${missing.join(', ')}` : '全部存在' })

    let doc = null
    try {
      doc = this.#documentDoc()
      checks.push({ name: 'document.xml 可解析', ok: true })
    } catch (err) {
      checks.push({ name: 'document.xml 可解析', ok: false, detail: err.message })
      return { valid: false, checks, not_checked: NOT_CHECKED, fonts: [], rel_targets: [] }
    }

    const body = find(doc.root, 'body')
    if (!body) {
      checks.push({ name: '正文 body 存在', ok: false, detail: 'document.xml 缺少 body' })
      return { valid: false, checks, not_checked: NOT_CHECKED, fonts: [], rel_targets: [] }
    }
    checks.push({ name: '正文存在', ok: true, detail: `${this.bodyParagraphs().length} 个正文段落` })

    // 页眉页脚部件可解析
    const badParts = []
    for (const part of [...this.info.headers, ...this.info.footers]) {
      try {
        XmlDoc.parse(pkg.readText(part))
      } catch (err) {
        badParts.push(`${part}: ${err.message}`)
      }
    }
    checks.push({ name: '页眉页脚 XML 合法', ok: badParts.length === 0, detail: badParts.length ? badParts.join('; ') : `${this.info.headers.length} 页眉 / ${this.info.footers.length} 页脚` })

    // 关系有效性
    const rels = readRelationshipTargets(pkg, 'word/_rels/document.xml.rels')
    const missingTargets = []
    for (const [id, target] of rels) {
      const external = target.startsWith('http://') || target.startsWith('https://') || target.startsWith('mailto:')
      if (external) continue
      const resolved = resolvePart('word', target)
      if (!names.has(resolved)) missingTargets.push(`${id} → ${target}`)
    }
    checks.push({
      name: '关系目标部件存在',
      ok: missingTargets.length === 0,
      detail: missingTargets.length ? `悬空关系：${missingTargets.join(', ')}` : `${rels.size} 条关系全部有效`
    })

    // 文档中引用的 r:id / r:embed 是否都有对应关系
    const referenced = new Set()
    for (const node of findAll(doc.root, 'blip')) {
      const embed = attr(node, 'embed')
      if (embed) referenced.add(embed)
    }
    for (const node of findAll(doc.root, 'hyperlink')) {
      const id = attr(node, 'id')
      if (id) referenced.add(id)
    }
    for (const node of findAll(doc.root, 'headerReference')) {
      const id = attr(node, 'id')
      if (id) referenced.add(id)
    }
    for (const node of findAll(doc.root, 'footerReference')) {
      const id = attr(node, 'id')
      if (id) referenced.add(id)
    }
    const dangling = [...referenced].filter((id) => !rels.has(id))
    checks.push({
      name: '图片/超链接/页眉页脚引用可解析',
      ok: dangling.length === 0,
      detail: dangling.length ? `无对应关系的引用：${dangling.join(', ')}` : `${referenced.size} 个引用全部可解析`
    })

    // 样式引用完整性（pStyle / rStyle 指向的样式必须存在）
    const styleIds = readStyleIds(pkg)
    const styleRefs = new Set()
    for (const tag of ['pStyle', 'rStyle', 'tblStyle']) {
      for (const node of findAll(doc.root, tag)) {
        const val = attr(node, 'val')
        if (val) styleRefs.add(val)
      }
    }
    const undefinedStyles = styleIds === null ? [] : [...styleRefs].filter((s) => !styleIds.has(s))
    checks.push({
      name: '样式引用有定义',
      ok: undefinedStyles.length === 0,
      detail:
        styleIds === null
          ? '文档没有 styles.xml，跳过'
          : undefinedStyles.length
            ? `引用了未定义的样式：${undefinedStyles.join(', ')}`
            : `${styleRefs.size} 个样式引用全部有定义`
    })

    // 编号引用完整性（numId 必须存在于 numbering.xml）
    const numIds = readNumberingIds(pkg)
    const usedNumIds = new Set(findAll(doc.root, 'numId').map((n) => attr(n, 'val')).filter(Boolean))
    const undefinedNums = numIds === null ? [] : [...usedNumIds].filter((n) => !numIds.has(n))
    checks.push({
      name: '编号引用有定义',
      ok: undefinedNums.length === 0,
      detail:
        numIds === null
          ? usedNumIds.size > 0
            ? `⚠️ 文档使用了 ${usedNumIds.size} 个编号，但缺少 numbering.xml（编号会断裂）`
            : '文档未使用编号'
          : undefinedNums.length
            ? `引用了未定义的 numId：${undefinedNums.join(', ')}`
            : `${usedNumIds.size} 个编号引用全部有定义`,
      warning: numIds === null && usedNumIds.size > 0
    })

    // 字体清单（不做「缺失」判定：跨平台系统字体不可比，交由调用方判断）
    const fonts = new Set()
    for (const node of findAll(doc.root, 'rFonts')) {
      for (const key of ['ascii', 'eastAsia', 'hAnsi', 'cs']) {
        const val = attr(node, key)
        if (val) fonts.add(val)
      }
    }

    // 域代码与目录（未更新会导致显示陈旧）
    const rawAll = [pkg.readText('word/document.xml'), ...[...this.info.headers, ...this.info.footers].map((p) => pkg.readText(p))].join('')
    checks.push({
      name: '域代码状态',
      ok: true,
      detail: /<w:fldChar|<w:instrText|<w:fldSimple/.test(rawAll) ? '含域代码，插件不会更新域；请在 Word 中按 F9 更新' : '无域代码',
      warning: /<w:fldChar|<w:instrText|<w:fldSimple/.test(rawAll)
    })
    checks.push({
      name: '目录状态',
      ok: true,
      detail: /TOC\s/.test(rawAll) ? '含目录域，改动段落后需在 Word 中更新目录' : '无目录',
      warning: /TOC\s/.test(rawAll)
    })

    // 评估：批注与修订仍在
    const structure = this.structure()
    checks.push({
      name: '批注保留',
      ok: structure.has_comments ? structure.comments > 0 : true,
      detail: structure.has_comments ? `${structure.comments} 条批注` : '无批注'
    })
    checks.push({
      name: '修订保留',
      ok: true,
      detail: structure.has_revisions ? '含修订标记，未被删除' : '无修订'
    })

    return {
      valid: checks.every((c) => c.ok),
      checks,
      not_checked: NOT_CHECKED,
      fonts: [...fonts],
      rel_targets: [...rels.entries()].map(([id, target]) => ({ id, target }))
    }
  }

  /**
   * 设置页眉或页脚的文本，可选在末尾附加 PAGE 页码域。
   *
   * 只改写目标部件里第一个 `<w:p>` 的 run 内容，保留其段落属性；
   * 页眉页脚部件本身的关系、内容类型与 sectPr 引用一律不动。
   *
   * 当前版本只支持**修改已存在的**页眉页脚部件。文档没有该部件时返回
   * `UNSUPPORTED_FEATURE` 并给出可操作的替代方案，而不是悄悄新建一个
   * 可能不符合用户预期的部件。
   *
   * @param {object} args - 参数。
   * @param {'header'|'footer'} args.kind - 类型。
   * @param {string} args.text - 文本内容。
   * @param {boolean} [args.pageNumber] - 是否在文本后附加页码域。
   * @param {number} [args.index] - 第几个页眉/页脚部件（默认 0，即默认页眉页脚）。
   * @returns {object} 变更信息。
   */
  setHeaderFooter({ kind, text, pageNumber = false, index = 0 }) {
    const listKey = kind === 'header' ? 'headers' : 'footers'
    const existedBefore = Boolean(this.info[listKey][index])
    const part = this.#ensureHeaderFooterPart(kind, index)

    const doc = XmlDoc.parse(this.#pkg.readText(part))
    const p = find(doc.root, 'p')
    if (!p) throw new OfficeError('CORRUPTED_DOCUMENT', `${part} 中没有段落。`)
    const previous = paragraphText(doc, p)
    const pPr = find(p, 'pPr')
    const from = pPr ? pPr.end : p.startTagEnd
    doc.patch(from, p.contentEnd, runsXml(text ?? '') + (pageNumber ? PAGE_FIELD_XML : ''))
    this.#pkg.write(part, doc.toString())
    return {
      type: pageNumber ? `set_${kind}_with_page_number` : `set_${kind}`,
      part,
      created: !existedBefore,
      from: preview(previous),
      to: preview(text ?? ''),
      page_number_field: pageNumber
    }
  }

  /**
   * 读取批注（`word/comments.xml`）。
   *
   * 同时报告 Word 2013+ 的**线程化批注**附加部件是否在场（`commentsExtended` /
   * `commentsIds` / `commentsExtensible` / `people`）：批注的「回复关系、已解决状态、
   * 人员标识」都存在那里面，本工具只读主部件，因此会如实说明这些信息未解析，
   * 而不是假装批注没有更多字段。
   *
   * @returns {object} 批注列表与部件情况。
   */
  comments() {
    const part = this.info.commentsPart
    if (!part) return { part: null, count: 0, comments: [], extended_parts: [], note: '文档没有批注部件。' }
    const doc = XmlDoc.parse(this.#pkg.readText(part))
    const paragraphs = this.bodyParagraphs()
    // 批注锚点：正文里 <w:commentReference w:id> 所在段落
    const docXml = this.#documentDoc()
    const anchors = new Map()
    for (const ref of findAll(docXml.root, 'commentReference')) {
      const id = attr(ref, 'id')
      if (id === undefined || anchors.has(id)) continue
      const paragraph = enclosingParagraph(ref)
      anchors.set(id, paragraph ? paragraphs.indexOf(paragraph) : -1)
    }
    const comments = findAll(doc.root, 'comment').map((node) => ({
      id: Number(attr(node, 'id') ?? -1),
      author: attr(node, 'author') ?? null,
      initials: attr(node, 'initials') ?? null,
      date: attr(node, 'date') ?? null,
      text: findAll(node, 't')
        .map((t) => t.children?.map((c) => c.raw ?? '').join('') ?? '')
        .join(''),
      anchored_paragraph: anchors.get(attr(node, 'id')) ?? null
    }))
    const extended = ['commentsExtended', 'commentsIds', 'commentsExtensible', 'people']
      .map((name) => `word/${name}.xml`)
      .filter((p) => this.#pkg.has(p))
    return {
      part,
      count: comments.length,
      comments,
      extended_parts: extended,
      note: extended.length
        ? `含 ${extended.length} 个线程化批注附加部件（回复关系/已解决状态/人员），本工具只读主批注部件，未解析这些字段。`
        : '无附加批注部件（Word 2010 式批注）。'
    }
  }

  /**
   * 读取修订（`w:ins` / `w:del` / `w:moveFrom` / `w:moveTo` 与各类 `*Change`）。
   *
   * 与读取段落时的语义保持一致：**已插入**的文字计入正文、**已删除**的不计入，
   * 因此这里给出的 `text` 是各自的真实内容（删除的取 `w:delText`，
   * 不然会读到空串 —— 删除的文字不在 `w:t` 里）。
   *
   * @returns {object} 修订列表与统计。
   */
  revisions() {
    const doc = this.#documentDoc()
    const rows = []
    const push = (node, type) => {
      const tag = localName(node)
      const parts = tag === 'del' || tag === 'moveFrom' ? ['delText'] : ['t']
      const text = parts
        .flatMap((name) => findAll(node, name))
        .map((t) => t.children?.map((c) => c.raw ?? '').join('') ?? '')
        .join('')
      const paragraph = enclosingParagraph(node)
      // `w:ins` / `w:del` 出现在 `w:pPr/w:rPr` 里时表示**段落标记**本身被增删
      // （即「与下一段合并/拆分」），它天然没有文字，不能和「删掉一段文字」混为一谈。
      const parentName = node.parent ? localName(node.parent) : ''
      const paragraphMark = parentName === 'rPr' || parentName === 'pPr'
      rows.push({
        type: paragraphMark ? `${CHANGE_TYPES[tag] ?? tag}（段落标记）` : (CHANGE_TYPES[tag] ?? tag),
        author: attr(node, 'author') ?? null,
        date: attr(node, 'date') ?? null,
        id: attr(node, 'id') ?? null,
        text,
        paragraph_mark: paragraphMark,
        paragraph: paragraph ? this.bodyParagraphs().indexOf(paragraph) : -1
      })
    }
    for (const tag of ['ins', 'del', 'moveFrom', 'moveTo']) {
      for (const node of findAll(doc.root, tag)) push(node, tag)
    }
    const formatting = []
    for (const tag of ['rPrChange', 'pPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange', 'sectPrChange']) {
      for (const node of findAll(doc.root, tag)) {
        const paragraph = enclosingParagraph(node)
        formatting.push({
          type: CHANGE_TYPES[tag] ?? tag,
          author: attr(node, 'author') ?? null,
          date: attr(node, 'date') ?? null,
          paragraph: paragraph ? this.bodyParagraphs().indexOf(paragraph) : -1
        })
      }
    }
    rows.sort((a, b) => (a.paragraph - b.paragraph) || String(a.date ?? '').localeCompare(String(b.date ?? '')))
    return {
      part: 'word/document.xml',
      count: rows.length,
      formatting_count: formatting.length,
      by_type: rows.reduce((acc, r) => ({ ...acc, [r.type]: (acc[r.type] ?? 0) + 1 }), {}),
      revisions: rows,
      formatting_changes: formatting,
      note: '插件不修改也不接受/拒绝修订；本视图只读。'
    }
  }

  /**
   * 设置页面尺寸与方向（末节）。
   *
   * 页面尺寸写在**末节**的 `<w:sectPr><w:pgSz/></w:sectPr>` 上：只有 body 的直接子
   * `sectPr` 才代表最后一节；段落内部的 `sectPr` 属分节符，改它只影响前面那一节。
   * 文档没有 `sectPr` 时补一个带该子元素的节属性（Word 接受，表现为默认节）。
   *
   * 尺寸以 twips（1/20 磅）存储：1 厘米 = 566.93 twips，1 英寸 = 1440 twips。
   * 方向通过交换宽高实现，横向时写 `w:orient="landscape"`（纵向是默认值，不写）。
   *
   * @param {object} args - 参数。
   * @param {string} [args.paper] - 纸张预设：A4 / A3 / A5 / Letter / Legal。
   * @param {'portrait'|'landscape'} [args.orientation] - 方向。
   * @param {number} [args.widthCm] - 自定义页宽（厘米）；给定时覆盖纸张预设。
   * @param {number} [args.heightCm] - 自定义页高（厘米）。
   * @returns {object} 变更信息。
   */
  setPageLayout({ paper, orientation, widthCm, heightCm }) {
    if (paper !== undefined && !PAPER_SIZES[paper]) {
      throw new OfficeError('INVALID_REQUEST', `未知纸张「${paper}」，支持 ${Object.keys(PAPER_SIZES).join(' / ')}。`)
    }
    if (orientation !== undefined && orientation !== 'portrait' && orientation !== 'landscape') {
      throw new OfficeError('INVALID_REQUEST', `orientation 只支持 portrait / landscape，实际「${orientation}」。`)
    }
    for (const [name, value] of [['widthCm', widthCm], ['heightCm', heightCm]]) {
      if (value !== undefined && (!Number.isFinite(value) || value <= 0)) {
        throw new OfficeError('INVALID_REQUEST', `${name} 必须是正数（厘米），实际 ${value}。`)
      }
    }

    const doc = this.#documentDoc()
    const sectPr = lastBodySectPr(doc)
    const current = sectPr ? find(sectPr, 'pgSz') : null
    const before = current
      ? {
          width_twips: Number(attr(current, 'w') ?? 0) || null,
          height_twips: Number(attr(current, 'h') ?? 0) || null,
          orientation: attr(current, 'orient') === 'landscape' ? 'landscape' : 'portrait'
        }
      : null

    let width = before?.width_twips ?? 0
    let height = before?.height_twips ?? 0
    if (paper) {
      width = PAPER_SIZES[paper].width
      height = PAPER_SIZES[paper].height
    }
    if (widthCm !== undefined) width = cmToTwips(widthCm)
    if (heightCm !== undefined) height = cmToTwips(heightCm)
    if (!width || !height) {
      width = PAPER_SIZES.A4.width
      height = PAPER_SIZES.A4.height
    }

    let landscape = width > height
    if (orientation !== undefined) landscape = orientation === 'landscape'
    if (landscape !== width > height) {
      const swap = width
      width = height
      height = swap
    }

    const xml = `<w:pgSz w:w="${width}" w:h="${height}"${landscape ? ' w:orient="landscape"' : ''}/>`
    this.#applySectionChild(doc, 'pgSz', xml)
    this.#markDirty()

    return {
      type: 'set_page_layout',
      part: 'word/document.xml',
      section: 'last',
      from: before,
      to: {
        width_twips: width,
        height_twips: height,
        width_cm: twipsToCm(width),
        height_cm: twipsToCm(height),
        orientation: landscape ? 'landscape' : 'portrait',
        paper: paper ?? null
      }
    }
  }

  /**
   * 设置末节的页边距。
   *
   * 文档已有 `pgMar` 时**逐项改写**（只覆盖传入的项，其余保持原值）；
   * 没有时新建一个，未传入的项按 Word 简体中文默认值补齐
   * （上下 2.54 cm、左右 3.17 cm、页眉页脚 1.5 cm、装订线 0）——
   * 缺属性的 `pgMar` 会让不同阅读器各自取默认值，结果不一致。
   *
   * @param {object} args - 参数。
   * @param {number} [args.topCm] - 上边距（厘米）。
   * @param {number} [args.bottomCm] - 下边距（厘米）。
   * @param {number} [args.leftCm] - 左边距（厘米）。
   * @param {number} [args.rightCm] - 右边距（厘米）。
   * @param {number} [args.headerCm] - 页眉距边界（厘米）。
   * @param {number} [args.footerCm] - 页脚距边界（厘米）。
   * @param {number} [args.gutterCm] - 装订线（厘米）。
   * @returns {object} 变更信息。
   */
  setMargins({ topCm, bottomCm, leftCm, rightCm, headerCm, footerCm, gutterCm }) {
    const given = { top: topCm, right: rightCm, bottom: bottomCm, left: leftCm, header: headerCm, footer: footerCm, gutter: gutterCm }
    if (Object.values(given).every((v) => v === undefined)) {
      throw new OfficeError(
        'INVALID_REQUEST',
        '至少要给出一个页边距参数（top_cm / bottom_cm / left_cm / right_cm / header_cm / footer_cm / gutter_cm）。'
      )
    }
    for (const [name, value] of Object.entries(given)) {
      if (value !== undefined && (!Number.isFinite(value) || value < 0)) {
        throw new OfficeError('INVALID_REQUEST', `${name}_cm 必须是非负数（厘米），实际 ${value}。`)
      }
    }

    const doc = this.#documentDoc()
    const sectPr = lastBodySectPr(doc)
    const current = sectPr ? find(sectPr, 'pgMar') : null
    const values = {}
    for (const key of Object.keys(MARGIN_DEFAULTS)) {
      const existing = current ? Number(attr(current, key) ?? NaN) : NaN
      const base = Number.isFinite(existing) ? existing : MARGIN_DEFAULTS[key]
      values[key] = given[key] === undefined ? base : cmToTwips(given[key])
    }

    const pgSz = sectPr ? find(sectPr, 'pgSz') : null
    const pageWidth = Number(attr(pgSz ?? { attrs: new Map() }, 'w') ?? 0) || PAPER_SIZES.A4.width
    const pageHeight = Number(attr(pgSz ?? { attrs: new Map() }, 'h') ?? 0) || PAPER_SIZES.A4.height
    if (values.left + values.right >= pageWidth || values.top + values.bottom >= pageHeight) {
      throw new OfficeError(
        'INVALID_REQUEST',
        `页边距过大：左右合计 ${twipsToCm(values.left + values.right)} cm，页宽 ${twipsToCm(pageWidth)} cm。`
      )
    }

    const attrs = Object.keys(MARGIN_DEFAULTS)
      .map((k) => ` w:${k}="${values[k]}"`)
      .join('')
    this.#applySectionChild(doc, 'pgMar', `<w:pgMar${attrs}/>`)
    this.#markDirty()

    const to = {}
    for (const key of Object.keys(values)) to[`${key}_cm`] = twipsToCm(values[key])
    return { type: 'set_margins', part: 'word/document.xml', section: 'last', updated: current ? 'in_place' : 'created', to }
  }

  /**
   * 插入页码域。
   *
   * 页码是**域**（`PAGE` + fldChar 三段式），由 Word 按页自动更新，插件不写死数字。
   * 复用默认页脚/页眉部件（没有就新建并接好关系与内容类型），
   * 把该部件的第一个段落改写成「前缀 + 域 + 后缀」。
   *
   * @param {object} args - 参数。
   * @param {'footer'|'header'} [args.position] - 放在页脚（默认）还是页眉。
   * @param {'left'|'center'|'right'} [args.align] - 水平对齐，默认 center。
   * @param {string} [args.prefix] - 页码前文本，如 `第 `。
   * @param {string} [args.suffix] - 页码后文本，如 ` 页`。
   * @returns {object} 变更信息。
   */
  insertPageNumber({ position = 'footer', align = 'center', prefix = '', suffix = '' }) {
    if (position !== 'footer' && position !== 'header') {
      throw new OfficeError('INVALID_REQUEST', `position 只支持 footer / header，实际「${position}」。`)
    }
    if (!ALIGN_VALUES[align]) throw new OfficeError('INVALID_REQUEST', `align 只支持 left / center / right，实际「${align}」。`)

    const listKey = position === 'header' ? 'headers' : 'footers'
    const existedBefore = Boolean(this.info[listKey][0])
    const part = this.#ensureHeaderFooterPart(position, 0)
    const doc = XmlDoc.parse(this.#pkg.readText(part))
    const p = find(doc.root, 'p')
    if (!p) throw new OfficeError('CORRUPTED_DOCUMENT', `${part} 中没有段落。`)
    const previous = paragraphText(doc, p)
    const pPr = find(p, 'pPr')
    const body = runsXml(prefix ?? '') + PAGE_FIELD_XML + runsXml(suffix ?? '')
    if (pPr) {
      // 段落属性已存在：只在它内部改 jc，正文整段替换（两个补丁区间不重叠）
      doc.patch(pPr.end, p.contentEnd, body)
      setAlignmentInPPr(doc, pPr, align)
    } else {
      // 段落属性不存在：**一次**补丁同时写 pPr 与正文。
      // 不能分两次（先替换正文、再在正文起点插入 pPr）：同一起点上的零长插入
      // 排在替换补丁之后，结果正文会被写两遍——旧内容留在段尾。
      doc.patch(p.startTagEnd, p.contentEnd, `<w:pPr>${paragraphAlignXml(align)}</w:pPr>${body}`)
    }
    this.#pkg.write(part, doc.toString())

    return {
      type: 'insert_page_number',
      part,
      position,
      created: !existedBefore,
      align,
      prefix: prefix ?? '',
      suffix: suffix ?? '',
      field: 'PAGE',
      from: preview(previous)
    }
  }

  /**
   * 把某个子元素写到末节的 `sectPr` 上：已有同名子元素就地替换，没有就按 schema 顺序插入，
   * 整篇文档没有 sectPr 时补一个只含该子元素的节属性。
   *
   * @param {object} doc - 文档。
   * @param {string} name - 子元素本地名。
   * @param {string} xml - 子元素 XML。
   * @returns {void}
   */
  #applySectionChild(doc, name, xml) {
    const body = find(doc.root, 'body')
    if (!body) throw new OfficeError('CORRUPTED_DOCUMENT', 'document.xml 缺少 body 节点。')
    const sections = (body.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'sectPr')
    const sectPr = sections.at(-1)
    if (!sectPr) {
      doc.appendChild(body, `<w:sectPr>${xml}</w:sectPr>`)
      return
    }
    const existing = find(sectPr, name)
    if (existing) {
      doc.patch(existing.start, existing.end, xml)
      return
    }
    ensureOrderedChild(doc, sectPr, name, xml, SECTPR_ORDER)
  }

  /**
   * 确保存在可写的页眉/页脚部件，不存在则新建并接线。
   *
   * 新建一个页眉需要补齐四处，缺一处 Word 就会忽略它或判定文档损坏：
   *   1. `word/headerN.xml` 部件本身
   *   2. `[Content_Types].xml` 的 Override
   *   3. `word/_rels/document.xml.rels` 的 header 关系
   *   4. 最后一个 `<w:sectPr>` 里的 `<w:headerReference>`（决定它真的生效）
   *
   * @param {'header'|'footer'} kind - 类型。
   * @param {number} index - 目标部件下标；等于现有数量时新建。
   * @returns {string} 可写的部件路径。
   */
  #ensureHeaderFooterPart(kind, index) {
    const key = kind === 'header' ? 'headers' : 'footers'
    const list = this.info[key]
    if (list[index]) return list[index]
    if (index !== list.length) {
      throw new OfficeError('INVALID_REQUEST', `只有 ${list.length} 个${kind === 'header' ? '页眉' : '页脚'}部件，无法直接创建第 ${index} 个。`)
    }

    let n = 1
    while (this.#pkg.has(`word/${kind}${n}.xml`)) n += 1
    const part = `word/${kind}${n}.xml`
    const rootTag = kind === 'header' ? 'hdr' : 'ftr'
    this.#pkg.write(
      part,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:${rootTag} xmlns:w="${NS_W}"><w:p></w:p></w:${rootTag}>`
    )
    ensureContentTypeOverride(this.#pkg, `/${part}`, kind === 'header' ? CT_HEADER : CT_FOOTER)
    const rid = addRelationshipTo(this.#pkg, 'word/_rels/document.xml.rels', `${NS_R}/${kind}`, part.replace(/^word\//, ''))
    this.#attachSectionReference(kind, rid)
    list.push(part)
    return part
  }

  /**
   * 把页眉页脚引用写进最后一个 `<w:sectPr>`。
   *
   * 没有 sectPr 时补一个：节属性是页眉页脚引用的唯一挂载点，
   * 部件写了但没挂上去，Word 会当作未被引用的孤儿部件而忽略。
   *
   * @param {'header'|'footer'} kind - 类型。
   * @param {string} rid - 关系 ID。
   * @returns {void}
   */
  #attachSectionReference(kind, rid) {
    const doc = this.#documentDoc()
    const body = find(doc.root, 'body')
    if (!body) throw new OfficeError('CORRUPTED_DOCUMENT', 'document.xml 缺少 body 节点。')

    // 末节的 sectPr 是 body 的直接子元素；段落内的 sectPr 属分节符，不动它
    const bodySections = (body.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'sectPr')
    const sectPr = bodySections.at(-1)
    const xml = `<w:${kind}Reference w:type="default" r:id="${rid}"/>`

    if (sectPr) {
      // CT_SectPr 有顺序约束：headerReference 必须排在 footerReference 与 pgSz 之前
      const footerRef = kind === 'header' ? find(sectPr, 'footerReference') : null
      const pgSz = find(sectPr, 'pgSz')
      const anchor = footerRef ?? pgSz
      if (anchor) doc.insertBefore(anchor, xml)
      else doc.appendChild(sectPr, xml)
    } else {
      doc.appendChild(body, `<w:sectPr>${xml}</w:sectPr>`)
    }
    this.#markDirty()
  }

  /**
   * 改写表格中一个单元格的文本。
   *
   * 只改目标 `<w:tc>` 里第一个段落的 run 内容，保留单元格宽度、边框、
   * 底纹与段落属性；表格其余单元格与整篇文档的其它部件一律不动。
   *
   * 行列定位只数**直接子元素**，因此嵌套表格不会被误当成外层表格的行列。
   *
   * @param {object} args - 参数。
   * @param {number} [args.table] - 第几个表格（0 基，默认 0）。
   * @param {number} args.row - 行号（0 基）。
   * @param {number} args.column - 列号（0 基）。
   * @param {string} args.text - 新文本。
   * @returns {object} 变更信息。
   */
  updateTableCell({ table = 0, row, column, text }) {
    const doc = this.#documentDoc()
    const body = find(doc.root, 'body')
    const tables = (body?.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'tbl')
    const tbl = tables[table]
    if (!tbl) {
      throw new OfficeError('FILE_NOT_FOUND', `不存在第 ${table} 个表格（文档共 ${tables.length} 个）。`, { table_count: tables.length })
    }
    const rows = (tbl.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'tr')
    const tr = rows[row]
    if (!tr) {
      throw new OfficeError('FILE_NOT_FOUND', `表格 ${table} 不存在第 ${row} 行（共 ${rows.length} 行）。`, { row_count: rows.length })
    }
    const cells = (tr.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'tc')
    const tc = cells[column]
    if (!tc) {
      throw new OfficeError('FILE_NOT_FOUND', `第 ${row} 行不存在第 ${column} 列（共 ${cells.length} 列）。`, { column_count: cells.length })
    }
    const p = find(tc, 'p')
    if (!p) throw new OfficeError('CORRUPTED_DOCUMENT', '目标单元格里没有段落，无法写入文本。')

    const previous = paragraphText(doc, p)
    const pPr = find(p, 'pPr')
    const from = pPr ? pPr.end : p.startTagEnd
    doc.patch(from, p.contentEnd, runsXml(text ?? ''))
    this.#markDirty()
    return { type: 'update_table_cell', table, row, column, from: preview(previous), to: preview(text ?? '') }
  }

  /**
   * 在指定段落之后插入一张新表格。
   *
   * 表格是 body 级元素（与段落平级），所以插到「第 N 段之后」就是把这个 `<w:tbl>`
   * 节点插到那个 `<w:p>` 后面，段落编号不受影响（`bodyParagraphs()` 只数 `<w:p>`）。
   *
   * 三个细节按 Word 的实际要求处理：
   *   1. `w:tblGrid` 必须给出每列宽度（dxa），否则 Word 会按内容猜列宽；
   *   2. 每个单元格都要有 `w:tcPr/w:tcW`，宽度写死才不会被重排；
   *   3. 文档**以表格结尾**时补一个空段落 —— Word 自己也会这么加，否则尾部表格没有落脚点。
   *
   * @param {object} args - 参数。
   * @param {number} args.after - 插入到第几段之后（0 基）；-1 表示插到正文最前。
   * @param {string[][]} args.rows - 二维文本数组（第一维是行）。
   * @param {boolean} [args.header] - 是否把首行当表头（加粗），默认 true。
   * @param {string} [args.styleId] - 表格样式 ID，默认 TableGrid。
   * @param {number[]} [args.columnWidths] - 每列宽度（dxa，1/20 磅）；省略时按正文宽度均分。
   * @returns {object} 变更信息。
   */
  createTable({ after, rows, header = true, styleId = 'TableGrid', columnWidths = null }) {
    if (!Array.isArray(rows) || rows.length === 0 || !Array.isArray(rows[0])) {
      throw new OfficeError('INVALID_REQUEST', 'rows 必须是非空的二维数组（第一维是行）。')
    }
    const columnCount = rows.reduce((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0)
    if (columnCount === 0) throw new OfficeError('INVALID_REQUEST', 'rows 里没有任何单元格。')

    const doc = this.#documentDoc()
    const body = find(doc.root, 'body')
    if (!body) throw new OfficeError('CORRUPTED_DOCUMENT', 'document.xml 缺少 body 节点。')
    const paragraphs = this.bodyParagraphs()
    const index = after + 1
    if (index < 0 || index > paragraphs.length) {
      throw new OfficeError('INVALID_REQUEST', `插入位置 ${index} 超出范围（文档共 ${paragraphs.length} 段）。`, {
        paragraph_count: paragraphs.length
      })
    }

    // 正文可用宽度：取 sectPr 的页面尺寸与页边距，拿不到就按 A4 + 1 英寸页边距
    const sectPr = find(body, 'sectPr')
    const pageWidth = Number(attr(find(sectPr ?? { children: [] }, 'pgSz') ?? { attrs: new Map() }, 'w') ?? 0) || 11906
    const marginLeft = Number(attr(find(sectPr ?? { children: [] }, 'pgMar') ?? { attrs: new Map() }, 'left') ?? 0) || 1440
    const marginRight = Number(attr(find(sectPr ?? { children: [] }, 'pgMar') ?? { attrs: new Map() }, 'right') ?? 0) || 1440
    const usable = Math.max(1200, pageWidth - marginLeft - marginRight)

    const widths =
      Array.isArray(columnWidths) && columnWidths.length === columnCount
        ? columnWidths.map((value) => Math.max(120, Math.round(Number(value) || 0)))
        : Array.from({ length: columnCount }, () => Math.round(usable / columnCount))

    const grid = widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')
    const rowsXml = rows
      .map((row, rowIndex) => {
        const isHeader = header && rowIndex === 0
        const cells = Array.from({ length: columnCount }, (_, columnIndex) => {
          const text = Array.isArray(row) ? row[columnIndex] : undefined
          const runs = runsXml(text ?? '')
          const body = isHeader && runs !== '' ? runs.replace(/<w:r>/g, '<w:r><w:rPr><w:b/></w:rPr>') : runs
          return (
            `<w:tc><w:tcPr><w:tcW w:w="${widths[columnIndex]}" w:type="dxa"/></w:tcPr>` +
            `<w:p>${body}</w:p></w:tc>`
          )
        }).join('')
        return `<w:tr>${isHeader ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${cells}</w:tr>`
      })
      .join('')
    const tableXml =
      `<w:tbl><w:tblPr><w:tblStyle w:val="${escapeXmlAttr(styleId)}"/>` +
      `<w:tblW w:w="${usable}" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="${header ? 1 : 0}" w:lastRow="0" ` +
      `w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>` +
      `<w:tblGrid>${grid}</w:tblGrid>${rowsXml}</w:tbl>`

    if (index === 0) {
      // 插到正文最前：落在第一个段落之前（body 的第一个元素）
      const firstElement = (body.children ?? []).find((n) => n.type === 'element')
      if (firstElement) doc.insertBefore(firstElement, tableXml)
      else doc.appendChild(body, tableXml)
    } else {
      const anchor = paragraphs[index - 1]
      // 表格成为正文最后一个元素时，**同一次插入**里补一个空段落（Word 也这么做）。
      // 不能在插入后再从 body.children 里找最后一个节点 —— 那是改前的快照。
      const isLast = index === paragraphs.length
      doc.insertAfter(anchor, isLast ? `${tableXml}<w:p/>` : tableXml)
    }
    // 样式定义缺失时补一个，避免 Word 静默退回普通表格
    const styleCreated = styleId === 'TableGrid' ? false : this.#ensureTableStyle(styleId)
    this.#markDirty()
    return {
      type: 'create_table',
      after,
      rows: rows.length,
      columns: columnCount,
      style_id: styleId,
      style_defined_in_document: styleCreated,
      header,
      column_widths: widths
    }
  }

  /**
   * 列出文档正文里的图片（内联与浮动都算）。
   *
   * 索引按它们在 `document.xml` 里出现的先后顺序排列，与 `office_read_docx` 报告的图片数一致。
   *
   * @returns {object[]} 图片清单：`{index, kind, width_px, height_px, name, alt, relationship_id}`。
   */
  images() {
    return this.#imageNodes().map((item, index) => ({
      index,
      kind: item.kind,
      wrap: wrapModeOf(item.inlineNode),
      width_px: emuToPx(Number(attr(item.extent, 'cx') ?? 0)),
      height_px: emuToPx(Number(attr(item.extent, 'cy') ?? 0)),
      name: attr(item.docPr ?? { attrs: new Map() }, 'name') ?? null,
      alt: attr(item.docPr ?? { attrs: new Map() }, 'descr') ?? null,
      relationship_id: item.relationshipId
    }))
  }

  /**
   * 调整图片显示尺寸（像素，按 96 DPI 换算成 EMU）。
   *
   * Word 的布局尺寸取自 `<wp:extent>`，而图形自身还带一份 `<a:xfrm><a:ext>`；
   * 两处都要改 —— 只改一处时，Word 会在下次编辑时把另一处当作真实尺寸，图片「弹回」原大小。
   * 只给一个方向时按原比例缩放，避免把图拉变形。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 第几张图片（0 基）。
   * @param {number} [args.widthPx] - 目标宽度（像素）。
   * @param {number} [args.heightPx] - 目标高度（像素）。
   * @returns {object} 变更信息。
   */
  resizeImage({ index, widthPx = null, heightPx = null }) {
    if (widthPx === null && heightPx === null) {
      throw new OfficeError('INVALID_REQUEST', 'widthPx 与 heightPx 至少要给一个。')
    }
    if ((widthPx !== null && !(widthPx > 0)) || (heightPx !== null && !(heightPx > 0))) {
      throw new OfficeError('INVALID_REQUEST', '尺寸必须是正数。')
    }
    const doc = this.#documentDoc()
    const items = this.#imageNodes()
    const item = items[index]
    if (!item) {
      throw new OfficeError('FILE_NOT_FOUND', `文档里不存在第 ${index} 张图片（共 ${items.length} 张）。`, {
        image_count: items.length
      })
    }
    const fromWidth = Number(attr(item.extent, 'cx') ?? 0)
    const fromHeight = Number(attr(item.extent, 'cy') ?? 0)
    const fromPx = { width: emuToPx(fromWidth), height: emuToPx(fromHeight) }
    let targetWidth = widthPx === null ? Math.round((heightPx * fromWidth) / fromHeight) : widthPx
    let targetHeight = heightPx === null ? Math.round((widthPx * fromHeight) / fromWidth) : heightPx
    targetWidth = Math.max(1, Math.round(targetWidth))
    targetHeight = Math.max(1, Math.round(targetHeight))
    const cx = pxToEmu(targetWidth)
    const cy = pxToEmu(targetHeight)

    doc.setAttr(item.extent, 'cx', String(cx))
    doc.setAttr(item.extent, 'cy', String(cy))
    // 图形内部那份 a:ext 同步改掉
    for (const node of findAll(item.inlineNode, 'ext')) {
      if (attr(node, 'cx') !== undefined && attr(node, 'cy') !== undefined) {
        doc.setAttr(node, 'cx', String(cx))
        doc.setAttr(node, 'cy', String(cy))
      }
    }
    this.#markDirty()
    return {
      type: 'resize_image',
      index,
      kind: item.kind,
      from: fromPx,
      to: { width: targetWidth, height: targetHeight },
      emu: { cx, cy }
    }
  }

  /**
   * 删除一张图片。
   *
   * 图片所在的段落如果只剩这一张图（没有其它文字），整段一起删掉 —— 否则会留下一个空行。
   * 删段落同样走**保护性检查**：段内有书签/批注锚点/修订/域代码时默认拒绝。
   * 媒体部件与关系**保留**（同一张媒体可能被别处引用，这里不做垃圾回收），返回值里明确说明。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 第几张图片（0 基）。
   * @param {boolean} [args.allowMarkupLoss] - 明知会丢失结构化标记仍继续。
   * @returns {object} 变更信息。
   */
  deleteImage({ index, allowMarkupLoss = false }) {
    const doc = this.#documentDoc()
    const items = this.#imageNodes()
    const item = items[index]
    if (!item) {
      throw new OfficeError('FILE_NOT_FOUND', `文档里不存在第 ${index} 张图片（共 ${items.length} 张）。`, {
        image_count: items.length
      })
    }
    const drawing = item.drawing
    const paragraph = item.paragraph
    // 段落里是否只剩这一张图：数一下段内的 drawing 个数与文字长度。
    // 不能用「段落子元素去掉 drawing 后是否为空」——drawing 外面还包着 run，
    // 那个判断永远为假（实测因此留下空段落）。
    const drawingCount = paragraph ? findAll(paragraph, 'drawing').length + findAll(paragraph, 'pict').length : 0
    const paragraphTextLength = paragraph ? findAll(paragraph, 't').map((t) => nodeText(doc, t)).join('').length : 0
    const dropParagraph = Boolean(paragraph) && drawingCount === 1 && paragraphTextLength === 0
    let markup = []
    if (dropParagraph) {
      markup = this.#protectedMarkup(paragraph)
      if (markup.length > 0 && !allowMarkupLoss) {
        throw new OfficeError(
          'INVALID_REQUEST',
          `图片所在段落含结构化标记（${[...new Set(markup)].join('、')}），删除整段会让它们失效；确认请传 allowMarkupLoss=true，或保留段落只删图片。`,
          { markup: [...new Set(markup)], needsConfirmation: true }
        )
      }
    }
    // 删整段时**只删段落**：drawing 的区间在段落内部，两个删除补丁会嵌套重叠，
    // 补丁引擎会（正确地）拒绝这种重叠修改。
    if (dropParagraph) doc.remove(paragraph)
    else doc.remove(drawing)
    this.#markDirty()
    return {
      type: 'delete_image',
      index,
      kind: item.kind,
      removed_paragraph: Boolean(dropParagraph),
      markup_lost: [...new Set(markup)],
      note: '媒体部件与关系保留在包里（同一媒体可能被别处引用），未做垃圾回收。'
    }
  }

  /**
   * 枚举正文里的图片节点（按源码位置排序）。
   * @returns {object[]} `{kind, inlineNode, extent, docPr, drawing, paragraph, relationshipId}`。
   */
  #imageNodes() {
    const doc = this.#documentDoc()
    const containers = [...findAll(doc.root, 'inline'), ...findAll(doc.root, 'anchor')]
      .map((node) => ({ node, kind: localName(node) }))
      .sort((a, b) => a.node.startTagEnd - b.node.startTagEnd)
    return containers.map(({ node, kind }) => {
      const drawing = node.parent
      return {
        kind,
        inlineNode: node,
        extent: find(node, 'extent'),
        docPr: find(node, 'docPr'),
        drawing,
        // 注意：`drawing.parent` 是 **run**（`w:r`），不是段落 —— 要沿 parent 链往上找到 `w:p`。
        // 直接把它当段落会「删掉 run、留下空段落」，段落数不变。
        paragraph: enclosingParagraph(drawing),
        relationshipId: attr(find(node, 'blip') ?? { attrs: new Map() }, 'embed') ?? null
      }
    })
  }

  /**
   * 在指定段落上插入一个书签。
   *
   * 书签是**一对**标记：`<w:bookmarkStart w:id w:name/>` 放在段落属性之后、
   * `<w:bookmarkEnd w:id/>` 放在段落内容末尾。两处的 `w:id` 必须一致且**全文档唯一**——
   * 重号会让 Word 认不出书签之间的范围（目录、交叉引用、书签跳转都会失效）。
   *
   * 空段落上两处插入落在同一偏移，按**登记顺序**生效（先 start 后 end），正好是正确顺序。
   *
   * 名称规则按 Word 的来：以字母（含中文等任意文字字符）或下划线开头，后续只能是
   * 文字字符、数字与下划线，最长 40 字符，且**不区分大小写地**在文档内唯一。
   *
   * @param {object} args - 参数。
   * @param {number} args.paragraph - 段落下标（0 基，只数 body 的直接子段落）。
   * @param {string} args.name - 书签名。
   * @returns {object} 变更信息。
   */
  insertBookmark({ paragraph, name }) {
    const label = String(name ?? '').trim()
    if (!BOOKMARK_NAME.test(label)) {
      throw new OfficeError(
        'INVALID_REQUEST',
        `书签名「${name}」不合法：需以字母（含中文）或下划线开头，后续只能是文字、数字、下划线，最多 40 字符。`
      )
    }
    const paragraphs = this.bodyParagraphs()
    const p = paragraphs[paragraph]
    if (!p) {
      throw new OfficeError('FILE_NOT_FOUND', `不存在第 ${paragraph} 段（共 ${paragraphs.length} 段）。`, {
        paragraph_count: paragraphs.length
      })
    }
    const doc = this.#documentDoc()
    const existing = findAll(doc.root, 'bookmarkStart')
    const lower = label.toLowerCase()
    if (existing.some((node) => (attr(node, 'name') ?? '').toLowerCase() === lower)) {
      throw new OfficeError('INVALID_REQUEST', `书签名「${label}」在文档里已存在（书签名不区分大小写）。`, {
        existing: existing.map((node) => attr(node, 'name'))
      })
    }
    const maxId = existing.reduce((max, node) => Math.max(max, Number(attr(node, 'id') ?? -1)), -1)
    const id = maxId + 1

    const pPr = find(p, 'pPr')
    const from = pPr ? pPr.end : p.startTagEnd
    doc.patch(from, from, `<w:bookmarkStart w:id="${id}" w:name="${escapeXmlAttr(label)}"/>`)
    doc.patch(p.contentEnd, p.contentEnd, `<w:bookmarkEnd w:id="${id}"/>`)
    this.#markDirty()

    return {
      type: 'insert_bookmark',
      part: 'word/document.xml',
      paragraph,
      name: label,
      id,
      text: paragraphText(doc, p),
      bookmark_count: existing.length + 1
    }
  }

  /**
   * 列出正文里的书签（按出现顺序），用于写入后读回核对。
   * @returns {object[]} 书签列表。
   */
  bookmarks() {
    const doc = this.#documentDoc()
    const paragraphs = this.bodyParagraphs()
    const ends = new Set(findAll(doc.root, 'bookmarkEnd').map((node) => attr(node, 'id')))
    return findAll(doc.root, 'bookmarkStart').map((node) => {
      const paragraph = enclosingParagraph(node)
      return {
        name: attr(node, 'name') ?? null,
        id: Number(attr(node, 'id') ?? -1),
        has_end: ends.has(attr(node, 'id')),
        paragraph: paragraph ? paragraphs.indexOf(paragraph) : -1,
        text: paragraph ? paragraphText(doc, paragraph) : ''
      }
    })
  }

  /**
   * 设置图片的环绕方式（行内 ↔ 浮动）。
   *
   * Word 里这是两种**不同的 XML 容器**，不是同一个元素上的一个属性：
   *   - 行内图：`<wp:inline>`，随文字排版，不能自由拖动
   *   - 浮动图：`<wp:anchor>`，带 `simplePos` + `positionH`/`positionV` + 环绕元素，
   *     并且**八个属性全是必填**（`distT/distB/distL/distR/simplePos/relativeHeight/behindDoc/locked/layoutInCell/allowOverlap`）——
   *     少写一个 Word 就报「内容有问题」。
   *
   * 元素顺序同样受 schema 约束：`simplePos → positionH → positionV → extent → effectExtent → 环绕元素 → docPr → ...`，
   * 所以环绕元素要插在 `wp:docPr` **之前**，位置元素要插在 `wp:extent` **之前**。
   *
   * 不支持 `tight` / `through`（紧密、穿越）：它们要求 `wrapPolygon` 包围多边形，
   * 而多边形要按图片实际轮廓算 —— 那属于需要渲染的能力，这里明确拒绝而不是伪造一个矩形。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 第几张图片（0 基）。
   * @param {string} args.wrap - `inline` | `square` | `topAndBottom` | `none` | `behind` | `inFront`。
   * @param {number} [args.offsetXEmu] - 水平偏移（EMU），默认 0。
   * @param {number} [args.offsetYEmu] - 垂直偏移（EMU），默认 0。
   * @param {string} [args.relativeFromH] - 水平参照：`column`（默认）/ `page` / `margin` / `character`。
   * @param {string} [args.relativeFromV] - 垂直参照：`paragraph`（默认）/ `page` / `margin` / `line`。
   * @returns {object} 变更信息。
   */
  setImageWrap({ index, wrap, offsetXEmu = 0, offsetYEmu = 0, relativeFromH = 'column', relativeFromV = 'paragraph' }) {
    // 无环绕在 Word 里没有独立选项：`wrapNone` + `behindDoc="0"` 就是「浮于文字上方」，
    // 因此 `none` 是 `inFront` 的别名，返回与读回统一用 inFront，避免同一份文档两个名字。
    const requested = wrap === 'none' ? 'inFront' : wrap
    const mode = WRAP_MODES[requested]
    if (wrap !== 'inline' && !mode) {
      throw new OfficeError(
        'INVALID_REQUEST',
        `wrap 只支持 inline / ${Object.keys(WRAP_MODES).join(' / ')}；tight / through 需要多边形包围盒，本插件不实现。`
      )
    }
    if (!H_RELATIVE_FROM[relativeFromH]) throw new OfficeError('INVALID_REQUEST', `relativeFromH 不支持「${relativeFromH}」。`)
    if (!V_RELATIVE_FROM[relativeFromV]) throw new OfficeError('INVALID_REQUEST', `relativeFromV 不支持「${relativeFromV}」。`)
    for (const [name, value] of [['offsetXEmu', offsetXEmu], ['offsetYEmu', offsetYEmu]]) {
      if (!Number.isFinite(value)) throw new OfficeError('INVALID_REQUEST', `${name} 必须是数字（EMU）。`)
    }

    const items = this.#imageNodes()
    const item = items[index]
    if (!item) {
      throw new OfficeError('FILE_NOT_FOUND', `不存在第 ${index} 张图片（共 ${items.length} 张）。`, { image_count: items.length })
    }
    const doc = this.#documentDoc()
    const node = item.inlineNode
    const wasInline = item.kind === 'inline'

    if (wrap === 'inline') {
      if (wasInline) {
        return { type: 'set_image_wrap', index, from: 'inline', to: 'inline', changed: false, offset_emu: null }
      }
      // 浮动 → 行内：换标签 + 删掉浮动专有的子元素
      rebuildStartTag(doc, node, 'wp:inline', {}, ANCHOR_ONLY_ATTRS)
      for (const tag of ['simplePos', 'positionH', 'positionV', 'wrapSquare', 'wrapTight', 'wrapThrough', 'wrapTopAndBottom', 'wrapNone']) {
        for (const child of findAll(node, tag)) doc.remove(child)
      }
      this.#markDirty()
      return {
        type: 'set_image_wrap',
        index,
        from: item.kind === 'anchor' ? 'anchor' : item.kind,
        to: 'inline',
        changed: true,
        offset_emu: null
      }
    }

    // 行内 / 浮动 → 浮动
    const anchorAttrs = {
      simplePos: '0',
      relativeHeight: String(mode.behind ? 0 : 251658240),
      behindDoc: mode.behind ? '1' : '0',
      locked: '0',
      layoutInCell: '1',
      allowOverlap: '1'
    }
    if (wasInline) {
      rebuildStartTag(doc, node, 'wp:anchor', anchorAttrs)
      const extent = item.extent
      // 等长零长插入按登记顺序生效：simplePos → positionH → positionV，正好是 schema 要求的顺序
      const posH = `<wp:positionH relativeFrom="${relativeFromH}"><wp:posOffset>${Math.round(offsetXEmu)}</wp:posOffset></wp:positionH>`
      const posV = `<wp:positionV relativeFrom="${relativeFromV}"><wp:posOffset>${Math.round(offsetYEmu)}</wp:posOffset></wp:positionV>`
      if (!extent) throw new OfficeError('CORRUPTED_DOCUMENT', '图片缺少 wp:extent，无法计算位置。')
      doc.insertBefore(extent, '<wp:simplePos x="0" y="0"/>')
      doc.insertBefore(extent, posH)
      doc.insertBefore(extent, posV)
    } else {
      rebuildStartTag(doc, node, 'wp:anchor', anchorAttrs)
      const posH = find(node, 'positionH')
      const posV = find(node, 'positionV')
      if (!posH || !posV) throw new OfficeError('CORRUPTED_DOCUMENT', '浮动图片缺少 positionH / positionV。')
      doc.patch(posH.start, posH.end, `<wp:positionH relativeFrom="${relativeFromH}"><wp:posOffset>${Math.round(offsetXEmu)}</wp:posOffset></wp:positionH>`)
      doc.patch(posV.start, posV.end, `<wp:positionV relativeFrom="${relativeFromV}"><wp:posOffset>${Math.round(offsetYEmu)}</wp:posOffset></wp:positionV>`)
    }

    // 环绕元素要排在 wp:docPr 之前
    const docPr = item.docPr
    if (!docPr) throw new OfficeError('CORRUPTED_DOCUMENT', '图片缺少 wp:docPr。')
    const existingWrap = ['wrapSquare', 'wrapTopAndBottom', 'wrapNone'].map((tag) => find(node, tag)).find(Boolean)
    if (existingWrap) doc.patch(existingWrap.start, existingWrap.end, mode.xml)
    else doc.insertBefore(docPr, mode.xml)

    this.#markDirty()
    return {
      type: 'set_image_wrap',
      index,
      from: wasInline ? 'inline' : 'anchor',
      to: requested,
      changed: true,
      offset_emu: { x: Math.round(offsetXEmu), y: Math.round(offsetYEmu) },
      relative_from: { horizontal: relativeFromH, vertical: relativeFromV },
      behind_text: mode.behind
    }
  }

  /**
   * 给表格套用样式（并设置 Word 的「表格样式选项」）。
   *
   * 两件事要一起做，只做一件在 Word 里都看不出效果：
   *   1. `w:tblPr/w:tblStyle w:val="…"` 指向 styles.xml 里的表格样式；
   *   2. `w:tblPr/w:tblLook` 声明首行/首列/镶边行等条件格式开关 —— 样式本身只是定义，
   *      哪些条件生效由 tblLook 决定（`w:noHBand="0"` 才表示镶边行开启）。
   *
   * 指定的样式在 styles.xml 里不存在时，会**补一个最小定义**（名称 + 单线边框），
   * 而不是留一个指向空气的引用 —— 那种情况下 Word 会静默退回「普通表格」。
   *
   * @param {object} args - 参数。
   * @param {number} [args.table] - 第几个表格（0 基）。
   * @param {string} args.styleId - 表格样式 ID，如 TableGrid、LightShading-Accent1。
   * @param {boolean} [args.firstRow] - 强调首行，默认 true。
   * @param {boolean} [args.lastRow] - 强调末行，默认 false。
   * @param {boolean} [args.firstColumn] - 强调首列，默认 false。
   * @param {boolean} [args.lastColumn] - 强调末列，默认 false。
   * @param {boolean} [args.bandedRows] - 镶边行，默认 true。
   * @returns {object} 变更信息。
   */
  setTableStyle({
    table = 0,
    styleId,
    firstRow = true,
    lastRow = false,
    firstColumn = false,
    lastColumn = false,
    bandedRows = true
  }) {
    if (typeof styleId !== 'string' || !/^[A-Za-z][A-Za-z0-9_-]*$/.test(styleId)) {
      throw new OfficeError('INVALID_REQUEST', `样式 ID 不合法：${styleId}（应形如 TableGrid、LightShading-Accent1）。`)
    }
    const { doc, tbl } = this.#tableContext(table)
    const styleCreated = this.#ensureTableStyle(styleId)

    // 1) w:tblPr/w:tblStyle —— 按 CT_TblPr 顺序，tblStyle 必须排在最前
    let tblPr = find(tbl, 'tblPr')
    if (!tblPr) {
      const firstChild = (tbl.children ?? []).find((n) => n.type === 'element')
      const xml = '<w:tblPr><w:tblStyle w:val="' + escapeXmlAttr(styleId) + '"/></w:tblPr>'
      if (firstChild) doc.insertBefore(firstChild, xml)
      else doc.appendChild(tbl, xml)
      tblPr = find(tbl, 'tblPr')
    } else {
      // 注意：删除节点后**不能**再从 children 里找插入锚点 —— 那是改前的快照，
      // 拿到的可能正是刚被删掉的节点，结果新旧节点同时留在文档里（实测写出两个 tblStyle）。
      // 锚点必须提前取。
      const children = (tblPr.children ?? []).filter((n) => n.type === 'element')
      const anchor = children.find((n) => localName(n) !== 'tblStyle')
      const xml = `<w:tblStyle w:val="${escapeXmlAttr(styleId)}"/>`
      for (const node of children.filter((n) => localName(n) === 'tblStyle')) doc.remove(node)
      if (anchor) doc.insertBefore(anchor, xml)
      else doc.appendChild(tblPr, xml)
    }

    // 2) w:tblLook —— 条件格式开关。Word 的写法：不用某条件时对应位为 1
    const lookAttrs = [
      `w:val="${firstRow ? '04A0' : '0000'}"`,
      `w:firstRow="${firstRow ? 1 : 0}"`,
      `w:lastRow="${lastRow ? 1 : 0}"`,
      `w:firstColumn="${firstColumn ? 1 : 0}"`,
      `w:lastColumn="${lastColumn ? 1 : 0}"`,
      `w:noHBand="${bandedRows ? 0 : 1}"`,
      'w:noVBand="1"'
    ].join(' ')
    const lookChildren = (tblPr.children ?? []).filter((n) => n.type === 'element')
    // tblLook 之后还有 tblCaption / tblDescription / tblPrChange，锚点同样提前取
    const lookAnchor = lookChildren.find((n) => ['tblCaption', 'tblDescription', 'tblPrChange'].includes(localName(n)))
    for (const node of lookChildren.filter((n) => localName(n) === 'tblLook')) doc.remove(node)
    const lookXml = `<w:tblLook ${lookAttrs}/>`
    if (lookAnchor) doc.insertBefore(lookAnchor, lookXml)
    else doc.appendChild(tblPr, lookXml)

    this.#markDirty()
    return {
      type: 'set_table_style',
      table,
      style_id: styleId,
      style_defined_in_document: styleCreated,
      look: { first_row: firstRow, last_row: lastRow, first_column: firstColumn, last_column: lastColumn, banded_rows: bandedRows }
    }
  }

  /**
   * 确保 styles.xml 里有指定 styleId 的表格样式；没有就补一个最小定义。
   * @param {string} styleId - 样式 ID。
   * @returns {boolean} 是否新建了样式定义。
   */
  #ensureTableStyle(styleId) {
    return this.#ensureStyle('table', styleId)
  }

  /**
   * 确保 styles.xml 里有指定 styleId 的某种样式；没有就补一个最小定义。
   *
   * 为什么必须补：`w:pStyle` / `w:rStyle` / `w:tblStyle` 指向一个不存在的 styleId 时，
   * Word 不报错，而是**静默按默认样式显示** —— 用户会以为「设置了没生效」。
   *
   * @param {string} kind - `table` | `paragraph` | `character`。
   * @param {string} styleId - 样式 ID。
   * @returns {boolean} 是否新建了样式定义。
   */
  #ensureStyle(kind, styleId) {
    if (!['table', 'paragraph', 'character'].includes(kind)) {
      throw new OfficeError('INVALID_REQUEST', `不支持的样式类型：${kind}`)
    }
    if (!this.#pkg.has('word/styles.xml')) {
      throw new OfficeError('CORRUPTED_DOCUMENT', '文档缺少 word/styles.xml，无法套用样式。')
    }
    const doc = XmlDoc.parse(this.#pkg.readText('word/styles.xml'))
    const exists = findAll(doc.root, 'style').some(
      (node) => attr(node, 'styleId') === styleId && attr(node, 'type') === kind
    )
    if (exists) return false
    const root = doc.root.children.find((c) => c.type === 'element')
    // 尽量少写内容：只给名称（表格样式再给单线边框）。
    // 不写 basedOn，避免引用到文档里并不存在的样式定义。
    const body =
      kind === 'table'
        ? '<w:tblPr><w:tblBorders>' +
          ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
            .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
            .join('') +
          '</w:tblBorders></w:tblPr>'
        : ''
    const xml =
      `<w:style w:type="${kind}" w:styleId="${escapeXmlAttr(styleId)}">` +
      `<w:name w:val="${escapeXmlAttr(styleId)}"/><w:uiPriority w:val="99"/>${body}</w:style>`
    doc.appendChild(root, xml)
    // styles.xml 不参与增量脏标记，改完直接写回包
    this.#pkg.write('word/styles.xml', doc.toString())
    return true
  }

  /**
   * 给某个段落套用段落样式（`w:pPr/w:pStyle`）。
   *
   * 只动段落属性里的样式引用，不改 run、不动编号与对齐 ——
   * 与原 `pPr` 里的其它设置（缩进、间距、大纲级别）共存。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 段落下标（0 基，只计正文直接段落）。
   * @param {string} args.styleId - 段落样式 ID，如 Heading2、Quote。
   * @returns {object} 变更信息。
   */
  setParagraphStyle({ index, styleId }) {
    if (typeof styleId !== 'string' || !/^[A-Za-z][A-Za-z0-9_-]*$/.test(styleId)) {
      throw new OfficeError('INVALID_REQUEST', `样式 ID 不合法：${styleId}`)
    }
    const doc = this.#documentDoc()
    const paragraph = this.bodyParagraphs()[index]
    if (!paragraph) {
      const count = this.bodyParagraphs().length
      throw new OfficeError('FILE_NOT_FOUND', `不存在第 ${index} 段（文档共 ${count} 段）。`, { paragraph_count: count })
    }
    const previous = attr(find(paragraph, 'pPr') ? find(find(paragraph, 'pPr'), 'pStyle') ?? { attrs: new Map() } : { attrs: new Map() }, 'val')
    this.#applyStyleRef(doc, paragraph, 'pPr', 'pStyle', styleId)
    const created = this.#ensureStyle('paragraph', styleId)
    this.#markDirty()
    return { type: 'set_paragraph_style', index, style_id: styleId, from: previous ?? null, style_defined_in_document: created }
  }

  /**
   * 给某个段落里的所有 run 套用字符样式（`w:rPr/w:rStyle`）。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 段落下标（0 基）。
   * @param {string} args.styleId - 字符样式 ID，如 Strong、Emphasis。
   * @returns {object} 变更信息。
   */
  setCharacterStyle({ index, styleId }) {
    if (typeof styleId !== 'string' || !/^[A-Za-z][A-Za-z0-9_-]*$/.test(styleId)) {
      throw new OfficeError('INVALID_REQUEST', `样式 ID 不合法：${styleId}`)
    }
    const doc = this.#documentDoc()
    const paragraph = this.bodyParagraphs()[index]
    if (!paragraph) {
      const count = this.bodyParagraphs().length
      throw new OfficeError('FILE_NOT_FOUND', `不存在第 ${index} 段（文档共 ${count} 段）。`, { paragraph_count: count })
    }
    const runs = findAll(paragraph, 'r')
    for (const run of runs) this.#applyStyleRef(doc, run, 'rPr', 'rStyle', styleId)
    const created = this.#ensureStyle('character', styleId)
    this.#markDirty()
    return { type: 'set_character_style', index, style_id: styleId, runs: runs.length, style_defined_in_document: created }
  }

  /**
   * 在段落属性或 run 属性里写入样式引用（`w:pStyle` / `w:rStyle`）。
   *
   * 属性容器不存在时先建一个，并且必须放在**第一个子元素**位置（`w:pPr`/`w:rPr` 是段落/run 的首个子元素）。
   *
   * @param {object} doc - 文档。
   * @param {object} node - 段落或 run 节点。
   * @param {string} containerName - `pPr` 或 `rPr`。
   * @param {string} refName - `pStyle` 或 `rStyle`。
   * @param {string} styleId - 样式 ID。
   * @returns {void}
   */
  #applyStyleRef(doc, node, containerName, refName, styleId) {
    let container = find(node, containerName)
    if (!container) {
      const firstChild = (node.children ?? []).find((n) => n.type === 'element')
      const wrapper = `<w:${containerName}><w:${refName} w:val="${escapeXmlAttr(styleId)}"/></w:${containerName}>`
      if (firstChild) doc.insertBefore(firstChild, wrapper)
      else doc.appendChild(node, wrapper)
      return
    }
    // 先取锚点再删除（children 是改前快照）
    const children = (container.children ?? []).filter((n) => n.type === 'element')
    const anchor = children.find((n) => localName(n) !== refName)
    for (const existing of children.filter((n) => localName(n) === refName)) doc.remove(existing)
    const xml = `<w:${refName} w:val="${escapeXmlAttr(styleId)}"/>`
    if (anchor) doc.insertBefore(anchor, xml)
    else doc.appendChild(container, xml)
  }

  /**
   * 在指定段落之后插入一个分页符（新起一个只含分页符的段落）。
   *
   * Word 的分页符就是一个 run 里的 `<w:br w:type="page"/>`；放在独立段落里最稳，
   * 不会把后面段落的文字挤到同一行。
   *
   * @param {object} args - 参数。
   * @param {number} args.after - 在第几段之后插入（0 基）；-1 表示插到正文最前。
   * @returns {object} 变更信息。
   */
  insertPageBreak({ after }) {
    const doc = this.#documentDoc()
    const paragraphs = this.bodyParagraphs()
    const index = after + 1
    if (index < 0 || index > paragraphs.length) {
      throw new OfficeError('INVALID_REQUEST', `插入位置 ${index} 超出范围（文档共 ${paragraphs.length} 段）。`, {
        paragraph_count: paragraphs.length
      })
    }
    const xml = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    if (index === 0) {
      const body = find(doc.root, 'body')
      const firstChild = (body?.children ?? []).find((n) => n.type === 'element')
      if (firstChild) doc.insertBefore(firstChild, xml)
      else doc.appendChild(body, xml)
    } else {
      doc.insertAfter(paragraphs[index - 1], xml)
    }
    this.#markDirty()
    return { type: 'insert_page_break', after, index }
  }

  /**
   * 在表格里插入一行。
   *
   * 新行的格式从**参照行**复制（`w:trPr` 行属性 + 每个单元格的 `w:tcPr`），
   * 这样列宽、边框、底纹、行高都跟着走，不需要自己去拼表格属性 ——
   * 自己新造一个 `w:tr` 而不带 `w:tcPr`，Word 会按默认宽度重排整张表。
   *
   * @param {object} args - 参数。
   * @param {number} [args.table] - 第几个表格（0 基）。
   * @param {number} args.after - 插到第几行之后（0 基）；-1 表示插到表格最前。
   * @param {string[]} [args.cells] - 各单元格文本；缺省则插入空单元格。
   * @returns {object} 变更信息。
   */
  insertTableRow({ table = 0, after, cells = [] }) {
    const { doc, tbl, rows } = this.#tableContext(table)
    const referenceIndex = after < 0 ? 0 : after
    const reference = rows[referenceIndex]
    if (!reference) {
      throw new OfficeError('FILE_NOT_FOUND', `表格 ${table} 不存在第 ${after} 行（共 ${rows.length} 行）。`, {
        row_count: rows.length
      })
    }
    const referenceCells = (reference.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'tc')
    const trPr = find(reference, 'trPr')
    const rowXml =
      `<w:tr>${trPr ? doc.raw(trPr) : ''}` +
      referenceCells
        .map((tc, index) => {
          const tcPr = find(tc, 'tcPr')
          return `<w:tc>${tcPr ? doc.raw(tcPr) : ''}<w:p>${runsXml(cells[index] ?? '')}</w:p></w:tc>`
        })
        .join('') +
      `</w:tr>`

    if (after < 0) doc.insertBefore(rows[0], rowXml)
    else doc.insertAfter(reference, rowXml)
    this.#markDirty()
    return {
      type: 'insert_table_row',
      table,
      row: referenceIndex + (after < 0 ? 0 : 1),
      columns: referenceCells.length,
      row_count: rows.length + 1,
      cells: referenceCells.map((_, index) => cells[index] ?? '')
    }
  }

  /**
   * 删除表格里的一行。
   *
   * 与删除段落一样做**保护性检查**：行内含书签、批注锚点、修订标记或域代码时默认拒绝，
   * 因为它们会随行一起消失（Word 会静默丢弃对应功能），需要显式 `allowMarkupLoss`。
   *
   * @param {object} args - 参数。
   * @param {number} [args.table] - 第几个表格（0 基）。
   * @param {number} args.row - 要删除的行下标（0 基）。
   * @param {boolean} [args.allowMarkupLoss] - 明知会丢失结构化标记仍继续。
   * @returns {object} 变更信息。
   */
  deleteTableRow({ table = 0, row, allowMarkupLoss = false }) {
    const { doc, rows } = this.#tableContext(table)
    const tr = rows[row]
    if (!tr) {
      throw new OfficeError('FILE_NOT_FOUND', `表格 ${table} 不存在第 ${row} 行（共 ${rows.length} 行）。`, {
        row_count: rows.length
      })
    }
    if (rows.length <= 1) {
      throw new OfficeError('INVALID_REQUEST', '表格只剩一行，删除后表格不再有效（Word 要求表格至少一行）。')
    }
    const markup = []
    for (const p of findAll(tr, 'p')) markup.push(...this.#protectedMarkup(p))
    if (markup.length > 0 && !allowMarkupLoss) {
      throw new OfficeError(
        'INVALID_REQUEST',
        `该行含结构化标记（${[...new Set(markup)].join('、')}），删除会让它们失效；确认要删除请传 allowMarkupLoss=true。`,
        { markup: [...new Set(markup)], needsConfirmation: true }
      )
    }
    const text = findAll(tr, 't').map((t) => nodeText(doc, t)).join('')
    doc.remove(tr)
    this.#markDirty()
    return { type: 'delete_table_row', table, row, row_count: rows.length - 1, removed_text: preview(text), markup_lost: [...new Set(markup)] }
  }

  /**
   * 合并一个矩形单元格区域（Word 的「合并单元格」）。
   *
   * OOXML 的合并分两步，缺一步 Word 就显示成没合并：
   *   1. **横向**：区域首格写 `<w:gridSpan w:val="N"/>`，并把该行被合并掉的 N-1 个 `<w:tc>` 删掉；
   *   2. **纵向**：首行写 `<w:vMerge w:val="restart"/>`，其余行写 `<w:vMerge/>`（继续合并）。
   * 被合并掉的单元格里的文字按 Word 的行为**并入左上角单元格**（用换行分隔），不丢内容。
   *
   * @param {object} args - 参数。
   * @param {number} [args.table] - 第几个表格（0 基）。
   * @param {number} args.row - 区域左上角行（0 基）。
   * @param {number} args.column - 区域左上角列（0 基）。
   * @param {number} [args.rowSpan] - 纵向跨几行，默认 1。
   * @param {number} [args.colSpan] - 横向跨几列，默认 1。
   * @returns {object} 变更信息。
   */
  mergeTableCells({ table = 0, row, column, rowSpan = 1, colSpan = 1 }) {
    if (!Number.isInteger(rowSpan) || rowSpan < 1 || !Number.isInteger(colSpan) || colSpan < 1) {
      throw new OfficeError('INVALID_REQUEST', 'rowSpan / colSpan 必须是正整数。')
    }
    if (rowSpan === 1 && colSpan === 1) {
      throw new OfficeError('INVALID_REQUEST', 'rowSpan 与 colSpan 不能同时为 1（没有可合并的单元格）。')
    }
    const { doc, rows } = this.#tableContext(table)
    const region = []
    for (let r = row; r < row + rowSpan; r += 1) {
      const tr = rows[r]
      if (!tr) throw new OfficeError('INVALID_REQUEST', `第 ${r} 行不存在（表格共 ${rows.length} 行）。`, { row_count: rows.length })
      const cells = (tr.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'tc')
      if (column + colSpan > cells.length) {
        throw new OfficeError('INVALID_REQUEST', `第 ${r} 行只有 ${cells.length} 列，无法合并到第 ${column + colSpan} 列。`, {
          column_count: cells.length
        })
      }
      region.push({ tr, cells })
    }

    // 收集被并入的文字（Word 的行为：合并后内容保留在左上角）
    const others = []
    const topLeft = region[0].cells[column]
    for (const { cells } of region) {
      for (let c = column; c < column + colSpan; c += 1) {
        if (cells[c] === topLeft) continue
        const text = findAll(cells[c], 't').map((t) => nodeText(doc, t)).join('').trim()
        if (text !== '') others.push(text)
      }
    }

    // 1) 横向合并：首格加 <w:gridSpan>，其余格删掉
    for (const { cells } of region) {
      const first = cells[column]
      ensureTcPrChild(doc, first, 'gridSpan', `<w:gridSpan w:val="${colSpan}"/>`)
      for (let c = column + 1; c < column + colSpan; c += 1) doc.remove(cells[c])
    }
    // 2) 纵向合并：首行 restart，其余行继续
    if (rowSpan > 1) {
      for (let r = 0; r < rowSpan; r += 1) {
        const tr = rows[row + r]
        const cells = (tr.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'tc')
        const tc = cells[column]
        if (!tc) throw new OfficeError('INVALID_REQUEST', `第 ${row + r} 行缺少第 ${column} 列，无法纵向合并。`)
        ensureTcPrChild(doc, tc, 'vMerge', r === 0 ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>')
      }
    }

    // 3) 把并入的文字写进左上角单元格：原本空白就写进第一个段落，否则另起段落（不丢内容）
    if (others.length > 0) {
      const firstP = find(topLeft, 'p')
      const topText = firstP ? findAll(firstP, 't').map((t) => nodeText(doc, t)).join('').trim() : ''
      if (topText === '' && firstP) {
        const pPr = find(firstP, 'pPr')
        doc.patch(pPr ? pPr.end : firstP.startTagEnd, firstP.contentEnd, runsXml(others[0]))
        for (const text of others.slice(1)) doc.appendChild(topLeft, `<w:p>${runsXml(text)}</w:p>`)
      } else {
        for (const text of others) doc.appendChild(topLeft, `<w:p>${runsXml(text)}</w:p>`)
      }
    }

    // 4) 纵向合并时，被并入的那些单元格要**清空内容**：Word 合并后这些内容已经搬到首格，
    //    留着会在「取消合并」时冒出重复文字（Word 自己合并时也是这么处理的）。
    if (rowSpan > 1) {
      for (let r = 1; r < rowSpan; r += 1) {
        const tr = rows[row + r]
        const cells = (tr.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'tc')
        const tc = cells[column]
        if (!tc) continue
        for (const p of (tc.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'p')) {
          const pPr = find(p, 'pPr')
          doc.patch(pPr ? pPr.end : p.startTagEnd, p.contentEnd, '')
        }
      }
    }

    this.#markDirty()
    return {
      type: 'merge_table_cells',
      table,
      row,
      column,
      row_span: rowSpan,
      col_span: colSpan,
      merged_text: others
    }
  }

  /**
   * 取表格上下文（文档、表格节点、行节点）。
   * @param {number} table - 第几个表格（0 基）。
   * @returns {{doc: object, tbl: object, rows: object[]}} 上下文。
   */
  #tableContext(table) {
    const doc = this.#documentDoc()
    const body = find(doc.root, 'body')
    const tables = (body?.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'tbl')
    const tbl = tables[table]
    if (!tbl) {
      throw new OfficeError('FILE_NOT_FOUND', `不存在第 ${table} 个表格（文档共 ${tables.length} 个）。`, { table_count: tables.length })
    }
    const rows = (tbl.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'tr')
    return { doc, tbl, rows }
  }

  /**
   * 在指定位置之后插入一张图片（新起一个只含图片的段落）。
   *
   * 需要新增四个东西，缺一个 Word 都会报文档损坏：
   *   1. `word/media/` 下的媒体部件
   *   2. `[Content_Types].xml` 里该扩展名的 Default 声明
   *   3. `word/_rels/document.xml.rels` 里的 image 关系
   *   4. 正文里引用该关系的 `<w:drawing>`
   *
   * 命名空间（wp/a/pic）直接声明在插入的子树根上，而不是去改 `<w:document>`
   * 的根元素 —— 这样对文档其它部分的字节影响为零。
   *
   * @param {object} args - 参数。
   * @param {number} args.after - 插入到第几段之后（0 基）；-1 表示插到最前。
   * @param {Buffer} args.data - 图片字节。
   * @param {string} args.extension - 图片扩展名（已按真实字节判定）。
   * @param {number} [args.widthPx] - 显示宽度（像素，按 96 DPI 换算）；省略则用图片原始尺寸。
   * @param {number} [args.heightPx] - 显示高度（像素）；只给宽度时按比例缩放。
   * @param {string} [args.altText] - 替代文本。
   * @returns {object} 变更信息。
   */
  insertImage({ after, data, extension, widthPx, heightPx, altText }) {
    const doc = this.#documentDoc()
    const paragraphs = this.bodyParagraphs()
    const index = after + 1
    if (index < 0 || index > paragraphs.length) {
      throw new OfficeError('INVALID_REQUEST', `插入位置 ${index} 超出范围（文档共 ${paragraphs.length} 段）。`, {
        paragraph_count: paragraphs.length
      })
    }

    const mediaName = this.#nextMediaName(extension)
    this.#pkg.write(mediaName, data)
    ensureContentTypeDefault(this.#pkg, extension, imageContentType(extension))
    const rid = addRelationshipTo(this.#pkg, 'word/_rels/document.xml.rels', REL_IMAGE, mediaName.replace(/^word\//, ''))

    const natural = readImageSize(data, extension)
    const emu = resolveImageExtent({ natural, widthPx, heightPx })
    const drawingId = this.#nextDrawingId()
    const xml = buildImageParagraphXml({ rid, cx: emu.cx, cy: emu.cy, id: drawingId, altText })

    const anchor = paragraphs[index]
    if (anchor) doc.insertBefore(anchor, xml)
    else doc.appendChild(find(doc.root, 'body'), xml)
    this.#markDirty()

    return {
      type: 'insert_image',
      index,
      media_part: mediaName,
      relationship_id: rid,
      natural_size: natural,
      display_size_emu: { cx: emu.cx, cy: emu.cy },
      display_size_px: { width: Math.round(emu.cx / EMU_PER_PX), height: Math.round(emu.cy / EMU_PER_PX) },
      bytes: data.length
    }
  }

  /**
   * 取下一个可用的媒体部件名。
   * @param {string} extension - 扩展名。
   * @returns {string} 形如 `word/media/image1.png`。
   */
  #nextMediaName(extension) {
    let n = 1
    while (this.#pkg.has(`word/media/image${n}.${extension}`)) n += 1
    return `word/media/image${n}.${extension}`
  }

  /**
   * 取下一个可用的 drawing 对象 id。
   *
   * `wp:docPr/@id` 在同一文档内必须唯一，Word 对重复 id 会报错，
   * 因此扫描一遍既有 drawing 取最大值 +1，而不是从 1 开始猜。
   *
   * @returns {number} 可用的 id。
   */
  #nextDrawingId() {
    const doc = this.#documentDoc()
    let max = 0
    for (const node of findAll(doc.root, 'docPr')) {
      const id = Number(attr(node, 'id') ?? 0)
      if (Number.isFinite(id) && id > max) max = id
    }
    return max + 1
  }

  /** @returns {ZipPackage} 底层 OPC 包（供校验工具复用）。 */
  get pkg() {
    return this.#pkg
  }
}

/**
 * 复杂域形式的 PAGE 页码域。
 *
 * 用 fldChar 三段式而不是 fldSimple：fldSimple 在某些读取器里不刷新，
 * 三段式是 Word 自己写出来的形态，兼容性最好。
 */
const PAGE_FIELD_XML =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>1</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

/* EMU_PER_PX / IMAGE_TYPES / imageContentType / detectImageType / readImageSize
 * 已下沉到 ooxml.js，与 PPTX 适配器共用同一份图片工具。 */

/** 关系类型：图片。 */
const REL_IMAGE = `${NS_R}/image`

/** 修订标记的中文名。 */
const CHANGE_TYPES = Object.freeze({
  ins: '插入',
  del: '删除',
  moveFrom: '移动源',
  moveTo: '移动目标',
  rPrChange: '字符格式变更',
  pPrChange: '段落格式变更',
  tblPrChange: '表格属性变更',
  trPrChange: '行属性变更',
  tcPrChange: '单元格属性变更',
  sectPrChange: '节属性变更'
})

/** 合法书签名：字母（含中文等文字字符）或下划线开头，后续文字/数字/下划线，最长 40 字符。 */
const BOOKMARK_NAME = /^[\p{L}_][\p{L}\p{N}_]{0,39}$/u

/** 结构化标记的中文名，用于生成人能看懂的拒绝原因。 */
const MARKUP_LABELS = Object.freeze({
  bookmarkStart: '书签起始标记',
  bookmarkEnd: '书签结束标记',
  commentRangeStart: '批注范围起始',
  commentRangeEnd: '批注范围结束',
  commentReference: '批注引用',
  ins: '修订插入标记',
  del: '修订删除标记',
  fldChar: '域字符',
  instrText: '域代码',
  fldSimple: '简单域',
  hyperlink: '超链接',
  sdt: '内容控件',
  footnoteReference: '脚注引用',
  endnoteReference: '尾注引用'
})

/* detectImageType / imageContentType / readImageSize / EMU_PER_PX 已下沉到
 * ooxml.js，与 PPTX 适配器共用同一份图片工具（见文件顶部 import）。 */

/* resolveImageExtent 已下沉到 ooxml.js，与 PPTX 适配器共用。 */

/**
 * 构造「只含一张图片的段落」的 XML。
 *
 * 命名空间声明在子树根上，避免改动 `<w:document>` 根元素。
 *
 * @param {object} args - `{rid, cx, cy, id, altText}`。
 * @returns {string} 段落 XML。
 */
function buildImageParagraphXml({ rid, cx, cy, id, altText }) {
  const descr = altText ? ` descr="${escapeXmlAttr(altText)}"` : ''
  const name = `Picture ${id}`
  return (
    '<w:p><w:r><w:drawing>' +
    `<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    `<wp:docPr id="${id}" name="${name}"${descr}/>` +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
  )
}

/* 内容类型与关系操作已下沉到 ooxml.js（ensureContentTypeDefault /
 * ensureContentTypeOverride / addRelationshipTo），三个适配器共用一份实现。 */

/**
 * CT_SectPr 的子元素顺序（ECMA-376 §17.6.17）。
 *
 * Word 对节属性的子元素顺序敏感：`pgMar` 必须紧跟 `pgSz`，`cols` 在其后，
 * 顺序错了 Word 会报「文档内容有问题」。所以插入子元素要按这张表找位置。
 */
const SECTPR_ORDER = Object.freeze([
  'headerReference',
  'footerReference',
  'footnotePr',
  'endnotePr',
  'type',
  'pgSz',
  'pgMar',
  'paperSrc',
  'pgBorders',
  'lnNumType',
  'pgNumType',
  'cols',
  'formProt',
  'vAlign',
  'noEndnote',
  'titlePg',
  'textDirection',
  'bidi',
  'rtlGutter',
  'docGrid',
  'printerSettings'
])

/** CT_PPr 子元素顺序（只列到本插件会用到的部分）：`jc` 在 `ind` 之后、`rPr`/`sectPr` 之前。 */
const PPR_ORDER = Object.freeze([
  'pStyle',
  'keepNext',
  'keepLines',
  'pageBreakBefore',
  'framePr',
  'widowControl',
  'numPr',
  'suppressLineNumbers',
  'pBdr',
  'shd',
  'tabs',
  'suppressAutoHyphens',
  'kinsoku',
  'wordWrap',
  'overflowPunct',
  'topLinePunct',
  'autoSpaceDE',
  'autoSpaceDN',
  'bidi',
  'adjustRightInd',
  'snapToGrid',
  'spacing',
  'ind',
  'contextualSpacing',
  'mirrorIndents',
  'suppressOverlap',
  'jc',
  'textDirection',
  'textAlignment',
  'textboxTightWrap',
  'outlineLvl',
  'divId',
  'cnfStyle',
  'rPr',
  'sectPr',
  'pPrChange'
])

/** 对齐值白名单（`w:jc` 的取值）。 */
const ALIGN_VALUES = Object.freeze({ left: 'left', center: 'center', right: 'right' })

/**
 * 纸张预设（twips，1/20 磅；1 英寸 = 1440 twips）。
 *
 * A4 210×297mm、A3 297×420mm、A5 148×210mm、Letter 8.5×11in、Legal 8.5×14in，
 * 换算按 1 英寸 = 25.4mm 精确折算后四舍五入——与 Word 写入的值一致。
 */
const PAPER_SIZES = Object.freeze({
  A4: { width: 11906, height: 16838 },
  A3: { width: 16838, height: 23811 },
  A5: { width: 8391, height: 11906 },
  Letter: { width: 12240, height: 15840 },
  Legal: { width: 12240, height: 20160 }
})

/**
 * 新建 `pgMar` 时的默认页边距（twips）。
 *
 * 取 Word 简体中文默认版式：上下 2.54cm、左右 3.17cm、页眉页脚 1.5cm、装订线 0。
 * 缺属性的 `pgMar` 会让不同阅读器各取各的默认值，所以新建时全部写齐。
 */
const MARGIN_DEFAULTS = Object.freeze({ top: 1440, right: 1797, bottom: 1440, left: 1797, header: 851, footer: 851, gutter: 0 })

const TWIPS_PER_CM = 1440 / 2.54

/**
 * 厘米 → twips。
 * @param {number} cm - 厘米。
 * @returns {number} twips（整数）。
 */
function cmToTwips(cm) {
  return Math.round(cm * TWIPS_PER_CM)
}

/**
 * twips → 厘米（保留两位小数）。
 * @param {number} twips - twips。
 * @returns {number} 厘米。
 */
function twipsToCm(twips) {
  return Math.round((twips / TWIPS_PER_CM) * 100) / 100
}

/**
 * 取末节的 `w:sectPr`。
 *
 * 只看 `w:body` 的**直接子** sectPr：段落内部的 sectPr 是分节符，
 * 改动它只影响前面那一节，不是「整篇文档的页面设置」。
 *
 * @param {object} doc - 文档。
 * @returns {object|null} sectPr 节点或 null。
 */
function lastBodySectPr(doc) {
  const body = find(doc.root, 'body')
  if (!body) return null
  const sections = (body.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'sectPr')
  return sections.at(-1) ?? null
}

/**
 * 确保父元素下有某个子元素（按 schema 顺序插入；已有同名元素先删再插）。
 *
 * 锚点必须在删除之前取：`children` 是改前快照，删完再找可能拿到刚被删掉的节点。
 *
 * @param {object} doc - 文档。
 * @param {object} parent - 父元素。
 * @param {string} name - 子元素本地名。
 * @param {string} xml - 子元素 XML。
 * @param {readonly string[]} order - 该父元素的子元素顺序表。
 * @returns {void}
 */
function ensureOrderedChild(doc, parent, name, xml, order) {
  const children = (parent.children ?? []).filter((n) => n.type === 'element')
  const target = order.indexOf(name)
  const anchor = children.find((n) => localName(n) !== name && order.indexOf(localName(n)) > target)
  for (const node of children.filter((n) => localName(n) === name)) doc.remove(node)
  if (anchor) doc.insertBefore(anchor, xml)
  else doc.appendChild(parent, xml)
}

/**
 * `w:jc` 元素 XML。
 * @param {'left'|'center'|'right'} align - 对齐方式。
 * @returns {string} XML。
 */
function paragraphAlignXml(align) {
  return `<w:jc w:val="${ALIGN_VALUES[align]}"/>`
}

/**
 * 在已有的 `w:pPr` 内部设置水平对齐（`w:jc`）。
 *
 * 只改 `w:pPr` 自己的区间，**不碰段落正文**：正文的替换补丁从 `pPr.end` 开始，
 * 两者区间不重叠。若段落没有 `pPr`，调用方应把 `<w:pPr>` 与正文写在同一个补丁里。
 *
 * @param {object} doc - 文档。
 * @param {object} pPr - `w:pPr` 节点。
 * @param {'left'|'center'|'right'} align - 对齐方式。
 * @returns {void}
 */
function setAlignmentInPPr(doc, pPr, align) {
  const jcXml = paragraphAlignXml(align)
  const existing = find(pPr, 'jc')
  if (existing) {
    doc.patch(existing.start, existing.end, jcXml)
    return
  }
  ensureOrderedChild(doc, pPr, 'jc', jcXml, PPR_ORDER)
}

/**
 * 图片环绕方式（`office_set_docx_image_wrap` 的白名单）。
 *
 * `behind` 与 `inFront` 在 XML 里都是 `<wp:wrapNone/>`，区别只在 `behindDoc`：
 * Word 的「衬于文字下方 / 浮于文字上方」就是同一个元素的两个取值 ——
 * Word 里**没有**独立的「无环绕」选项，所以 `none` 只是 `inFront` 的别名
 * （适配器会统一归一成 `inFront` 返回，避免同一份文档出现两个名字）。
 * `tight`（紧密）与 `through`（穿越）需要 `wrapPolygon` 包围多边形，本插件不实现。
 */
const WRAP_MODES = Object.freeze({
  square: { xml: '<wp:wrapSquare wrapText="bothSides"/>', behind: false },
  topAndBottom: { xml: '<wp:wrapTopAndBottom/>', behind: false },
  none: { xml: '<wp:wrapNone/>', behind: false },
  behind: { xml: '<wp:wrapNone/>', behind: true },
  inFront: { xml: '<wp:wrapNone/>', behind: false }
})

/** `wp:anchor` 专有属性：转回 `wp:inline` 时要删掉，否则 Word 会当成未知属性。 */
const ANCHOR_ONLY_ATTRS = Object.freeze(
  new Set(['simplePos', 'relativeHeight', 'behindDoc', 'locked', 'layoutInCell', 'allowOverlap'])
)

/** `wp:positionH` 的 relativeFrom 取值。 */
const H_RELATIVE_FROM = Object.freeze({ column: 'column', page: 'page', margin: 'margin', character: 'character' })

/** `wp:positionV` 的 relativeFrom 取值。 */
const V_RELATIVE_FROM = Object.freeze({ paragraph: 'paragraph', page: 'page', margin: 'margin', line: 'line' })

/**
 * 重写元素的**起始标签与结束标签**（换标签名 + 增改属性），保留原有属性的顺序与转义。
 *
 * 结束标签必须一起改：只改起始标签会得到 `<wp:anchor>…</wp:inline>`，
 * 连自家解析器都会在重新扫描时报「XML 标签不匹配」。
 *
 * 只在「整个标签都要重写」时用：它的补丁区间是 `[node.start, node.startTagEnd)` 与
 * `[node.contentEnd, node.end)`，插在标签内部的零长属性补丁（`setAttr`）会与前者冲突，
 * 因此同一轮里不要既 `setAttr` 又调用它。
 *
 * @param {object} doc - 文档。
 * @param {object} node - 元素节点。
 * @param {string} tagName - 新标签名（含前缀，如 `wp:anchor`）。
 * @param {object} [overrides] - 要设置/覆盖的属性。
 * @param {Set<string>} [drop] - 要删除的属性名。
 * @returns {void}
 */
function rebuildStartTag(doc, node, tagName, overrides = {}, drop = new Set()) {
  const parts = []
  const used = new Set()
  for (const key of node.attrOrder ?? []) {
    if (drop.has(key)) continue
    const value = node.attrs.get(key)?.value
    if (value === undefined) continue
    parts.push(` ${key}="${value}"`)
    used.add(key)
  }
  for (const [key, value] of Object.entries(overrides)) {
    if (used.has(key)) continue
    parts.push(` ${key}="${escapeXmlAttr(String(value))}"`)
  }
  doc.patch(node.start, node.startTagEnd, `<${tagName}${parts.join('')}>`)
  if (!node.selfClosing) doc.patch(node.contentEnd, node.end, `</${tagName}>`)
}

/**
 * 读出一张图片当前的环绕方式。
 *
 * `<wp:inline>` 就是行内；`<wp:anchor>` 下再看子元素与 `behindDoc`：
 * `wrapNone` + `behindDoc=1` 是「衬于文字下方」，否则是「浮于文字上方」——
 * 这两种在 Word 里是两个选项，在 XML 里只差一个属性。
 *
 * @param {object} node - `wp:inline` / `wp:anchor` 节点。
 * @returns {string} `inline` | `square` | `topAndBottom` | `none` | `behind` | `inFront`。
 */
function wrapModeOf(node) {
  if (localName(node) === 'inline') return 'inline'
  if (find(node, 'wrapSquare')) return 'square'
  if (find(node, 'wrapTopAndBottom')) return 'topAndBottom'
  if (find(node, 'wrapTight')) return 'tight'
  if (find(node, 'wrapThrough')) return 'through'
  if (find(node, 'wrapNone')) return attr(node, 'behindDoc') === '1' ? 'behind' : 'inFront'
  return 'none'
}

/**
 * 取节点本地名（忽略命名空间前缀）。
 * @param {object} node - 元素节点。
 * @returns {string} 本地名。
 */
function localName(node) {
  return node.name.includes(':') ? node.name.slice(node.name.indexOf(':') + 1) : node.name
}

/**
 * 必须渲染成图片才能判定的检查项。
 *
 * 明确列出来而不是假装通过：读 XML 无法可靠判断空白页、元素重叠、表格越界、
 * 页眉页脚重叠与字体替换，硬猜只会产生假阳性/假阴性。
 */
const NOT_CHECKED = Object.freeze([
  { name: '空白页面', reason: '需要渲染分页后才能判定' },
  { name: '元素重叠', reason: '需要渲染后才能判定' },
  { name: '表格超出页面边界', reason: '需要按字体度量排版后才能判定' },
  { name: '页眉页脚重叠', reason: '需要渲染后才能判定' },
  { name: '字体替换', reason: '需要目标机器的字体环境才能判定；本工具改为输出文档使用的字体清单' }
])

/**
 * 读取关系文件，返回 `rId → Target`。
 * @param {ZipPackage} pkg - OPC 包。
 * @param {string} part - 关系部件路径。
 * @returns {Map<string, string>} 关系映射。
 */
function readRelationshipTargets(pkg, part) {
  const map = new Map()
  if (!pkg.has(part)) return map
  const doc = XmlDoc.parse(pkg.readText(part))
  for (const rel of findAll(doc.root, 'Relationship')) {
    const id = attr(rel, 'Id')
    const target = attr(rel, 'Target')
    if (id && target) map.set(id, target)
  }
  return map
}

/**
 * 把关系目标解析为包内绝对部件路径。
 * @param {string} baseDir - 宿主部件所在目录。
 * @param {string} target - 关系目标。
 * @returns {string} 归一化路径。
 */
function resolvePart(baseDir, target) {
  if (target.startsWith('/')) return target.slice(1)
  const stack = []
  for (const seg of `${baseDir}/${target}`.split('/')) {
    if (seg === '' || seg === '.') continue
    if (seg === '..') stack.pop()
    else stack.push(seg)
  }
  return stack.join('/')
}

/**
 * 读取 styles.xml 中定义的全部样式 ID。
 * @param {ZipPackage} pkg - OPC 包。
 * @returns {Set<string>|null} 样式 ID 集合；无 styles.xml 时返回 null。
 */
function readStyleIds(pkg) {
  if (!pkg.has('word/styles.xml')) return null
  const doc = XmlDoc.parse(pkg.readText('word/styles.xml'))
  return new Set(findAll(doc.root, 'style').map((s) => attr(s, 'styleId')).filter(Boolean))
}

/**
 * 读取 numbering.xml 中定义的全部 numId。
 * @param {ZipPackage} pkg - OPC 包。
 * @returns {Set<string>|null} numId 集合；无 numbering.xml 时返回 null。
 */
function readNumberingIds(pkg) {
  if (!pkg.has('word/numbering.xml')) return null
  const doc = XmlDoc.parse(pkg.readText('word/numbering.xml'))
  return new Set(findAll(doc.root, 'num').map((n) => attr(n, 'numId')).filter(Boolean))
}

/**
 * 把纯文本转成 WordprocessingML 的 run 序列。
 *
 * `\n` 转 `<w:br/>`（软换行）而非新段落，保证段落下标与段落属性不被破坏；
 * 含首尾空格的文本加 `xml:space="preserve"`，否则 Word 会吃掉空白。
 *
 * @param {string} text - 纯文本。
 * @returns {string} run 序列 XML；空文本返回空串。
 */
function runsXml(text) {
  const value = String(text ?? '')
  if (value === '') return ''
  return value
    .split('\n')
    .map((line) => `<w:r><w:t xml:space="preserve">${escapeXmlText(line)}</w:t></w:r>`)
    .join('<w:r><w:br/></w:r>')
}

/**
 * `w:tcPr` 子元素的 schema 顺序（ECMA-376 §17.4.70）。
 *
 * 顺序错了 Word 会报「文档内容有问题」——`gridSpan` 在 `tcW` 之后、`vMerge` 之前，
 * 不是 `tcPr` 的属性。所以插入子元素时要按这张表找位置，而不是无脑 append。
 */
const TCPR_ORDER = Object.freeze([
  'cnfStyle',
  'tcW',
  'gridSpan',
  'hMerge',
  'vMerge',
  'tcBorders',
  'shd',
  'noWrap',
  'tcMar',
  'textDirection',
  'tcFitText',
  'vAlign',
  'hideMark'
])

/**
 * 确保单元格的 `w:tcPr` 里有某个子元素（按 schema 顺序插入；没有 tcPr 就先建一个）。
 * @param {object} doc - 文档。
 * @param {object} tc - `w:tc` 节点。
 * @param {string} name - 子元素本地名。
 * @param {string} xml - 子元素 XML。
 * @returns {void}
 */
function ensureTcPrChild(doc, tc, name, xml) {
  let tcPr = find(tc, 'tcPr')
  if (!tcPr) {
    // w:tcPr 必须是 w:tc 的第一个子元素
    const firstChild = (tc.children ?? []).find((n) => n.type === 'element')
    const wrapper = `<w:tcPr>${xml}</w:tcPr>`
    if (firstChild) doc.insertBefore(firstChild, wrapper)
    else doc.appendChild(tc, wrapper)
    return
  }
  // 锚点必须在删除之前取：children 是改前快照，删完再找可能拿到刚被删掉的节点
  const children = (tcPr.children ?? []).filter((n) => n.type === 'element')
  const target = TCPR_ORDER.indexOf(name)
  const anchor = children.find((n) => localName(n) !== name && TCPR_ORDER.indexOf(localName(n)) > target)
  for (const node of children.filter((n) => localName(n) === name)) doc.remove(node)
  if (anchor) doc.insertBefore(anchor, xml)
  else doc.appendChild(tcPr, xml)
}

/**
 * 沿 parent 链往上找最近的 `w:p` 段落节点。
 * @param {object|null} node - 起始节点。
 * @returns {object|null} 段落节点或 null。
 */
function enclosingParagraph(node) {
  let current = node
  for (let depth = 0; depth < 12 && current; depth += 1) {
    if (localName(current) === 'p') return current
    current = current.parent ?? null
  }
  return null
}

/**
 * EMU → 像素（96 DPI）。
 * @param {number} emu - EMU 值。
 * @returns {number} 像素。
 */
function emuToPx(emu) {
  return Math.round(emu / EMU_PER_PX)
}

/**
 * 像素 → EMU（96 DPI）。
 * @param {number} px - 像素。
 * @returns {number} EMU。
 */
function pxToEmu(px) {
  return Math.round(px * EMU_PER_PX)
}

/**
 * 构造一个完整的 `<w:p>` 段落。
 * @param {object} args - `{text, style, alignment}`。
 * @returns {string} 段落 XML。
 */
function buildParagraphXml({ text, style, alignment }) {
  const props = []
  if (style) props.push(`<w:pStyle w:val="${escapeXmlAttr(style)}"/>`)
  if (alignment) props.push(`<w:jc w:val="${escapeXmlAttr(alignment)}"/>`)
  const pPr = props.length > 0 ? `<w:pPr>${props.join('')}</w:pPr>` : ''
  return `<w:p>${pPr}${runsXml(text)}</w:p>`
}

/**
 * 生成用于日志与变更记录的短预览，避免把正文写进日志（§十二.3）。
 * @param {string} text - 原文。
 * @param {number} [max] - 最大长度。
 * @returns {string} 预览文本。
 */
function preview(text, max = 30) {
  const value = String(text ?? '')
  return value.length > max ? `${value.slice(0, max)}…` : value
}

/**
 * 判断 haystack 是否包含 needle。
 * @param {string} haystack - 被查找文本。
 * @param {string} needle - 查找内容。
 * @param {boolean} matchCase - 是否区分大小写。
 * @returns {boolean} 是否包含。
 */
function contains(haystack, needle, matchCase) {
  return matchCase ? haystack.includes(needle) : haystack.toLowerCase().includes(needle.toLowerCase())
}

/**
 * 全量替换（非正则，避免用户输入被当作模式解释）。
 * @param {string} original - 原文本。
 * @param {string} needle - 查找内容。
 * @param {string} replacement - 替换内容。
 * @param {boolean} matchCase - 是否区分大小写。
 * @returns {string} 替换结果。
 */
function replaceAll(original, needle, replacement, matchCase) {
  if (matchCase) return original.split(needle).join(replacement)
  const escaped = needle.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return original.replace(new RegExp(escaped, 'gi'), () => replacement)
}

/**
 * 提取一个段落的可见文本。
 *
 * 跳过 `w:del`（已删除的修订）、`w:pPr`、`w:rPr`；
 * `w:t` 取文本，`w:tab` 转制表符，`w:br`/`w:cr` 转换行。
 *
 * @param {XmlDoc} doc - 文档对象。
 * @param {object} node - 段落或任意容器节点。
 * @returns {string} 文本。
 */
function paragraphText(doc, node) {
  let out = ''
  const walk = (current) => {
    for (const child of current.children ?? []) {
      if (child.type !== 'element') continue
      const name = localName(child)
      if (name === 't') out += nodeText(doc, child)
      else if (name === 'tab') out += '\t'
      else if (name === 'br' || name === 'cr') out += '\n'
      else if (name === 'del' || name === 'delText' || name === 'pPr' || name === 'rPr' || name === 'tblPr') continue
      else walk(child)
    }
  }
  walk(node)
  return out
}

/**
 * 提取段落的结构化信息。
 * @param {XmlDoc} doc - 文档对象。
 * @param {object} p - `<w:p>` 节点。
 * @param {number} index - 段落序号。
 * @returns {object} 段落信息。
 */
function paragraphInfo(doc, p, index) {
  const pPr = find(p, 'pPr')
  const styleNode = pPr ? find(pPr, 'pStyle') : null
  const outlineNode = pPr ? find(pPr, 'outlineLvl') : null
  const jcNode = pPr ? find(pPr, 'jc') : null
  const numPr = pPr ? find(pPr, 'numPr') : null
  const runs = findAll(p, 'r')
  return {
    index,
    text: paragraphText(doc, p),
    style: styleNode ? (attr(styleNode, 'val') ?? null) : null,
    outline_level: outlineNode ? Number(attr(outlineNode, 'val') ?? 0) : null,
    alignment: jcNode ? (attr(jcNode, 'val') ?? null) : null,
    is_list_item: numPr !== null,
    run_count: runs.length,
    has_bold_run: runs.some((r) => {
      const rPr = find(r, 'rPr')
      return rPr ? find(rPr, 'b') !== undefined : false
    })
  }
}

/**
 * 构造一份最小的合法 .docx，供测试与 `office_create_document`（docx）使用。
 *
 * 包含：标题段、含加粗行的正文段、带表头的表格、页眉、页脚（含页码域）、
 * 核心与应用属性。刻意覆盖结构识别需要探测的多种部件。
 *
 * @param {object} [options] - 内容选项。
 * @returns {Buffer} .docx 字节。
 */
export function buildBasicDocx({ title = '示例文档', paragraph = '本文件由 dsh-exp-office 生成。', withHeaderFooter = true } = {}) {
  const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { maxEntries: 64, maxEntryBytes: 1 << 20, maxTotalBytes: 1 << 24, maxRatio: 200 })
  const esc = (s) => String(s).replace(/[&<>]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : '&gt;'))

  const headerFooterTypes = withHeaderFooter
    ? `<Override PartName="/word/header1.xml" ContentType="${CT_HEADER}"/><Override PartName="/word/footer1.xml" ContentType="${CT_FOOTER}"/>`
    : ''
  const headerFooterRels = withHeaderFooter
    ? `<Relationship Id="rId2" Type="${NS_R}/header" Target="header1.xml"/><Relationship Id="rId3" Type="${NS_R}/footer" Target="footer1.xml"/>`
    : ''
  const headerFooterRefs = withHeaderFooter
    ? '<w:headerReference w:type="default" r:id="rId2"/><w:footerReference w:type="default" r:id="rId3"/>'
    : ''

  pkg.write(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_DOCUMENT}"/><Override PartName="/word/styles.xml" ContentType="${CT_STYLES}"/>${headerFooterTypes}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`
  )
  pkg.write(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${NS_R}/extended-properties" Target="docProps/app.xml"/></Relationships>`
  )
  pkg.write(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS_R}/styles" Target="styles.xml"/>${headerFooterRels}</Relationships>`
  )
  pkg.write(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>${esc(title)}</w:t></w:r></w:p><w:p><w:r><w:t>${esc(paragraph)}</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>加粗片段</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>月份</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>金额</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>1月</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>12000</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>正文结束。</w:t></w:r></w:p><w:sectPr>${headerFooterRefs}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr></w:body></w:document>`
  )
  if (withHeaderFooter) {
    pkg.write(
      'word/header1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:hdr xmlns:w="${NS_W}"><w:p><w:r><w:t>季度销售报告</w:t></w:r></w:p></w:hdr>`
    )
    pkg.write(
      'word/footer1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:ftr xmlns:w="${NS_W}"><w:p><w:r><w:t>第 </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t> 页</w:t></w:r></w:p></w:ftr>`
    )
  }
  pkg.write(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:styles xmlns:w="${NS_W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>`
  )
  pkg.write(
    'docProps/core.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(title)}</dc:title><dc:creator>dsh-exp-office</dc:creator><cp:lastModifiedBy>dsh-exp-office</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2026-09-21T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-21T00:00:00Z</dcterms:modified></cp:coreProperties>`
  )
  pkg.write(
    'docProps/app.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>dsh-exp-office</Application><Pages>1</Pages><Words>42</Words><Paragraphs>4</Paragraphs></Properties>`
  )
  return pkg.toBuffer()
}
