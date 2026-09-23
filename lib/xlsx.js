/**
 * XLSX 适配器：直接操作 OOXML 部件，采用最小修改模式。
 *
 * 对齐开发要求 §五（功能规划与开发注意事项）与 §十六.2（validate_workbook）：
 *   - 只改目标单元格所在的 XML 区间，图表、图片、宏、外部链接、
 *     条件格式、数据验证、透视表等部件逐字节不变。
 *   - 区分「公式本身」与「公式缓存结果」，默认不重算（FORMULA_NOT_RECALCULATED 警告）。
 *   - 正确处理 1900/1904 日期系统与日期序列值。
 *   - 修改前检测宏、外部链接、数字签名与工作表保护状态。
 *
 * @module dsh-exp-office/xlsx
 */

import zlib from 'node:zlib'
import {
  ZipPackage,
  XmlDoc,
  findAll,
  find,
  attr,
  escapeXmlAttr,
  escapeXmlText,
  decodeEntities,
  ensureContentTypeOverride,
  removeContentTypeOverride,
  addRelationshipTo,
  removeRelationshipFrom,
  DEFAULT_ZIP_LIMITS
} from './ooxml.js'
import { OfficeError } from './errors.js'

/** 工作簿主部件路径。 */
const WORKBOOK_PART = 'xl/workbook.xml'

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const REL_WORKSHEET = `${NS_REL}/worksheet`
const REL_SHARED_STRINGS = `${NS_REL}/sharedStrings`
const CT_WORKSHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
const CT_TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const NS_CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const NS_DRAWING_MAIN = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
/** 1 像素 = 9525 EMU（96 DPI）。 */
const EMU_PER_PX_LOCAL = 9525

/** 空绘图部件：新建图表/图片时先写它，再往里加锚点。 */
const EMPTY_DRAWING_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_DRAWING_MAIN}" xmlns:r="${NS_REL}"></xdr:wsDr>`

/** Excel 内置日期/时间数字格式 ID（ECMA-376 §18.8.30）。 */
const BUILTIN_DATE_FORMAT_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57])

/** 1900 日期系统的纪元（含 Excel 的 1900 闰年错误补偿）。 */
const EPOCH_1900 = Date.UTC(1899, 11, 30)
const EPOCH_1904 = Date.UTC(1904, 0, 1)

/**
 * 列字母转 1 基列号。`A` → 1，`AA` → 27。
 * @param {string} letters - 列字母。
 * @returns {number} 1 基列号。
 */
export function colToIndex(letters) {
  let n = 0
  for (const ch of letters.toUpperCase()) {
    const code = ch.charCodeAt(0)
    if (code < 65 || code > 90) throw new OfficeError('INVALID_REQUEST', `非法列标识：${letters}`)
    n = n * 26 + (code - 64)
  }
  return n
}

/**
 * 1 基列号转列字母。
 * @param {number} index - 1 基列号。
 * @returns {string} 列字母。
 */
export function indexToCol(index) {
  if (!Number.isInteger(index) || index < 1) throw new OfficeError('INVALID_REQUEST', `非法列号：${index}`)
  let n = index
  let out = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(65 + rem) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

/**
 * 解析 A1 形式的单元格引用。
 * @param {string} ref - 形如 `B12` 的引用。
 * @returns {{col: number, row: number}} 1 基行列号。
 */
export function parseRef(ref) {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(String(ref).trim())
  if (!match) throw new OfficeError('INVALID_REQUEST', `非法单元格引用：${ref}`)
  return { col: colToIndex(match[1]), row: Number(match[2]) }
}

/**
 * 由 1 基行列号生成 A1 引用。
 * @param {number} col - 列号。
 * @param {number} row - 行号。
 * @returns {string} A1 引用。
 */
export function formatRef(col, row) {
  return `${indexToCol(col)}${row}`
}

/**
 * 解析区域引用，如 `A1:C10`、`B2:B20`、`A1`。
 * @param {string} range - 区域引用。
 * @returns {{start: {col: number, row: number}, end: {col: number, row: number}}} 区域边界。
 */
export function parseRange(range) {
  const parts = String(range).trim().split(':')
  const start = parseRef(parts[0])
  const end = parts.length > 1 ? parseRef(parts[1]) : { ...start }
  return {
    start: { col: Math.min(start.col, end.col), row: Math.min(start.row, end.row) },
    end: { col: Math.max(start.col, end.col), row: Math.max(start.row, end.row) }
  }
}

/**
 * 判断某个样式索引是否为日期格式。
 * @param {number} styleIndex - cellXfs 索引。
 * @param {object} styles - 解析出的样式表。
 * @returns {boolean} 是否日期格式。
 */
function isDateFormat(styleIndex, styles) {
  if (!Number.isInteger(styleIndex) || styleIndex < 0) return false
  const xf = styles.cellXfs[styleIndex]
  if (!xf) return false
  const numFmtId = xf.numFmtId ?? 0
  if (BUILTIN_DATE_FORMAT_IDS.has(numFmtId)) return true
  const code = styles.numFmts.get(numFmtId)
  if (!code) return false
  const stripped = code.replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, '')
  return /[ymdhs]/i.test(stripped)
}

/**
 * 日期序列值转 ISO 字符串。
 * @param {number} serial - Excel 日期序列值。
 * @param {boolean} date1904 - 是否使用 1904 日期系统。
 * @returns {string} ISO 8601 字符串。
 */
export function serialToIso(serial, date1904) {
  const epoch = date1904 ? EPOCH_1904 : EPOCH_1900
  const ms = epoch + Math.round(serial * 86400000)
  return new Date(ms).toISOString()
}

/**
 * ISO 日期字符串转 Excel 日期序列值。
 * @param {string} iso - ISO 8601 字符串。
 * @param {boolean} date1904 - 是否使用 1904 日期系统。
 * @returns {number} 日期序列值。
 */
export function isoToSerial(iso, date1904) {
  const ms = Date.parse(iso)
  if (Number.isNaN(ms)) throw new OfficeError('INVALID_REQUEST', `非法日期：${iso}`)
  const epoch = date1904 ? EPOCH_1904 : EPOCH_1900
  return (ms - epoch) / 86400000
}

/** 一个已打开的工作簿。 */
export class Workbook {
  #pkg
  #docCache = new Map()
  /** 已被写回、树与文本脱节的部件，下次读取树时重建。 */
  #dirtyDocs = new Set()

  /**
   * @param {ZipPackage} pkg - 底层 OPC 包。
   * @param {object} info - 工作簿结构信息。
   */
  constructor(pkg, info) {
    this.#pkg = pkg
    this.info = info
  }

  /**
   * 打开一个 .xlsx。
   * @param {Buffer} buffer - 文件字节。
   * @param {Partial<typeof DEFAULT_ZIP_LIMITS>} [limits] - ZIP 安全上限。
   * @returns {Workbook} 工作簿对象。
   */
  static open(buffer, limits) {
    // 先按 ZIP 读取，若为 OLE 复合文档则说明是加密或旧版格式。
    if (buffer.length >= 8 && buffer.readUInt32LE(0) === 0xe011cfd0) {
      throw new OfficeError('PASSWORD_REQUIRED', '文件是加密的 OOXML（OLE 复合文档容器），需要密码或解密副本。')
    }
    const pkg = ZipPackage.open(buffer, limits)
    if (!pkg.has('xl/workbook.xml')) {
      throw new OfficeError('CORRUPTED_DOCUMENT', '缺少 xl/workbook.xml，不是有效的 .xlsx 文件。')
    }
    const wbDoc = XmlDoc.parse(pkg.readText('xl/workbook.xml'))
    const date1904 = attr(find(wbDoc.root, 'workbookPr') ?? { attrs: new Map() }, 'date1904') === '1'
    const rels = readRels(pkg, 'xl/_rels/workbook.xml.rels')
    const sheets = []
    for (const node of findAll(find(wbDoc.root, 'sheets') ?? wbDoc.root, 'sheet')) {
      const name = attr(node, 'name') ?? ''
      const rid = attr(node, 'id') ?? node.attrs.get('r:id')?.value
      const target = rels.get(rid)
      if (!target) continue
      sheets.push({
        name,
        sheetId: Number(attr(node, 'sheetId') ?? 0),
        rid,
        path: normalizePart('xl', target),
        state: attr(node, 'state') ?? 'visible'
      })
    }
    const info = {
      date1904,
      sheets,
      workbookDoc: wbDoc,
      hasMacro: pkg.has('xl/vbaProject.bin') || pkg.names().some((n) => n.endsWith('.bin') && n.includes('vba')),
      hasExternalLinks: pkg.names().some((n) => n.startsWith('xl/externalLinks/')),
      hasSignatures: pkg.names().some((n) => n.includes('_xmlsignatures/')),
      hasCharts: pkg.names().some((n) => /^xl\/charts\//.test(n)),
      hasPivot: pkg.names().some((n) => /^xl\/pivot/.test(n)),
      hasSharedStrings: pkg.has('xl/sharedStrings.xml'),
      sheetProtection: new Map()
    }
    return new Workbook(pkg, info)
  }

  /** @returns {string[]} 工作表名列表（按工作簿顺序）。 */
  sheetNames() {
    return this.info.sheets.map((s) => s.name)
  }

  /**
   * 按名称取工作表结构信息。
   * @param {string} name - 工作表名。
   * @returns {object} 工作表信息。
   */
  sheet(name) {
    const found = this.info.sheets.find((s) => s.name === name)
    if (!found) {
      throw new OfficeError('FILE_NOT_FOUND', `工作簿中不存在工作表「${name}」。`, { available: this.sheetNames() })
    }
    return found
  }

  /**
   * 取得（并缓存）某个部件的 XML 文档对象。
   *
   * 写回后树与文本会脱节，这里按需重建（懒 rescan）。不这么做的话，
   * 「多次批量写入同一张大表」每次都要全量重解析，退化成 O(n²)。
   *
   * @param {string} part - 部件路径。
   * @returns {XmlDoc} XML 文档。
   */
  #doc(part) {
    let doc = this.#docCache.get(part)
    if (!doc) {
      doc = XmlDoc.parse(this.#pkg.readText(part))
      this.#docCache.set(part, doc)
      return doc
    }
    if (this.#dirtyDocs.delete(part)) doc.rescan()
    return doc
  }

  /**
   * 把缓存中的部件写回包，并把它标记为「树已过期」。
   *
   * 补丁模型下，本次会话新增的节点只存在于输出文本里。标记而非立即重建，
   * 让「先 write_cells 再 set_cell_style」能看到刚写入的单元格，
   * 同时避免每次写入都付一次全量解析的代价。
   *
   * @param {string} part - 部件路径。
   * @returns {void}
   */
  #flush(part) {
    const doc = this.#docCache.get(part)
    if (!doc) return
    this.#pkg.write(part, doc.toString())
    this.#dirtyDocs.add(part)
  }

  /**
   * 写回一个不在 `#docCache` 中托管的文档（例如工作簿主部件）。
   * @param {string} part - 部件路径。
   * @param {XmlDoc} doc - 文档对象。
   * @returns {void}
   */
  #commitDoc(part, doc) {
    this.#pkg.write(part, doc.toString())
    this.#dirtyDocs.add(part)
  }

  /**
   * 按工作表名定位 `xl/workbook.xml` 中当前的 `<sheet>` 节点。
   *
   * 每次现查而不是缓存节点引用：rescan 之后旧引用会失效，缓存引用会导致
   * 在错误的偏移量上打补丁。
   *
   * @param {string} name - 工作表名。
   * @returns {object|null} `<sheet>` 节点。
   */
  #sheetNode(name) {
    const doc = this.info.workbookDoc
    if (this.#dirtyDocs.delete(WORKBOOK_PART)) doc.rescan()
    const sheetsNode = find(doc.root, 'sheets')
    if (!sheetsNode) return null
    return (
      (sheetsNode.children ?? []).find((n) => n.type === 'element' && localName(n) === 'sheet' && attr(n, 'name') === name) ?? null
    )
  }

  /**
   * 读取共享字符串表。
   * @returns {{items: string[], doc: XmlDoc|null, part: string|null}} 共享字符串表。
   */
  #sharedStrings() {
    if (!this.info.hasSharedStrings) return { items: [], doc: null, part: null, refs: 0, dirty: false }
    const part = 'xl/sharedStrings.xml'
    const doc = this.#doc(part)
    const root = doc.root.children.find((c) => c.type === 'element')
    const items = findAll(root, 'si').map((si) => collectText(doc, si))
    return { items, doc, part, refs: Number(attr(root, 'count') ?? items.length), dirty: false }
  }

  /**
   * 读取工作表的稀疏单元格列表。
   * @param {string} name - 工作表名。
   * @returns {{cells: object[], styles: object}} 单元格与样式表。
   */
  #readCells(name) {
    const sheet = this.sheet(name)
    const doc = this.#doc(sheet.path)
    const styles = this.#styles()
    const sst = this.#sharedStrings()
    const protection = find(doc.root, 'sheetProtection')
    this.info.sheetProtection.set(name, protection ? (attr(protection, 'sheet') === '1' || attr(protection, 'objects') === '1') : false)

    const cells = []
    for (const c of findAll(doc.root, 'c')) {
      const cell = this.#cellOf(c, doc, styles, sst)
      if (cell) cells.push(cell)
    }
    return { cells, styles }
  }

  /**
   * 把一个 `<c>` 节点解析为值对象。
   *
   * 类型判定的完整分支集中在这里，`readSheet` 与 `readRange` 共用，
   * 避免两处实现出现语义漂移。
   *
   * @param {object} c - 单元格节点。
   * @param {XmlDoc} doc - 工作表文档。
   * @param {object} styles - 样式表。
   * @param {object} sst - 共享字符串表。
   * @returns {object|null} 单元格值对象。
   */
  #cellOf(c, doc, styles, sst) {
    const ref = attr(c, 'r')
    if (!ref) return null
    const { col, row } = parseRef(ref)
    const styleIndex = Number(attr(c, 's') ?? 0)
    const type = attr(c, 't') ?? 'n'
    const fNode = find(c, 'f')
    const vNode = find(c, 'v')
    const isNode = find(c, 'is')
    const formula = fNode ? collectText(doc, fNode) : null
    let value = null
    let valueType = 'number'
    if (type === 's') {
      const idx = Number(vNode ? doc.text(vNode) : -1)
      value = sst.items[idx] ?? null
      valueType = 'string'
    } else if (type === 'inlineStr') {
      value = isNode ? collectText(doc, isNode) : null
      valueType = 'string'
    } else if (type === 'str') {
      value = vNode ? doc.text(vNode) : null
      valueType = 'string'
    } else if (type === 'b') {
      value = vNode ? doc.text(vNode) === '1' : false
      valueType = 'boolean'
    } else if (type === 'e') {
      value = vNode ? doc.text(vNode) : null
      valueType = 'error'
    } else if (vNode) {
      const num = Number(doc.text(vNode))
      if (isDateFormat(styleIndex, styles)) {
        value = serialToIso(num, this.info.date1904)
        valueType = 'date'
      } else {
        value = num
        valueType = 'number'
      }
    }
    return { ref, col, row, value, valueType, formula, style: styleIndex, raw: type }
  }

  /**
   * 读取工作表全部单元格。
   * @param {string} name - 工作表名。
   * @param {object} [options] - 选项。
   * @param {number} [options.maxCells] - 单元格上限，超出即报错（防超大表 OOM）。
   * @returns {object} 工作表内容。
   */
  readSheet(name, { maxCells = 500000 } = {}) {
    const { cells } = this.#readCells(name)
    if (cells.length > maxCells) {
      throw new OfficeError('MEMORY_LIMIT', `工作表「${name}」含 ${cells.length} 个单元格，超过上限 ${maxCells}，请改用区域读取。`)
    }
    const maxRow = cells.reduce((m, c) => Math.max(m, c.row), 0)
    const maxCol = cells.reduce((m, c) => Math.max(m, c.col), 0)
    return { sheet: name, row_count: maxRow, column_count: maxCol, cells }
  }

  /**
   * 读取指定区域，返回二维数组。
   *
   * 只遍历落在区域内的行与单元格，**不物化整张表**。
   * 这是「只读取用户要求的区域」这条性能要求的实现点：在 20 万单元格的表上
   * 读 100 行，代价与 100 行成正比，而不是与整表成正比。
   *
   * @param {string} name - 工作表名。
   * @param {string} range - 区域引用，如 `A1:C10`。
   * @returns {object} 区域内容。
   */
  readRange(name, range) {
    const bounds = parseRange(range)
    const sheet = this.sheet(name)
    const doc = this.#doc(sheet.path)
    const styles = this.#styles()
    const sst = this.#sharedStrings()
    const width = bounds.end.col - bounds.start.col + 1
    const height = bounds.end.row - bounds.start.row + 1
    const rows = Array.from({ length: height }, () => new Array(width).fill(null))

    const sheetData = find(doc.root, 'sheetData')
    if (sheetData) {
      let impliedRow = 0
      for (const rowNode of sheetData.children ?? []) {
        if (rowNode.type !== 'element' || localName(rowNode) !== 'row') continue
        const rawRow = attr(rowNode, 'r')
        const rowNumber = rawRow === undefined ? impliedRow + 1 : Number(rawRow)
        impliedRow = rowNumber
        if (rowNumber < bounds.start.row || rowNumber > bounds.end.row) continue
        for (const c of rowNode.children ?? []) {
          if (c.type !== 'element' || localName(c) !== 'c') continue
          const cell = this.#cellOf(c, doc, styles, sst)
          if (!cell) continue
          if (cell.col < bounds.start.col || cell.col > bounds.end.col) continue
          rows[rowNumber - bounds.start.row][cell.col - bounds.start.col] = {
            ref: cell.ref,
            value: cell.value,
            type: cell.valueType,
            formula: cell.formula
          }
        }
      }
    }
    return { sheet: name, range, rows }
  }

  /**
   * 解析样式表：cellXfs 索引 → numFmtId，以及自定义数字格式。
   * @returns {{cellXfs: object[], numFmts: Map<number, string>}} 样式表。
   */
  #styles() {
    if (this.#styleCache) return this.#styleCache
    const empty = { cellXfs: [], numFmts: new Map() }
    if (!this.#pkg.has('xl/styles.xml')) {
      this.#styleCache = empty
      return empty
    }
    const doc = this.#doc('xl/styles.xml')
    const numFmts = new Map()
    for (const nf of findAll(doc.root, 'numFmt')) {
      numFmts.set(Number(attr(nf, 'numFmtId')), attr(nf, 'formatCode') ?? '')
    }
    const cellXfsNode = find(doc.root, 'cellXfs')
    const cellXfs = cellXfsNode
      ? (cellXfsNode.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'xf').map((xf) => ({ numFmtId: Number(attr(xf, 'numFmtId') ?? 0) }))
      : []
    this.#styleCache = { cellXfs, numFmts }
    return this.#styleCache
  }

  #styleCache = null

  /**
   * 批量写入单元格。
   *
   * 采用「整格替换」：保留原有 `s` 样式索引与 `r` 引用，只重建该 `<c>` 元素，
   * 因此同工作表内其它单元格、图表、图片等部件保持逐字节不变。
   *
   * @param {string} name - 工作表名。
   * @param {object[]} edits - 编辑列表，每项 `{ref, value, formula, type}`。
   * @returns {object[]} 变更记录。
   */
  writeCells(name, edits) {
    const sheet = this.sheet(name)
    const doc = this.#doc(sheet.path)
    const sst = this.#sharedStrings()
    const sheetData = find(doc.root, 'sheetData')
    if (!sheetData) throw new OfficeError('CORRUPTED_DOCUMENT', `工作表「${name}」缺少 sheetData 节点。`)

    // 行号 → row 节点，以及 (列:行) → c 节点索引，避免逐格全表扫描。
    const rowNodes = new Map()
    const cellIndex = new Map()
    for (const row of findAll(sheetData, 'row')) {
      const r = Number(attr(row, 'r') ?? 0)
      if (r <= 0) continue
      rowNodes.set(r, row)
      for (const c of row.children ?? []) {
        if (c.type !== 'element' || localName(c) !== 'c') continue
        const ref = attr(c, 'r')
        if (ref) cellIndex.set(`${parseRef(ref).col}:${r}`, c)
      }
    }

    // 按行分组：新建的行一次性插入完整 <row>，避免「先插空行再补单元格」的定位问题。
    const byRow = new Map()
    for (const edit of edits) {
      const { col, row } = edit.ref ? parseRef(edit.ref) : { col: edit.col, row: edit.row }
      if (!byRow.has(row)) byRow.set(row, [])
      byRow.get(row).push({ col, edit })
    }

    const changes = []
    const stats = { sstRefs: 0 }
    // 新行插入锚点用二分查找：若每建一行都全扫 rowNodes，批量建表会退化成 O(n²)。
    const existingRowNumbers = [...rowNodes.keys()].sort((a, b) => a - b)
    const anchorRowFor = (row) => {
      let lo = 0
      let hi = existingRowNumbers.length
      while (lo < hi) {
        const mid = (lo + hi) >> 1
        if (existingRowNumbers[mid] <= row) lo = mid + 1
        else hi = mid
      }
      return lo < existingRowNumbers.length ? rowNodes.get(existingRowNumbers[lo]) : undefined
    }
    for (const row of [...byRow.keys()].sort((a, b) => a - b)) {
      // 值为 null 表示清空单元格：删除节点，而不是留下一个空 `<c>`。
      const items = byRow
        .get(row)
        .sort((a, b) => a.col - b.col)
        .map((item) => ({ ...item, clear: !item.edit.formula && item.edit.value === null }))
      const rowNode = rowNodes.get(row)
      const toWrite = items.filter((i) => !i.clear)

      if (!rowNode) {
        for (const { col } of items.filter((i) => i.clear)) {
          changes.push({ type: 'cell_clear', sheet: name, range: formatRef(col, row), formula: null })
        }
        if (toWrite.length > 0) {
          const xml = `<row r="${row}">${toWrite.map(({ col, edit }) => buildCellXml(formatRef(col, row), edit, edit.style ?? 0, sst, stats)).join('')}</row>`
          const anchor = anchorRowFor(row)
          if (anchor) doc.insertBefore(anchor, xml)
          else doc.appendChild(sheetData, xml)
          for (const { col, edit } of toWrite) {
            changes.push({ type: 'cell_update', sheet: name, range: formatRef(col, row), formula: edit.formula ?? null })
          }
        }
        continue
      }

      for (const { col, clear, edit } of items) {
        const ref = formatRef(col, row)
        const existing = cellIndex.get(`${col}:${row}`)
        if (clear) {
          if (existing) {
            doc.remove(existing)
            changes.push({ type: 'cell_clear', sheet: name, range: ref, formula: null })
          }
          continue
        }
        const styleIndex = edit.style ?? (existing ? Number(attr(existing, 's') ?? 0) : 0)
        const cellXml = buildCellXml(ref, edit, styleIndex, sst, stats)
        if (existing) {
          doc.replace(existing, cellXml)
        } else {
          const anchorCell = this.#firstCellAfter(doc, rowNode, col)
          if (anchorCell) doc.insertBefore(anchorCell, cellXml)
          else doc.appendChild(rowNode, cellXml)
        }
        changes.push({ type: 'cell_update', sheet: name, range: ref, formula: edit.formula ?? null })
      }
    }

    bumpSharedStringRefs(sst, stats.sstRefs)
    this.#flushSst(sst)
    this.#flushStyles()
    this.#flush(sheet.path)
    return changes
  }

  /**
   * 找到行内第一个列号大于目标的单元格，用于保持列序。
   * @param {XmlDoc} doc - 工作表文档。
   * @param {object} rowNode - 行节点。
   * @param {number} col - 目标列号。
   * @returns {object|undefined} 单元格节点。
   */
  #firstCellAfter(doc, rowNode, col) {
    for (const c of rowNode.children ?? []) {
      if (c.type !== 'element' || localName(c) !== 'c') continue
      const ref = attr(c, 'r')
      if (ref && parseRef(ref).col > col) return c
    }
    return undefined
  }

  /**
   * 把共享字符串表变更写回包。
   * @param {object} sst - 共享字符串表。
   * @returns {void}
   */
  #flushSst(sst) {
    if (sst.part && sst.dirty) this.#flush(sst.part)
  }

  /**
   * 把样式表变更写回包。
   * @returns {void}
   */
  #flushStyles() {
    if (this.#pkg.has('xl/styles.xml') && this.#docCache.has('xl/styles.xml')) this.#flush('xl/styles.xml')
  }

  /**
   * 在当前工作簿上执行查找替换。
   *
   * 只改写命中的单元格文本；公式单元格默认跳过（避免破坏公式语义）。
   *
   * @param {object} args - 参数。
   * @returns {object} 变更统计。
   */
  findAndReplace({ find: needle, replace: replacement, sheet: sheetName = null, matchCase = false, wholeCell = false, includeFormulas = false, maxReplacements = 100000 }) {
    if (typeof needle !== 'string' || needle === '') {
      throw new OfficeError('INVALID_REQUEST', '查找内容不能为空。')
    }
    const targets = sheetName ? [this.sheet(sheetName)] : this.info.sheets
    const changes = []
    let count = 0
    for (const sheet of targets) {
      const doc = this.#doc(sheet.path)
      const sst = this.#sharedStrings()
      for (const c of findAll(doc.root, 'c')) {
        if (count >= maxReplacements) break
        const ref = attr(c, 'r')
        if (!ref) continue
        const fNode = find(c, 'f')
        if (fNode && !includeFormulas) continue
        const type = attr(c, 't') ?? 'n'
        let target = null
        let kind = null
        if (type === 's') {
          const vNode = find(c, 'v')
          if (!vNode) continue
          const idx = Number(doc.text(vNode))
          const original = sst.items[idx]
          if (typeof original !== 'string') continue
          const updated = applyReplace(original, needle, replacement, matchCase, wholeCell)
          if (updated === original) continue
          const newIdx = internSharedString(sst, updated)
          doc.setText(vNode, String(newIdx))
          target = original
          kind = 'shared'
        } else if (type === 'inlineStr') {
          const isNode = find(c, 'is')
          const tNode = isNode ? find(isNode, 't') : null
          if (!tNode) continue
          const original = doc.text(tNode)
          const updated = applyReplace(original, needle, replacement, matchCase, wholeCell)
          if (updated === original) continue
          doc.setText(tNode, updated)
          target = original
          kind = 'inline'
        } else if (type === 'str') {
          // 部分写入器（如 SheetJS）把普通字符串也存成 t="str" + <v>文本</v>。
          const vNode = find(c, 'v')
          if (!vNode) continue
          const original = doc.text(vNode)
          const updated = applyReplace(original, needle, replacement, matchCase, wholeCell)
          if (updated === original) continue
          doc.setText(vNode, updated)
          target = original
          kind = 'literal'
        } else if (fNode && includeFormulas) {
          const original = collectText(doc, fNode)
          const updated = applyReplace(original, needle, replacement, matchCase, wholeCell)
          if (updated === original) continue
          doc.setText(fNode, updated)
          target = original
          kind = 'formula'
        }
        if (target !== null) {
          count += 1
          changes.push({ type: 'replace', sheet: sheet.name, range: ref, kind, from: target, to: replacement })
        }
      }
      if (sst.dirty) this.#flush(sst.part)
      if (changes.some((c) => c.sheet === sheet.name)) this.#flush(sheet.path)
    }
    return { replacements: count, changes }
  }

  /**
   * 新增工作表。
   * @param {string} name - 新工作表名。
   * @param {object} [options] - 选项。
   * @returns {object} 新建信息。
   */
  addWorksheet(name, { index = null } = {}) {
    if (this.info.sheets.some((s) => s.name === name)) {
      throw new OfficeError('INVALID_REQUEST', `工作表「${name}」已存在。`)
    }
    if (name.length > 31 || /[\\/?*[\]:]/.test(name)) {
      throw new OfficeError('INVALID_REQUEST', `工作表名「${name}」非法（不超过 31 字符且不得含 \\ / ? * [ ] :）。`)
    }
    const usedIds = new Set(this.info.sheets.map((s) => s.sheetId))
    let sheetId = 1
    while (usedIds.has(sheetId)) sheetId += 1

    // 找一个未被占用的部件路径
    let n = 1
    while (this.#pkg.has(`xl/worksheets/sheet${n}.xml`)) n += 1
    const part = `xl/worksheets/sheet${n}.xml`

    this.#pkg.write(part, EMPTY_SHEET_XML)
    ensureContentTypeOverride(this.#pkg, `/${part}`, CT_WORKSHEET)
    // 用 addRelationshipTo 的返回值作为 r:id —— 不要另外算一个 ID：
    // 两处各算一次，一旦算法分叉，workbook.xml 的引用就会指向另一条关系。
    const rid = addRelationshipTo(this.#pkg, 'xl/_rels/workbook.xml.rels', REL_WORKSHEET, `worksheets/sheet${n}.xml`)

    const wbDoc = this.info.workbookDoc
    const sheetsNode = find(wbDoc.root, 'sheets')
    if (!sheetsNode) throw new OfficeError('CORRUPTED_DOCUMENT', 'workbook.xml 缺少 sheets 节点。')
    const xml = `<sheet name="${escapeXmlAttr(name)}" sheetId="${sheetId}" r:id="${rid}"/>`
    const before = index === null ? null : this.info.sheets[index]
    const beforeNode = before ? this.#sheetNode(before.name) : null
    if (beforeNode) wbDoc.insertBefore(beforeNode, xml)
    else wbDoc.appendChild(sheetsNode, xml)
    this.#commitDoc(WORKBOOK_PART, wbDoc)

    const entry = { name, sheetId, rid, path: part, state: 'visible' }
    if (before) this.info.sheets.splice(this.info.sheets.indexOf(before), 0, entry)
    else this.info.sheets.push(entry)
    return { sheet: name, path: part, index: this.info.sheets.indexOf(entry) }
  }

  /**
   * 删除工作表。至少保留一张可见工作表。
   * @param {string} name - 工作表名。
   * @returns {object} 删除信息。
   */
  deleteWorksheet(name) {
    const sheet = this.sheet(name)
    const remaining = this.info.sheets.filter((s) => s !== sheet)
    if (remaining.filter((s) => s.state === 'visible').length === 0) {
      throw new OfficeError('INVALID_REQUEST', '不能删除最后一张可见工作表，工作簿至少需要保留一张。')
    }
    const node = this.#sheetNode(name)
    if (node) this.info.workbookDoc.remove(node)
    this.#commitDoc(WORKBOOK_PART, this.info.workbookDoc)
    removeRelationshipFrom(this.#pkg, 'xl/_rels/workbook.xml.rels', sheet.rid)
    removeContentTypeOverride(this.#pkg, `/${sheet.path}`)
    this.#pkg.delete(sheet.path)
    this.#docCache.delete(sheet.path)
    this.info.sheets = remaining
    return { sheet: name, removed_path: sheet.path }
  }

  /**
   * 重命名工作表。
   * @param {string} from - 原名称。
   * @param {string} to - 新名称。
   * @returns {object} 变更信息。
   */
  renameWorksheet(from, to) {
    const sheet = this.sheet(from)
    if (this.info.sheets.some((s) => s.name === to)) {
      throw new OfficeError('INVALID_REQUEST', `工作表「${to}」已存在。`)
    }
    if (to.length > 31 || /[\\/?*[\]:]/.test(to)) {
      throw new OfficeError('INVALID_REQUEST', `工作表名「${to}」非法。`)
    }
    const node = this.#sheetNode(from)
    if (!node) throw new OfficeError('CORRUPTED_DOCUMENT', `workbook.xml 中找不到工作表「${from}」的声明。`)
    this.info.workbookDoc.setAttr(node, 'name', to)
    this.#commitDoc(WORKBOOK_PART, this.info.workbookDoc)
    sheet.name = to
    return { from, to }
  }

  /**
   * 合并单元格区域。
   * @param {string} name - 工作表名。
   * @param {string} range - 区域引用。
   * @returns {object} 变更信息。
   */
  mergeCells(name, range) {
    const sheet = this.sheet(name)
    const doc = this.#doc(sheet.path)
    const bounds = parseRange(range)
    const existing = find(doc.root, 'mergeCells')
    const xml = `<mergeCell ref="${escapeXmlAttr(range)}"/>`
    if (existing) {
      const duplicate = findAll(existing, 'mergeCell').some((m) => attr(m, 'ref') === range)
      if (duplicate) throw new OfficeError('INVALID_REQUEST', `区域 ${range} 已被合并。`)
      doc.appendChild(existing, xml)
      doc.setAttr(existing, 'count', String(Number(attr(existing, 'count') ?? 1) + 1))
    } else {
      const after = find(doc.root, 'sheetData')
      if (!after) throw new OfficeError('CORRUPTED_DOCUMENT', '工作表缺少 sheetData。')
      doc.insertAfter(after, `<mergeCells count="1">${xml}</mergeCells>`)
    }
    this.#flush(sheet.path)
    return { sheet: name, range, merged: `${bounds.start.row}:${bounds.end.row}` }
  }

  /**
   * 列出工作簿里的表格对象（ListObject）。
   * @returns {object[]} 表格清单，按 id 排序。
   */
  tables() {
    const out = []
    for (const part of this.#pkg.names().filter((name) => /^xl\/tables\/[^/]+\.xml$/.test(name))) {
      try {
        const doc = XmlDoc.parse(this.#pkg.readText(part))
        const root = doc.root.children.find((c) => c.type === 'element')
        const styleNode = find(root, 'tableStyleInfo')
        out.push({
          part,
          id: Number(attr(root, 'id') ?? 0),
          name: attr(root, 'name') ?? '',
          displayName: attr(root, 'displayName') ?? '',
          ref: attr(root, 'ref') ?? '',
          headerRowCount: Number(attr(root, 'headerRowCount') ?? 1),
          totalsRowCount: Number(attr(root, 'totalsRowCount') ?? 0),
          columns: findAll(root, 'tableColumn').map((c) => attr(c, 'name') ?? ''),
          style: styleNode ? attr(styleNode, 'name') : null
        })
      } catch {
        // 单个表格部件坏了不该让整本工作簿打不开：跳过，由 validate 报告
      }
    }
    return out.sort((a, b) => a.id - b.id)
  }

  /**
   * 取某张工作表引用的表格部件路径。
   * @param {object} sheet - 工作表信息。
   * @returns {string[]} 部件路径。
   */
  #tablePartsOf(sheet) {
    const relsPath = `xl/worksheets/_rels/${sheet.path.split('/').pop()}.rels`
    if (!this.#pkg.has(relsPath)) return []
    const rels = readRels(this.#pkg, relsPath)
    const doc = this.#doc(sheet.path)
    const parts = []
    for (const node of findAll(doc.root, 'tablePart')) {
      const rid = attr(node, 'id') ?? node.attrs.get('r:id')?.value
      const target = rels.get(rid)
      if (target) parts.push(normalizePart('xl/worksheets', target))
    }
    return parts
  }

  /**
   * 创建 Excel 表格对象（ListObject）。
   *
   * 一个合法的表格至少要写四处，少任何一处 Excel 都会当作损坏文件要求修复：
   *   1. `xl/tables/tableN.xml` 本体（列名、区域、样式）；
   *   2. 工作表里的 `<tableParts><tablePart r:id="…"/></tableParts>`（且必须排在 `extLst` 之前）；
   *   3. 工作表关系部件里指向表格部件的 relationship；
   *   4. `[Content_Types].xml` 里表格部件的 Override。
   *
   * 另外两条硬约束：**列名必须非空且互不重复**，且应与表头单元格的值一致 ——
   * 表头为空或是重复值时，这里会把生成的列名**写回表头单元格**（Excel 自己建表也是这个行为），
   * 而不是留下一个「表头空着但表格说叫列1」的别扭状态。
   *
   * @param {string} name - 工作表名。
   * @param {object} args - 参数。
   * @param {string} args.ref - 表格区域（含表头行；`show_totals` 为 true 时也要含汇总行）。
   * @param {string} [args.name] - 表格名（工作簿内唯一，默认 `表1`/`表2`…）。
   * @param {string} [args.style] - 表格样式名，默认 `TableStyleMedium2`。
   * @param {boolean} [args.showTotals] - 是否含汇总行（`ref` 需已包含该行）。
   * @param {boolean} [args.hasHeader] - 是否含表头行，默认 true。
   * @returns {object} 表格信息与表头修正。
   */
  createTable(name, { ref, name: tableName = null, style = 'TableStyleMedium2', showTotals = false, hasHeader = true } = {}) {
    const sheet = this.sheet(name)
    if (typeof ref !== 'string' || !ref.includes(':')) {
      throw new OfficeError('INVALID_REQUEST', 'ref 必须是完整区域（如 A1:D5），不能是单个单元格。')
    }
    const bounds = parseRange(ref)
    const columnCount = bounds.end.col - bounds.start.col + 1
    const rowCount = bounds.end.row - bounds.start.row + 1
    if (columnCount < 1 || rowCount < 1) {
      throw new OfficeError('INVALID_REQUEST', `区域 ${ref} 无效。`)
    }
    if (showTotals && !hasHeader) {
      throw new OfficeError('INVALID_REQUEST', '汇总行需要表头行（Excel 不支持无表头 + 汇总行的组合）。')
    }
    if (showTotals && rowCount < 3) {
      throw new OfficeError('INVALID_REQUEST', '含汇总行时区域至少要有「表头 + 1 行数据 + 汇总行」。')
    }
    if (hasHeader && !showTotals && rowCount < 2) {
      throw new OfficeError('INVALID_REQUEST', '含表头时区域至少要有「表头 + 1 行数据」。')
    }
    if (!/^TableStyle[A-Za-z0-9]+$/.test(style)) {
      throw new OfficeError('INVALID_REQUEST', `样式名不合法：${style}（应形如 TableStyleMedium2）。`)
    }

    const existing = this.tables()
    const displayName = tableName ?? `表${existing.length + 1}`
    if (!/^[A-Za-z_\u4e00-\u9fff\\][A-Za-z0-9_.\u4e00-\u9fff\\]*$/.test(displayName)) {
      throw new OfficeError('INVALID_REQUEST', `表格名不合法：${displayName}（不能含空格，需以字母/下划线/汉字开头）。`)
    }
    if (/^[A-Za-z]{1,3}\d+$/.test(displayName)) {
      throw new OfficeError('INVALID_REQUEST', `表格名不能与单元格引用同形：${displayName}。`)
    }
    if (existing.some((t) => t.displayName.toLowerCase() === displayName.toLowerCase())) {
      throw new OfficeError('INVALID_REQUEST', `表格名已存在：${displayName}。`, { existing: existing.map((t) => t.displayName) })
    }

    // 同表内不允许区域重叠（Excel 的规定）
    const sheetTableParts = new Set(this.#tablePartsOf(sheet))
    for (const table of existing) {
      if (!sheetTableParts.has(table.part)) continue
      const other = parseRange(table.ref)
      const overlaps =
        bounds.start.row <= other.end.row &&
        bounds.end.row >= other.start.row &&
        bounds.start.col <= other.end.col &&
        bounds.end.col >= other.start.col
      if (overlaps) {
        throw new OfficeError('INVALID_REQUEST', `区域 ${ref} 与已有表格「${table.displayName}」(${table.ref}) 重叠。`)
      }
    }

    // 列名：取表头行的值；空/重复时生成并**写回表头单元格**
    const columnNames = []
    const headerFixes = []
    if (hasHeader) {
      const headerRange = `${formatRef(bounds.start.col, bounds.start.row)}:${formatRef(bounds.end.col, bounds.start.row)}`
      const header = this.readRange(name, headerRange).rows[0]
      for (let i = 0; i < columnCount; i += 1) {
        const cell = header[i]
        const raw = cell?.value
        let columnName = raw === null || raw === undefined ? '' : String(raw).trim()
        if (columnName === '' || columnNames.includes(columnName)) {
          let generated = `列${i + 1}`
          let suffix = 2
          while (columnNames.includes(generated)) {
            generated = `列${i + 1}_${suffix}`
            suffix += 1
          }
          headerFixes.push({ ref: formatRef(bounds.start.col + i, bounds.start.row), name: generated })
          columnName = generated
        }
        columnNames.push(columnName)
      }
    } else {
      for (let i = 0; i < columnCount; i += 1) columnNames.push(`列${i + 1}`)
    }
    if (headerFixes.length > 0) {
      this.writeCells(name, headerFixes.map((fix) => ({ ref: fix.ref, value: fix.name })))
    }

    const id = existing.reduce((max, table) => Math.max(max, table.id), 0) + 1
    const part = `xl/tables/table${id}.xml`
    const columnsXml = columnNames.map((columnName, index) => `<tableColumn id="${index + 1}" name="${escapeXmlAttr(columnName)}"/>`).join('')
    const attrs = [
      `id="${id}"`,
      `name="${escapeXmlAttr(displayName)}"`,
      `displayName="${escapeXmlAttr(displayName)}"`,
      `ref="${escapeXmlAttr(ref)}"`,
      `headerRowCount="${hasHeader ? 1 : 0}"`
    ]
    if (showTotals) attrs.push('totalsRowCount="1"')
    const tableXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
      `<table xmlns="${NS_MAIN}" ${attrs.join(' ')}>` +
      (hasHeader ? `<autoFilter ref="${escapeXmlAttr(ref)}"/>` : '') +
      `<tableColumns count="${columnNames.length}">${columnsXml}</tableColumns>` +
      `<tableStyleInfo name="${escapeXmlAttr(style)}" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>` +
      `</table>`
    this.#pkg.write(part, tableXml)
    ensureContentTypeOverride(this.#pkg, `/${part}`, CT_TABLE)

    const relsPath = `xl/worksheets/_rels/${sheet.path.split('/').pop()}.rels`
    const rid = addRelationshipTo(this.#pkg, relsPath, `${NS_REL}/table`, `../tables/table${id}.xml`)

    const doc = this.#doc(sheet.path)
    const tableParts = find(doc.root, 'tableParts')
    const partXml = `<tablePart r:id="${rid}"/>`
    if (tableParts) {
      doc.appendChild(tableParts, partXml)
      doc.setAttr(tableParts, 'count', String(Number(attr(tableParts, 'count') ?? 0) + 1))
    } else {
      // schema 顺序要求 tableParts 在 drawing 之后、extLst 之前。
      // 注意：`doc.root` 是**文档节点**，往它 appendChild 会写出第二个根元素 ——
      // 我们自己的解析器宽松能读，Excel 会直接判定文件损坏。必须挂到 worksheet 元素上。
      const wsRoot = doc.root.children.find((c) => c.type === 'element')
      if (!wsRoot) throw new OfficeError('CORRUPTED_DOCUMENT', `工作表 ${sheet.path} 没有根元素。`)
      const xml = `<tableParts count="1">${partXml}</tableParts>`
      const extLst = find(wsRoot, 'extLst')
      if (extLst) doc.insertBefore(extLst, xml)
      else doc.appendChild(wsRoot, xml)
    }
    this.#flush(sheet.path)

    return {
      sheet: name,
      display_name: displayName,
      ref,
      columns: columnNames,
      style,
      show_totals: showTotals,
      has_header: hasHeader,
      part,
      rid,
      header_fixes: headerFixes
    }
  }

  /**
   * 创建图表（柱/条形/折线/饼图）。
   *
   * 一个图表要写齐五处，少任何一处 Excel 都会要求修复文件：
   *   1. `xl/charts/chartN.xml` 图表本体；
   *   2. `xl/drawings/drawingN.xml` 里的锚点（graphicFrame 引用图表部件）；
   *   3. 工作表的 `<drawing r:id="…"/>`；
   *   4. 工作表关系 → 绘图部件，绘图部件关系 → 图表部件；
   *   5. `[Content_Types].xml` 里图表与绘图两个部件的 Override。
   *
   * 工作表已有绘图（图片或别的图表）时**复用**它并追加一个锚点，而不是新建第二个 drawing ——
   * 一张工作表只能有一个 `<drawing>` 引用。
   *
   * 同时写入分类与数值的**缓存值**（`c:cat` / `c:val` 里的 strCache/numCache）：
   * 规范里缓存可省略，但省略后没有任何阅读器能立刻画出图，得先重算；我们本来就拿得到数据。
   *
   * @param {string} name - 工作表名（放图表的那张表）。
   * @param {object} args - 参数。
   * @param {string} args.type - `column` / `bar` / `line` / `pie`。
   * @param {object[]} args.series - 数据系列：`{values: 'D2:D4', name?: '金额', nameRef?: 'D1'}`。
   * @param {string} [args.categories] - 分类区域，如 `A2:A4`；饼图也用它做标签。
   * @param {string} [args.title] - 图表标题。
   * @param {string} [args.anchor] - 左上角锚点单元格，默认 `F2`。
   * @param {number} [args.widthPx] - 宽（像素），默认 480。
   * @param {number} [args.heightPx] - 高（像素），默认 300。
   * @param {string} [args.dataSheet] - 数据所在工作表；默认与图表同一张。
   * @returns {object} 图表信息。
   */
  createChart(name, { type, series, categories = null, title = null, anchor = 'F2', widthPx = 480, heightPx = 300, dataSheet = null } = {}) {
    const sheet = this.sheet(name)
    const CHART_TYPES = { column: 'bar', bar: 'bar', line: 'line', pie: 'pie' }
    if (!CHART_TYPES[type]) {
      throw new OfficeError('INVALID_REQUEST', `不支持的图表类型：${type}（可用 column / bar / line / pie）。`)
    }
    if (!Array.isArray(series) || series.length === 0) {
      throw new OfficeError('INVALID_REQUEST', 'series 不能为空。')
    }
    if (type === 'pie' && series.length > 1) {
      throw new OfficeError('INVALID_REQUEST', '饼图只支持一个数据系列。')
    }
    const sourceSheet = dataSheet ? this.sheet(dataSheet).name : name

    // 1) 图表本体
    const chartParts = this.#pkg.names().filter((n) => /^xl\/charts\/chart\d+\.xml$/.test(n))
    const chartIndex = chartParts.length + 1
    const chartPart = `xl/charts/chart${chartIndex}.xml`
    const seriesXml = series
      .map((item, index) => this.#chartSeriesXml(sourceSheet, item, categories, index))
      .join('')
    const chartXml = buildChartXml({ kind: CHART_TYPES[type], type, seriesXml, title, seriesCount: series.length })
    this.#pkg.write(chartPart, chartXml)
    ensureContentTypeOverride(this.#pkg, `/${chartPart}`, CT_CHART)

    // 2) 绘图部件：已有就复用，否则新建并接线到工作表
    const sheetRelsPath = `xl/worksheets/_rels/${sheet.path.split('/').pop()}.rels`
    const sheetRels = readRels(this.#pkg, sheetRelsPath)
    let drawingPart = null
    for (const [rid, target] of sheetRels) {
      void rid
      if (/drawings\/drawing\d+\.xml$/.test(target)) {
        drawingPart = normalizePart('xl/worksheets', target)
        break
      }
    }
    const doc = this.#doc(sheet.path)
    if (!drawingPart) {
      const drawingIndex = this.#pkg.names().filter((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n)).length + 1
      drawingPart = `xl/drawings/drawing${drawingIndex}.xml`
      this.#pkg.write(drawingPart, EMPTY_DRAWING_XML)
      ensureContentTypeOverride(this.#pkg, `/${drawingPart}`, CT_DRAWING)
      const rid = addRelationshipTo(this.#pkg, sheetRelsPath, `${NS_REL}/drawing`, `../drawings/${drawingPart.split('/').pop()}`)
      const wsRoot = doc.root.children.find((c) => c.type === 'element')
      // `<drawing>` 在 schema 顺序里排在 tableParts 之前、pageSetup 之后
      const before = find(wsRoot, 'tableParts') ?? find(wsRoot, 'extLst')
      const xml = `<drawing r:id="${rid}"/>`
      if (before) doc.insertBefore(before, xml)
      else doc.appendChild(wsRoot, xml)
    }

    const drawingRelsPath = `xl/drawings/_rels/${drawingPart.split('/').pop()}.rels`
    const chartRid = addRelationshipTo(this.#pkg, drawingRelsPath, `${NS_REL}/chart`, `../charts/${chartPart.split('/').pop()}`)
    const drawingDoc = this.#doc(drawingPart)
    const drawingRoot = drawingDoc.root.children.find((c) => c.type === 'element')
    const existingAnchors = (drawingRoot.children ?? []).filter((c) => c.type === 'element' && localName(c).endsWith('Anchor')).length
    drawingDoc.appendChild(
      drawingRoot,
      buildChartAnchorXml({ anchor, widthPx, heightPx, rid: chartRid, index: existingAnchors + 1, metrics: this.#sheetMetrics(sheet) })
    )
    this.#flush(drawingPart)
    this.#flush(sheet.path)

    return {
      sheet: name,
      data_sheet: sourceSheet,
      type,
      chart_part: chartPart,
      drawing_part: drawingPart,
      drawing_rid: chartRid,
      anchor,
      width_px: widthPx,
      height_px: heightPx,
      series_count: series.length,
      title
    }
  }

  /**
   * 读取工作表列宽/行高（像素），用于把「目标像素尺寸」换算成准确的锚点落点。
   *
   * 换算按 Excel 自己的规则：列宽以字符为单位（Calibri 11 下 1 字符 ≈ 7px，再加 5px 内边距），
   * 行高以磅为单位（1pt = 96/72 px）。不读这些参数、直接按默认值估算，
   * 图表尺寸会偏（实测 480×300 会落成 ~563×294px）。
   *
   * @param {object} sheet - 工作表信息。
   * @returns {{columnPx: (index: number) => number, rowPx: (index: number) => number}} 度量函数。
   */
  #sheetMetrics(sheet) {
    const doc = this.#doc(sheet.path)
    const fmt = find(doc.root, 'sheetFormatPr')
    const defaultColWidth = Number(attr(fmt ?? { attrs: new Map() }, 'defaultColWidth') ?? 8.43)
    const defaultRowHeight = Number(attr(fmt ?? { attrs: new Map() }, 'defaultRowHeight') ?? 15)
    const overrides = new Map()
    const colsNode = find(doc.root, 'cols')
    if (colsNode) {
      for (const col of findAll(colsNode, 'col')) {
        const min = Number(attr(col, 'min') ?? 0)
        const max = Number(attr(col, 'max') ?? 0)
        const width = Number(attr(col, 'width') ?? defaultColWidth)
        for (let index = min; index <= max; index += 1) overrides.set(index - 1, width)
      }
    }
    return {
      columnPx: (index) => Math.round((overrides.get(index) ?? defaultColWidth) * 7 + 5),
      rowPx: () => Math.round((defaultRowHeight * 96) / 72)
    }
  }

  /**
   * 生成一个数据系列的 XML（含引用公式与缓存值）。
   *
   * @param {string} sheetName - 数据所在工作表名。
   * @param {object} item - `{values, name?, nameRef?}`。
   * @param {string|null} categories - 分类区域。
   * @param {number} index - 系列序号。
   * @returns {string} `<c:ser>` XML。
   */
  #chartSeriesXml(sheetName, item, categories, index) {
    if (!item || typeof item.values !== 'string' || !item.values.includes(':')) {
      throw new OfficeError('INVALID_REQUEST', `series[${index}].values 必须是区域（如 D2:D4）。`)
    }
    const valuesRange = this.#resolveRangeRef(sheetName, item.values)
    const values = this.readRange(valuesRange.sheet, valuesRange.range).rows.flat()
    const numbers = values.map((cell) => (typeof cell?.value === 'number' ? cell.value : Number(cell?.value))).map((n) => (Number.isFinite(n) ? n : 0))

    let titleXml = `<c:tx><c:v>${escapeXmlText(item.name ?? `系列${index + 1}`)}</c:v></c:tx>`
    if (typeof item.nameRef === 'string') {
      const nameRange = this.#resolveRangeRef(sheetName, item.nameRef)
      const nameCell = this.readRange(nameRange.sheet, nameRange.range).rows[0]?.[0]
      const nameText = nameCell?.value === null || nameCell?.value === undefined ? '' : String(nameCell.value)
      titleXml = `<c:tx><c:strRef><c:f>${escapeXmlText(nameRange.formula)}</c:f>${strCacheXml([nameText])}</c:strRef></c:tx>`
    }

    let catXml = ''
    if (typeof categories === 'string') {
      const catRange = this.#resolveRangeRef(sheetName, categories)
      const labels = this.readRange(catRange.sheet, catRange.range).rows.flat().map((cell) => (cell?.value === null || cell?.value === undefined ? '' : String(cell.value)))
      catXml = `<c:cat><c:strRef><c:f>${escapeXmlText(catRange.formula)}</c:f>${strCacheXml(labels)}</c:strRef></c:cat>`
    }
    return (
      `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>` +
      titleXml +
      catXml +
      `<c:val><c:numRef><c:f>${escapeXmlText(valuesRange.formula)}</c:f>${numCacheXml(numbers)}</c:numRef></c:val>` +
      `</c:ser>`
    )
  }

  /**
   * 把 `A2:A4` 或 `表名!A2:A4` 解析成工作表名 + 区域 + 公式文本（供图表引用）。
   * @param {string} defaultSheet - 默认工作表名。
   * @param {string} ref - 区域引用。
   * @returns {{sheet: string, range: string, formula: string}} 解析结果。
   */
  #resolveRangeRef(defaultSheet, ref) {
    const bang = ref.lastIndexOf('!')
    const sheetName = bang >= 0 ? ref.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'") : defaultSheet
    const range = bang >= 0 ? ref.slice(bang + 1) : ref
    const bounds = parseRange(range)
    this.sheet(sheetName)
    return {
      sheet: sheetName,
      range,
      formula: `${quoteSheetName(sheetName)}!$${indexToCol(bounds.start.col)}$${bounds.start.row}:$${indexToCol(bounds.end.col)}$${bounds.end.row}`
    }
  }

  /**
   * 插入或删除行列。
   *
   * 只重排行号与单元格引用；公式文本、合并区域、条件格式与数据验证中的
   * 区域引用不会自动改写，因此返回 warnings 明确提示（不静默改变语义）。
   *
   * @param {string} name - 工作表名。
   * @param {object} args - `{axis:'row'|'column', action:'insert'|'delete', start, count}`。
   * @returns {object} `{changes, warnings}` 变更与警告。
   */
  shiftRowsOrColumns(name, { axis, action, start, count = 1 }) {
    if (axis !== 'row' && axis !== 'column') throw new OfficeError('INVALID_REQUEST', "axis 必须是 'row' 或 'column'。")
    if (action !== 'insert' && action !== 'delete') throw new OfficeError('INVALID_REQUEST', "action 必须是 'insert' 或 'delete'。")
    if (!Number.isInteger(start) || start < 1) throw new OfficeError('INVALID_REQUEST', 'start 必须是正整数。')
    if (!Number.isInteger(count) || count < 1) throw new OfficeError('INVALID_REQUEST', 'count 必须是正整数。')

    const sheet = this.sheet(name)
    const doc = this.#doc(sheet.path)
    const sheetData = find(doc.root, 'sheetData')
    if (!sheetData) throw new OfficeError('CORRUPTED_DOCUMENT', '工作表缺少 sheetData。')

    const rows = (sheetData.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'row')
    if (action === 'delete') {
      const total = rows.reduce((m, r) => Math.max(m, Number(attr(r, 'r') ?? 0)), 0)
      if (start > total) {
        throw new OfficeError('INVALID_REQUEST', `起始${axis === 'row' ? '行' : '列'} ${start} 超出工作表范围（共 ${total} 行）。`)
      }
    }

    const warnings = []
    let shifted = 0
    let removed = 0

    for (const rowNode of rows) {
      const r = Number(attr(rowNode, 'r') ?? 0)
      if (r <= 0) continue

      if (axis === 'row' && action === 'delete' && r >= start && r < start + count) {
        doc.remove(rowNode)
        removed += 1
        continue
      }

      const newRow = axis === 'row' && action === 'insert' && r >= start ? r + count : axis === 'row' && action === 'delete' && r >= start + count ? r - count : r

      for (const c of rowNode.children ?? []) {
        if (c.type !== 'element' || localName(c) !== 'c') continue
        const ref = attr(c, 'r')
        if (!ref) continue
        const parsed = parseRef(ref)
        let newCol = parsed.col
        if (axis === 'column') {
          if (action === 'insert' && newCol >= start) newCol += count
          else if (action === 'delete' && newCol >= start && newCol < start + count) {
            doc.remove(c)
            continue
          } else if (action === 'delete' && newCol >= start + count) newCol -= count
        }
        if (newCol !== parsed.col || newRow !== r) {
          doc.setAttr(c, 'r', formatRef(newCol, newRow))
        }
      }
      if (newRow !== r) doc.setAttr(rowNode, 'r', String(newRow))
      shifted += 1
    }

    if (findAll(doc.root, 'f').length > 0) {
      warnings.push({
        code: 'FORMULA_REFERENCES_NOT_SHIFTED',
        message: `工作表「${name}」含公式，其内部单元格引用未随${axis === 'row' ? '行' : '列'}移位而调整，请在 Excel/WPS 中重算或改用重算模式。`
      })
    }
    if (find(doc.root, 'mergeCells')) {
      warnings.push({
        code: 'MERGE_RANGES_NOT_SHIFTED',
        message: `工作表「${name}」含合并单元格，其区域引用未随${axis === 'row' ? '行' : '列'}移位而调整。`
      })
    }
    for (const tag of ['conditionalFormatting', 'dataValidation', 'autoFilter', 'protectedRange']) {
      if (find(doc.root, tag)) {
        warnings.push({ code: 'RANGE_FEATURE_NOT_SHIFTED', message: `工作表「${name}」含 <${tag}>，其区域引用未自动调整。` })
      }
    }

    this.#flush(sheet.path)
    return {
      changes: [{ type: axis === 'row' ? 'row_shift' : 'column_shift', sheet: name, action, start, count, shifted, removed }],
      warnings
    }
  }

  /**
   * 应用单元格样式（语义化接口）。
   *
   * 调用方只需描述「加粗、红色字、黄底、千分位、居中」这类意图，适配器负责：
   * 在 styles.xml 中查重并追加 `<font>` / `<fill>` / `<numFmt>`，
   * 再组合出一份新的 `<xf>` 追加到 `<cellXfs>`，最后把单元格的 `s` 指过去。
   * 既有样式定义一律不动，因此不会影响其它已使用这些样式的单元格。
   *
   * @param {string} name - 工作表名。
   * @param {object[]} edits - 每项 `{ref, bold, italic, underline, fontColor, fontSize, fontName, fillColor, numberFormat, horizontal}`。
   * @returns {object[]} 变更记录。
   */
  applyCellStyle(name, edits) {
    if (!this.#pkg.has('xl/styles.xml')) {
      throw new OfficeError('UNSUPPORTED_FEATURE', '该工作簿没有 styles.xml，无法设置样式。')
    }
    const stylesDoc = this.#doc('xl/styles.xml')
    const stylesRoot = stylesDoc.root.children.find((c) => c.type === 'element')
    const cellXfs = find(stylesRoot, 'cellXfs')
    if (!cellXfs) throw new OfficeError('CORRUPTED_DOCUMENT', 'styles.xml 缺少 cellXfs。')
    const existingXfCount = (cellXfs.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'xf').length

    const sheet = this.sheet(name)
    const doc = this.#doc(sheet.path)
    const sheetData = find(doc.root, 'sheetData')
    if (!sheetData) throw new OfficeError('CORRUPTED_DOCUMENT', '工作表缺少 sheetData。')

    const cellIndex = new Map()
    for (const row of findAll(sheetData, 'row')) {
      const r = Number(attr(row, 'r') ?? 0)
      for (const c of row.children ?? []) {
        if (c.type !== 'element' || localName(c) !== 'c') continue
        const ref = attr(c, 'r')
        if (ref) cellIndex.set(`${parseRef(ref).col}:${r}`, c)
      }
    }

    const cache = new Map()
    const changes = []
    // 本次调用中新增到 fonts/fills/numFmts 的条目不在已解析的树里，
    // 必须单独记账，否则索引会与既有条目撞车（曾导致两个单元格共用同一字体）。
    const styleState = { appended: new Map() }
    for (const edit of edits) {
      const { col, row } = parseRef(edit.ref)
      const target = cellIndex.get(`${col}:${row}`)
      if (!target) {
        throw new OfficeError('FILE_NOT_FOUND', `单元格 ${edit.ref} 不存在，请先用 office_write_cells 写入内容，再设置样式。`, { ref: edit.ref })
      }
      // 缓存键必须排除 ref：多个单元格共用同一样式时应复用同一份定义，否则 styles.xml 会膨胀。
      const key = JSON.stringify([edit.bold, edit.italic, edit.underline, edit.fontColor, edit.fontSize, edit.fontName, edit.fillColor, edit.numberFormat, edit.horizontal])
      let styleIndex = cache.get(key)
      if (styleIndex === undefined) {
        const fontXml = buildFontXml(edit)
        const fontId = fontXml ? internStyleChild(stylesDoc, stylesRoot, 'fonts', fontXml, styleState) : 0
        const fillId = edit.fillColor ? internStyleChild(stylesDoc, stylesRoot, 'fills', buildFillXml(edit.fillColor), styleState) : 0
        const numFmtId = resolveNumFmtId(stylesDoc, stylesRoot, edit.numberFormat, styleState)
        const apply =
          (fontXml ? ' applyFont="1"' : '') +
          (edit.fillColor ? ' applyFill="1"' : '') +
          (edit.numberFormat ? ' applyNumberFormat="1"' : '') +
          (edit.horizontal ? ' applyAlignment="1"' : '')
        const body = edit.horizontal ? `><alignment horizontal="${escapeXmlAttr(edit.horizontal)}"/></xf>` : '/>'
        styleIndex = existingXfCount + cache.size
        stylesDoc.appendChild(cellXfs, `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="0" xfId="0"${apply}${body}`)
        cache.set(key, styleIndex)
      }
      doc.setAttr(target, 's', String(styleIndex))
      changes.push({ type: 'cell_style', sheet: name, range: edit.ref, style_index: styleIndex, applied: describeStyle(edit) })
    }
    stylesDoc.setAttr(cellXfs, 'count', String(existingXfCount + cache.size))
    this.#flush('xl/styles.xml')
    this.#flush(sheet.path)
    return changes
  }

  /** @returns {Buffer} 序列化后的工作簿字节。 */
  save() {
    return this.#pkg.toBuffer()
  }
}

/**
 * 取节点本地名。
 * @param {object} node - 元素节点。
 * @returns {string} 本地名。
 */
function localName(node) {
  return node.name.includes(':') ? node.name.slice(node.name.indexOf(':') + 1) : node.name
}

/** 与语言环境无关的内置数字格式（ECMA-376 §18.8.30 中语义明确的那部分）。 */
const BUILTIN_NUM_FMTS = new Map([
  ['General', 0],
  ['0', 1],
  ['0.00', 2],
  ['#,##0', 3],
  ['#,##0.00', 4],
  ['0%', 9],
  ['0.00%', 10],
  ['0.00E+00', 11],
  ['@', 49]
])

/** 自定义数字格式 ID 的起始值（0–163 保留给内置格式）。 */
const CUSTOM_NUM_FMT_START = 164

/**
 * 把颜色规范化为 ARGB 十六进制。
 * @param {string} color - `#RRGGBB` / `RRGGBB` / `AARRGGBB`。
 * @returns {string} 8 位 ARGB。
 */
function normalizeColor(color) {
  const hex = String(color).trim().replace(/^#/, '').toUpperCase()
  if (/^[0-9A-F]{6}$/.test(hex)) return `FF${hex}`
  if (/^[0-9A-F]{8}$/.test(hex)) return hex
  throw new OfficeError('INVALID_REQUEST', `非法颜色值：${color}（应为 #RRGGBB 或 AARRGGBB）。`)
}

/**
 * 生成 `<font>`；无字体相关选项时返回 null。
 *
 * 子元素顺序遵循 ECMA-376 中 CT_Font 的 sequence（b → i → u → sz → color → name），
 * 顺序错误时部分读取器会静默忽略后面的元素。
 *
 * @param {object} edit - 样式选项。
 * @returns {string|null} font XML。
 */
function buildFontXml(edit) {
  const parts = []
  if (edit.bold === true) parts.push('<b/>')
  if (edit.italic === true) parts.push('<i/>')
  if (edit.underline === true) parts.push('<u/>')
  parts.push(`<sz val="${edit.fontSize ? Number(edit.fontSize) : 11}"/>`)
  if (edit.fontColor) parts.push(`<color rgb="${normalizeColor(edit.fontColor)}"/>`)
  parts.push(`<name val="${escapeXmlAttr(edit.fontName ?? 'Calibri')}"/>`)
  const hasStyle = edit.bold === true || edit.italic === true || edit.underline === true || edit.fontColor || edit.fontSize || edit.fontName
  return hasStyle ? `<font>${parts.join('')}</font>` : null
}

/**
 * 生成纯色填充的 `<fill>`。
 * @param {string} color - 填充色。
 * @returns {string} fill XML。
 */
function buildFillXml(color) {
  const argb = normalizeColor(color)
  return `<fill><patternFill patternType="solid"><fgColor rgb="${argb}"/><bgColor indexed="64"/></patternFill></fill>`
}

/**
 * 生成样式表集合（fonts/fills）中的一项，已存在则复用其索引。
 *
 * 索引既要数已解析的子元素，也要数本次调用通过补丁追加的条目 ——
 * 补丁节点不在树里，只数树会导致新条目与既有条目索引撞车。
 *
 * @param {XmlDoc} doc - styles.xml 文档。
 * @param {object} root - styleSheet 根元素。
 * @param {string} containerName - 集合名（`fonts` / `fills`）。
 * @param {string} childXml - 待插入的子元素 XML。
 * @param {{appended: Map<string, string[]>}} state - 本次调用的追加记账。
 * @returns {number} 子元素索引。
 */
function internStyleChild(doc, root, containerName, childXml, state) {
  const itemName = containerName.replace(/s$/, '')
  const needle = normalizeXmlFragment(childXml)
  const appended = state.appended.get(containerName) ?? []

  const container = find(root, containerName)
  if (!container) {
    const anchor = find(root, 'cellStyleXfs') ?? find(root, 'cellXfs') ?? root.children.find((c) => c.type === 'element')
    const xml = `<${containerName} count="1">${childXml}</${containerName}>`
    if (anchor) doc.insertBefore(anchor, xml)
    else doc.appendChild(root, xml)
    state.appended.set(containerName, [needle])
    return 0
  }

  // 先在本轮新增里查重，再在既有内容里查重。本轮新增的条目排在既有条目之后。
  const base = countExisting(doc, container, itemName)
  const inAppended = appended.indexOf(needle)
  if (inAppended !== -1) return base + inAppended
  const children = (container.children ?? []).filter((n) => n.type === 'element' && localName(n) === itemName)
  for (let i = 0; i < children.length; i += 1) {
    if (normalizeXmlFragment(doc.raw(children[i])) === needle) return i
  }
  const index = base + appended.length
  doc.appendChild(container, childXml)
  appended.push(needle)
  state.appended.set(containerName, appended)
  doc.setAttr(container, 'count', String(base + appended.length))
  return index
}

/**
 * 数一个容器里已解析的子元素个数。
 * @param {XmlDoc} doc - 文档。
 * @param {object} container - 容器元素。
 * @param {string} itemName - 子元素本地名。
 * @returns {number} 数量。
 */
function countExisting(doc, container, itemName) {
  return (container.children ?? []).filter((n) => n.type === 'element' && localName(n) === itemName).length
}

/**
 * 解析或登记数字格式，返回 numFmtId。
 * @param {XmlDoc} doc - styles.xml 文档。
 * @param {object} root - styleSheet 根元素。
 * @param {string|undefined} formatCode - 格式代码，如 `#,##0.00`。
 * @param {{appended: Map<string, string[]>}} state - 本次调用的追加记账。
 * @returns {number} 数字格式 ID；未指定时为 0（General）。
 */
function resolveNumFmtId(doc, root, formatCode, state) {
  if (!formatCode) return 0
  const builtin = BUILTIN_NUM_FMTS.get(formatCode)
  if (builtin !== undefined) return builtin

  const containerKey = '__numFmts_codes'
  const appendedCodes = state.appended.get(containerKey) ?? []

  const existing = find(root, 'numFmts')
  if (existing) {
    for (const nf of findAll(existing, 'numFmt')) {
      if (attr(nf, 'formatCode') === formatCode) return Number(attr(nf, 'numFmtId') ?? 0)
    }
    // 本轮已登记过的同一格式直接复用（补丁节点不在树里）。
    const reusedIndex = appendedCodes.indexOf(formatCode)
    if (reusedIndex !== -1) return CUSTOM_NUM_FMT_START + reusedIndex
    const used = new Set(findAll(existing, 'numFmt').map((n) => Number(attr(n, 'numFmtId') ?? 0)))
    for (let i = 0; i < appendedCodes.length; i += 1) used.add(CUSTOM_NUM_FMT_START + i)
    let id = CUSTOM_NUM_FMT_START
    while (used.has(id)) id += 1
    doc.appendChild(existing, `<numFmt numFmtId="${id}" formatCode="${escapeXmlAttr(formatCode)}"/>`)
    appendedCodes.push(formatCode)
    state.appended.set(containerKey, appendedCodes)
    doc.setAttr(existing, 'count', String(used.size + 1))
    return id
  }

  // numFmts 必须是 styleSheet 的第一个子元素（ECMA-376 顺序约束）。
  const first = root.children.find((c) => c.type === 'element')
  const xml = `<numFmts count="1"><numFmt numFmtId="${CUSTOM_NUM_FMT_START}" formatCode="${escapeXmlAttr(formatCode)}"/></numFmts>`
  if (first) doc.insertBefore(first, xml)
  else doc.appendChild(root, xml)
  state.appended.set(containerKey, [formatCode])
  return CUSTOM_NUM_FMT_START
}

/**
 * 归一化 XML 片段用于查重比较：忽略标签之间的空白。
 * @param {string} xml - XML 片段。
 * @returns {string} 归一化结果。
 */
function normalizeXmlFragment(xml) {
  return xml.replace(/>\s+</g, '><').trim()
}

/**
 * 生成样式的人类可读描述。
 * @param {object} edit - 样式选项。
 * @returns {string} 描述文本。
 */
function describeStyle(edit) {
  const parts = []
  if (edit.bold === true) parts.push('加粗')
  if (edit.italic === true) parts.push('斜体')
  if (edit.underline === true) parts.push('下划线')
  if (edit.fontColor) parts.push(`字色 ${edit.fontColor}`)
  if (edit.fontSize) parts.push(`${edit.fontSize} 号字`)
  if (edit.fontName) parts.push(`字体 ${edit.fontName}`)
  if (edit.fillColor) parts.push(`底色 ${edit.fillColor}`)
  if (edit.numberFormat) parts.push(`数字格式 ${edit.numberFormat}`)
  if (edit.horizontal) parts.push(`水平对齐 ${edit.horizontal}`)
  return parts.join('、') || '无变化'
}

/**
 * 收集元素的全部文本。
 * @param {XmlDoc} doc - 文档。
 * @param {object} node - 元素节点。
 * @returns {string} 文本。
 */
function collectText(doc, node) {
  let text = ''
  for (const child of node.children ?? []) {
    if (child.type === 'text') text += decodeEntities(child.raw)
    else if (child.type === 'cdata') text += child.raw
    else if (child.type === 'element') text += collectText(doc, child)
  }
  return text
}

/**
 * 执行一次字符串替换。
 * @param {string} original - 原文本。
 * @param {string} needle - 查找内容。
 * @param {string} replacement - 替换内容。
 * @param {boolean} matchCase - 是否区分大小写。
 * @param {boolean} wholeCell - 是否整格匹配。
 * @returns {string} 替换结果。
 */
function applyReplace(original, needle, replacement, matchCase, wholeCell) {
  if (wholeCell) {
    const equal = matchCase ? original === needle : original.toLowerCase() === needle.toLowerCase()
    return equal ? replacement : original
  }
  if (matchCase) return original.split(needle).join(replacement)
  const escaped = needle.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return original.replace(new RegExp(escaped, 'gi'), () => replacement)
}

/**
 * 在共享字符串表中登记字符串，命中则复用索引。
 * @param {object} sst - 共享字符串表。
 * @param {string} value - 字符串。
 * @param {{touched: boolean}} dirty - 脏标记。
 * @returns {number} 共享字符串索引。
 */
function internSharedString(sst, value) {
  const hit = sst.items.indexOf(value)
  if (hit !== -1) return hit
  if (!sst.doc || !sst.part) return -1 // 无共享字符串表时退化为内联字符串
  const root = sst.doc.root.children.find((c) => c.type === 'element')
  sst.doc.appendChild(root, `<si><t xml:space="preserve">${escapeXmlText(value)}</t></si>`)
  sst.items.push(value)
  sst.doc.setAttr(root, 'uniqueCount', String(sst.items.length))
  sst.dirty = true
  return sst.items.length - 1
}

/**
 * 增加共享字符串引用计数。
 *
 * ECMA-376 中 `count` 是引用总数、`uniqueCount` 是唯一串数；每写入一个引用
 * `t="s"` 的单元格就要 +1，只新增唯一串时 `uniqueCount` 才变化。
 *
 * @param {object} sst - 共享字符串表。
 * @param {number} n - 本次新增的引用数。
 * @returns {void}
 */
function bumpSharedStringRefs(sst, n) {
  if (!sst.doc || !sst.part || n <= 0) return
  sst.refs += n
  const root = sst.doc.root.children.find((c) => c.type === 'element')
  sst.doc.setAttr(root, 'count', String(sst.refs))
  sst.dirty = true
}

/**
 * 构造一个单元格的 XML。
 * @param {string} ref - A1 引用。
 * @param {object} edit - 编辑定义。
 * @param {number} styleIndex - 样式索引。
 * @param {object} sst - 共享字符串表。
 * @returns {string} 单元格 XML。
 */
function buildCellXml(ref, edit, styleIndex, sst, stats = null) {
  const styleAttr = styleIndex ? ` s="${styleIndex}"` : ''
  const formula = edit.formula ? `<f>${escapeXmlText(String(edit.formula).replace(/^=/u, ''))}</f>` : ''
  const value = edit.value

  if (formula) {
    const cached = value === undefined || value === null ? '' : `<v>${escapeXmlText(String(value))}</v>`
    return `<c r="${ref}"${styleAttr}>${formula}${cached}</c>`
  }
  if (value === null || value === undefined || value === '') {
    return `<c r="${ref}"${styleAttr}/>`
  }
  if (typeof value === 'boolean') {
    return `<c r="${ref}"${styleAttr} t="b"><v>${value ? 1 : 0}</v></c>`
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return `<c r="${ref}"${styleAttr}><v>${value}</v></c>`
  }
  if (edit.type === 'date' || edit.type === 'datetime') {
    const serial = isoToSerial(String(value), edit.date1904 === true)
    return `<c r="${ref}"${styleAttr}><v>${serial}</v></c>`
  }
  // 字符串：优先共享字符串表（保持与 Excel 一致且体积更小）；
  // ponytail: 源文件没有 sst 部件时退化为内联字符串，避免为此新增部件与关系。
  const index = internSharedString(sst, String(value))
  if (index >= 0) {
    if (stats) stats.sstRefs += 1
    return `<c r="${ref}"${styleAttr} t="s"><v>${index}</v></c>`
  }
  return `<c r="${ref}"${styleAttr} t="inlineStr"><is><t xml:space="preserve">${escapeXmlText(String(value))}</t></is></c>`
}

/**
 * 生成图表部件 XML。
 *
 * 元素顺序是 schema 约束的（Excel 对顺序极敏感，错一处就要求修复文件）：
 * `c:chart` 内是 title → autoTitleDeleted → plotArea → legend → plotVisOnly；
 * `c:plotArea` 内是 layout → 图表组 → 坐标轴；
 * 柱状/条形图的图表组内是 barDir → grouping → varyColors → ser* → axId, axId。
 *
 * @param {object} args - 参数。
 * @param {string} args.kind - `bar` | `line` | `pie`。
 * @param {string} args.type - 用户给的语义类型（column/bar/line/pie）。
 * @param {string} args.seriesXml - 系列 XML。
 * @param {string|null} args.title - 标题。
 * @param {number} args.seriesCount - 系列数。
 * @returns {string} 图表 XML。
 */
function buildChartXml({ kind, type, seriesXml, title, seriesCount }) {
  const titleXml = title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXmlText(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : ''
  const autoTitleDeleted = title ? '' : '<c:autoTitleDeleted val="1"/>'
  const catAxId = '111111111'
  const valAxId = '222222222'
  let group = ''
  let axes = ''
  if (kind === 'bar') {
    group =
      `<c:barChart><c:barDir val="${type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
      seriesXml +
      `<c:gapWidth val="150"/><c:axId val="${catAxId}"/><c:axId val="${valAxId}"/></c:barChart>`
    axes = catAxXml(catAxId, valAxId) + valAxXml(valAxId, catAxId)
  } else if (kind === 'line') {
    group =
      `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` +
      seriesXml +
      `<c:marker val="1"/><c:axId val="${catAxId}"/><c:axId val="${valAxId}"/></c:lineChart>`
    axes = catAxXml(catAxId, valAxId) + valAxXml(valAxId, catAxId)
  } else {
    group = `<c:pieChart><c:varyColors val="1"/>${seriesXml}<c:firstSliceAng val="0"/></c:pieChart>`
  }
  const legend = seriesCount > 1 || kind === 'pie' ? '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>' : ''
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<c:chartSpace xmlns:c="${NS_CHART}" xmlns:a="${NS_DRAWING_MAIN}" xmlns:r="${NS_REL}">` +
    `<c:chart>${titleXml}${autoTitleDeleted}<c:plotArea><c:layout/>${group}${axes}</c:plotArea>${legend}` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></c:spPr>` +
    `</c:chartSpace>`
  )
}

/**
 * 分类轴 XML。
 * @param {string} id - 轴 id。
 * @param {string} crossId - 交叉轴 id。
 * @returns {string} `<c:catAx>`。
 */
function catAxXml(id, crossId) {
  return (
    `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>` +
    `<c:axPos val="b"/><c:crossAx val="${crossId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>`
  )
}

/**
 * 数值轴 XML。
 * @param {string} id - 轴 id。
 * @param {string} crossId - 交叉轴 id。
 * @returns {string} `<c:valAx>`。
 */
function valAxXml(id, crossId) {
  return (
    `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>` +
    `<c:axPos val="l"/><c:crossAx val="${crossId}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`
  )
}

/**
 * 生成绘图锚点 XML（twoCellAnchor + graphicFrame 引用图表部件）。
 * @param {object} args - 参数。
 * @param {string} args.anchor - 左上角单元格，如 F2。
 * @param {number} args.widthPx - 宽（像素）。
 * @param {number} args.heightPx - 高（像素）。
 * @param {string} args.rid - 指向图表部件的关系 ID。
 * @param {number} args.index - 锚点序号（用于形状 id 与名称）。
 * @param {object} args.metrics - 工作表的列宽/行高度量。
 * @returns {string} 锚点 XML。
 */
function buildChartAnchorXml({ anchor, widthPx, heightPx, rid, index, metrics }) {
  const from = parseRef(anchor)
  const columnPx = metrics ? metrics.columnPx : () => 64
  const rowPx = metrics ? metrics.rowPx : () => 20
  const startCol = from.col - 1
  const startRow = from.row - 1
  // 按实际列宽/行高累加，找出目标尺寸落在哪一列/哪一行
  let acc = 0
  let toCol = startCol
  while (acc < widthPx && toCol < startCol + 200) {
    acc += columnPx(toCol)
    toCol += 1
  }
  let accRow = 0
  let toRow = startRow
  while (accRow < heightPx && toRow < startRow + 500) {
    accRow += rowPx(toRow)
    toRow += 1
  }
  const cx = Math.round(widthPx * EMU_PER_PX_LOCAL)
  const cy = Math.round(heightPx * EMU_PER_PX_LOCAL)
  return (
    `<xdr:twoCellAnchor>` +
    `<xdr:from><xdr:col>${startCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${startRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:to><xdr:col>${toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>` +
    `<xdr:cNvPr id="${index + 1}" name="图表 ${index}"/><xdr:cNvGraphicFramePr/>` +
    `</xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${NS_CHART}">` +
    `<c:chart xmlns:c="${NS_CHART}" xmlns:r="${NS_REL}" r:id="${rid}"/>` +
    `</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`
  )
}

/**
 * 分类的字符串缓存。
 * @param {string[]} values - 文本值。
 * @returns {string} `<c:strCache>`。
 */
function strCacheXml(values) {
  const points = values.map((value, index) => `<c:pt idx="${index}"><c:v>${escapeXmlText(value)}</c:v></c:pt>`).join('')
  return `<c:strCache><c:ptCount val="${values.length}"/>${points}</c:strCache>`
}

/**
 * 数值缓存。
 * @param {number[]} values - 数值。
 * @returns {string} `<c:numCache>`。
 */
function numCacheXml(values) {
  const points = values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join('')
  return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numCache>`
}

/**
 * 给工作表名加引号（图表里的引用公式用）。
 * @param {string} name - 工作表名。
 * @returns {string} 引用用的表名。
 */
function quoteSheetName(name) {
  return /^[A-Za-z_\u4e00-\u9fff][A-Za-z0-9_.\u4e00-\u9fff]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`
}

/**
 * 读取关系文件。
 * @param {ZipPackage} pkg - OPC 包。
 * @param {string} part - 关系部件路径。
 * @returns {Map<string, string>} rId → 目标路径。
 */
function readRels(pkg, part) {
  const map = new Map()
  if (!pkg.has(part)) return map
  const doc = XmlDoc.parse(pkg.readText(part))
  for (const rel of findAll(doc.root, 'Relationship')) {
    map.set(attr(rel, 'Id'), attr(rel, 'Target'))
  }
  return map
}

/**
 * 把关系目标解析为包内绝对路径。
 * @param {string} baseDir - 宿主部件所在目录。
 * @param {string} target - 关系目标。
 * @returns {string} 归一化后的部件路径。
 */
function normalizePart(baseDir, target) {
  if (target.startsWith('/')) return target.slice(1)
  const segments = `${baseDir}/${target}`.split('/')
  const stack = []
  for (const seg of segments) {
    if (seg === '.' || seg === '') continue
    if (seg === '..') stack.pop()
    else stack.push(seg)
  }
  return stack.join('/')
}

/**
 * 计算下一个可用的关系 ID。
 * @param {ZipPackage} pkg - OPC 包。
 * @param {string} part - 关系部件路径。
 * @returns {string} 形如 `rIdN` 的 ID。
 */
function nextRelId(pkg, part) {
  if (!pkg.has(part)) return 'rId1'
  const doc = XmlDoc.parse(pkg.readText(part))
  let max = 0
  for (const rel of findAll(doc.root, 'Relationship')) {
    const id = attr(rel, 'Id') ?? ''
    const match = /^rId(\d+)$/.exec(id)
    if (match) max = Math.max(max, Number(match[1]))
  }
  return `rId${max + 1}`
}

/* 内容类型与关系操作已下沉到 ooxml.js（ensureContentTypeOverride /
 * removeContentTypeOverride / addRelationshipTo / removeRelationshipFrom），
 * 与 DOCX / PPTX 适配器共用同一份实现。 */

/** 新建工作表的最小合法 XML。 */
const EMPTY_SHEET_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><sheetData/></worksheet>`

/**
 * 构造 `c:chartSpace` 图表本体 XML。
 *
 * **导出给 PPTX 适配器复用**：XLSX 与 PPTX 的图表用的是同一套 DrawingML 图表 schema，
 * 只是宿主不同（一个放在 `xl/charts/`，一个放在 `ppt/charts/`）。复制一份迟早会漂移，
 * 所以这里显式导出，两个适配器共用同一份构造器。
 *
 * @param {object} args - 参数。
 * @param {'bar'|'line'|'pie'} args.kind - 图表族。
 * @param {string} args.type - 具体类型（`column` / `bar` / `line` / `pie`）。
 * @param {string} args.seriesXml - 系列 XML。
 * @param {string} [args.title] - 标题。
 * @param {number} args.seriesCount - 系列数（决定是否画图例）。
 * @returns {string} 图表部件 XML。
 */
export function buildChartSpaceXml({ kind, type, seriesXml, title = null, seriesCount }) {
  return buildChartXml({ kind, type, seriesXml, title, seriesCount })
}

/**
 * 构造一个数据系列的 `c:ser` XML。
 *
 * 默认写**字面量**（`c:strLit` / `c:numLit`）；给了 `nameRef` / `categoryRef` / `valuesRef`
 * 则改写成**引用 + 缓存**（`c:strRef` / `c:numRef`，缓存照旧带上）—— PPTX 的嵌入工作簿
 * 就是这么用的：引用让 PowerPoint 的「编辑数据」能对上源表，缓存让阅读器不必重算即可作图。
 *
 * @param {object} args - 参数。
 * @param {number} args.index - 系列下标。
 * @param {string} [args.name] - 系列名。
 * @param {string} [args.nameRef] - 系列名引用（如 `Sheet1!$B$1`）。
 * @param {string[]} args.categories - 分类标签。
 * @param {string} [args.categoryRef] - 分类引用（如 `Sheet1!$A$2:$A$4`）。
 * @param {number[]} args.values - 数值。
 * @param {string} [args.valuesRef] - 数值引用（如 `Sheet1!$B$2:$B$4`）。
 * @returns {string} `c:ser` XML。
 */
export function buildLiteralSeriesXml({ index, name = null, nameRef = null, categories, categoryRef = null, values, valuesRef = null }) {
  const tx = nameRef
    ? `<c:tx><c:strRef><c:f>${escapeXmlText(nameRef)}</c:f>${strCacheXml([name ?? ''])}</c:strRef></c:tx>`
    : name
      ? `<c:tx><c:v>${escapeXmlText(name)}</c:v></c:tx>`
      : ''
  const cat = categoryRef
    ? `<c:cat><c:strRef><c:f>${escapeXmlText(categoryRef)}</c:f>${strCacheXml(categories)}</c:strRef></c:cat>`
    : `<c:cat><c:strLit>${strCacheXml(categories)}</c:strLit></c:cat>`
  const val = valuesRef
    ? `<c:val><c:numRef><c:f>${escapeXmlText(valuesRef)}</c:f>${numCacheXml(values)}</c:numRef></c:val>`
    : `<c:val><c:numLit>${numCacheXml(values)}</c:numLit></c:val>`
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${tx}${cat}${val}</c:ser>`
}

export { EMPTY_SHEET_XML, NS_MAIN, NS_REL, CT_WORKSHEET }
