/**
 * 从零生成 PDF：自写页面树、内容流、字体资源与 xref。
 *
 * 设计边界（诚实说明）：
 *   - 只用 **PDF 标准 14 字体**（Helvetica / Times / Courier 四体），因此**不需要嵌入字体**，
 *     任何阅读器都能打开；代价是只支持 **WinAnsi 字符集**（拉丁字母、数字、常见标点）。
 *     中文/日文/韩文需要嵌入字体子集，属未实现能力（`pdf.create.cjk`），会**明确拒绝**而不是画乱码。
 *   - 文本按行排布、自动折行与分页；不做富文本、表格、图片、矢量图形。
 *   - 同时写一份 `/ToUnicode` CMap，因此**文本可被提取**（自己与第三方解析器都能读回），
 *     而不是只能"看"不能"读"。
 *
 * @module dsh-exp-office/pdf-writer
 */

import { deflateSync } from 'node:zlib'
import { OfficeError } from './errors.js'
import { serializePdfFile } from './pdf.js'

/** 标准纸张尺寸（点，1 pt = 1/72 inch）。 */
export const PDF_PAGE_SIZES = Object.freeze({
  A3: [841.89, 1190.55],
  A4: [595.28, 841.89],
  A5: [419.53, 595.28],
  Letter: [612, 792],
  Legal: [612, 1008]
})

/** 标准 14 字体里可用的四个字族（各自带粗体与斜体）。 */
export const PDF_BASE_FONTS = Object.freeze({
  helvetica: { regular: 'Helvetica', bold: 'Helvetica-Bold', italic: 'Helvetica-Oblique', boldItalic: 'Helvetica-BoldOblique' },
  times: { regular: 'Times-Roman', bold: 'Times-Bold', italic: 'Times-Italic', boldItalic: 'Times-BoldItalic' },
  courier: { regular: 'Courier', bold: 'Courier-Bold', italic: 'Courier-Oblique', boldItalic: 'Courier-BoldOblique' }
})

/**
 * WinAnsiEncoding 在 0x80–0x9F 段的码点（其余与 Latin-1 一致，可直接用码点本身）。
 * 依据 PDF 规范附录 D。未列出的码位（0x81/0x8D/0x8F/0x90/0x9D）在 WinAnsi 里无定义，视为不可编码。
 */
const WIN_ANSI_SPECIAL = Object.freeze({
  0x80: 0x20ac, // €
  0x82: 0x201a, // ‚
  0x83: 0x0192, // ƒ
  0x84: 0x201e, // „
  0x85: 0x2026, // …
  0x86: 0x2020, // †
  0x87: 0x2021, // ‡
  0x88: 0x02c6, // ˆ
  0x89: 0x2030, // ‰
  0x8a: 0x0160, // Š
  0x8b: 0x2039, // ‹
  0x8c: 0x0152, // Œ
  0x8e: 0x017d, // Ž
  0x91: 0x2018, // ‘
  0x92: 0x2019, // ’
  0x93: 0x201c, // “
  0x94: 0x201d, // ”
  0x95: 0x2022, // •
  0x96: 0x2013, // –
  0x97: 0x2014, // —
  0x98: 0x02dc, // ˜
  0x99: 0x2122, // ™
  0x9a: 0x0161, // š
  0x9b: 0x203a, // ›
  0x9c: 0x0153, // œ
  0x9e: 0x017e, // ž
  0x9f: 0x0178 // Ÿ
})

/** Unicode 码点 → WinAnsi 字节（仅包含可编码的部分）。 */
const UNICODE_TO_WIN_ANSI = new Map()
UNICODE_TO_WIN_ANSI.set(0x09, 0x09) // 制表符
for (let code = 0x20; code <= 0x7e; code += 1) UNICODE_TO_WIN_ANSI.set(code, code)
for (let code = 0xa0; code <= 0xff; code += 1) UNICODE_TO_WIN_ANSI.set(code, code)
for (const [byte, code] of Object.entries(WIN_ANSI_SPECIAL)) UNICODE_TO_WIN_ANSI.set(code, Number(byte))

/**
 * 找出文本里无法用 WinAnsi 表示的字符。
 * @param {string} text - 文本。
 * @returns {string[]} 无法编码的字符（去重，最多 10 个）。
 */
export function findUnencodableChars(text) {
  const bad = new Set()
  for (const ch of String(text)) {
    if (!UNICODE_TO_WIN_ANSI.has(ch.codePointAt(0))) bad.add(ch)
    if (bad.size >= 10) break
  }
  return [...bad]
}

/**
 * 把文本编码成 WinAnsi 字节（PDF 字符串里的原始字节）。
 * @param {string} text - 文本。
 * @returns {Buffer} 字节。
 */
function encodeWinAnsi(text) {
  const bytes = []
  for (const ch of String(text)) {
    const byte = UNICODE_TO_WIN_ANSI.get(ch.codePointAt(0))
    if (byte === undefined) {
      throw new OfficeError('UNSUPPORTED_FEATURE', `字符「${ch}」无法用 WinAnsi 表示（标准 14 字体不支持中日韩文字）。`, {
        char: ch,
        code_point: `U+${ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')}`
      })
    }
    bytes.push(byte)
  }
  return Buffer.from(bytes)
}

/**
 * 转义 PDF 字面字符串，并转成 `(...)` 形式。
 * @param {string} text - 文本。
 * @returns {string} 可直接写进内容流的字符串。
 */
function literalString(text) {
  const buf = encodeWinAnsi(text)
  let out = ''
  for (const byte of buf) {
    if (byte === 0x28) out += '\\('
    else if (byte === 0x29) out += '\\)'
    else if (byte === 0x5c) out += '\\\\'
    else if (byte < 0x20 || byte > 0x7e) out += `\\${byte.toString(8).padStart(3, '0')}`
    else out += String.fromCharCode(byte)
  }
  return `(${out})`
}

/**
 * 构造 WinAnsiEncoding 的 ToUnicode CMap（让文本可被提取）。
 * @returns {string} CMap 内容。
 */
function buildToUnicodeCMap() {
  const lines = [
    '/CIDInit /ProcSet findresource begin',
    '12 dict begin',
    'begincmap',
    '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
    '/CMapName /Adobe-Identity-UCS def',
    '/CMapType 2 def',
    '1 begincodespacerange',
    '<00> <FF>',
    'endcodespacerange'
  ]
  const ranges = [
    [0x20, 0x7e],
    [0xa0, 0xff]
  ]
  lines.push(`${ranges.length} beginbfrange`)
  for (const [from, to] of ranges) {
    lines.push(`<${from.toString(16).padStart(2, '0')}> <${to.toString(16).padStart(2, '0')}> <${from.toString(16).padStart(4, '0')}>`)
  }
  lines.push('endbfrange')
  const specials = Object.entries(WIN_ANSI_SPECIAL).map(([byte, code]) => [Number(byte), code])
  for (let i = 0; i < specials.length; i += 100) {
    const chunk = specials.slice(i, i + 100)
    lines.push(`${chunk.length} beginbfchar`)
    for (const [byte, code] of chunk) {
      lines.push(`<${byte.toString(16).padStart(2, '0')}> <${code.toString(16).padStart(4, '0')}>`)
    }
    lines.push('endbfchar')
  }
  lines.push('endcmap', 'CMapName currentdict /CMap defineresource pop', 'end', 'end')
  return lines.join('\n')
}

/**
 * 估算一行文本在给定字号下的宽度（单位：文本空间点）。
 *
 * 标准 14 字体自带宽度表，这里用各字族的**平均字宽近似**：
 * Helvetica/Times ≈ 0.5 em，Courier 是等宽 0.6 em。目的只是合理折行，
 * 不追求与阅读器逐像素一致（PDF 由阅读器自己排版，宽度不影响正确性）。
 *
 * @param {string} text - 文本。
 * @param {number} fontSize - 字号。
 * @param {string} family - 字族。
 * @returns {number} 宽度。
 */
function measure(text, fontSize, family) {
  const em = family === 'courier' ? 0.6 : 0.5
  return String(text).length * fontSize * em
}

/**
 * 按可用宽度折行（按空格断词；超长单词硬切）。
 * @param {string} text - 原始文本。
 * @param {number} maxWidth - 可用宽度。
 * @param {number} fontSize - 字号。
 * @param {string} family - 字族。
 * @returns {string[]} 折行后的行。
 */
function wrapLine(text, maxWidth, fontSize, family) {
  const source = String(text)
  if (source === '') return ['']
  if (measure(source, fontSize, family) <= maxWidth) return [source]
  const words = source.split(/(\s+)/).filter((w) => w !== '')
  const out = []
  let current = ''
  for (const word of words) {
    const candidate = current + word
    if (current !== '' && measure(candidate, fontSize, family) > maxWidth) {
      out.push(current.trimEnd())
      current = word.trimStart()
    } else {
      current = candidate
    }
    // 单个词就超宽：硬切
    while (measure(current, fontSize, family) > maxWidth && current.length > 1) {
      let cut = current.length
      while (cut > 1 && measure(current.slice(0, cut), fontSize, family) > maxWidth) cut -= 1
      out.push(current.slice(0, cut))
      current = current.slice(cut)
    }
  }
  if (current !== '') out.push(current.trimEnd())
  return out.length > 0 ? out : ['']
}

/**
 * 从零生成一份 PDF。
 *
 * @param {object} options - 内容选项。
 * @param {string[]} options.lines - 正文行（`''` 表示空行；`\f` 前缀强制分页）。
 * @param {string} [options.pageSize] - 纸张：A4/A3/A5/Letter/Legal，或 `[宽, 高]`（点）。
 * @param {'portrait'|'landscape'} [options.orientation] - 方向。
 * @param {number} [options.marginPt] - 页边距（点），默认 56.7（2 cm）。
 * @param {number} [options.fontSizePt] - 字号，默认 11。
 * @param {string} [options.fontFamily] - 字族：helvetica/times/courier。
 * @param {number} [options.lineHeightPt] - 行距，默认字号 × 1.4。
 * @param {object} [options.metadata] - `/Info` 字段（title/author/subject/keywords/creator/producer）。
 * @param {boolean} [options.compress] - 是否 Flate 压缩内容流，默认 true。
 * @returns {{buffer: Buffer, pages: number, lines: number}} 生成的字节与规模。
 */
export function buildPdf({
  lines = [],
  pageSize = 'A4',
  orientation = 'portrait',
  marginPt = 56.7,
  fontSizePt = 11,
  fontFamily = 'helvetica',
  lineHeightPt = null,
  metadata = {},
  compress = true
} = {}) {
  const family = PDF_BASE_FONTS[fontFamily]
  if (!family) {
    throw new OfficeError('INVALID_REQUEST', `不支持的字族：${fontFamily}（可用 ${Object.keys(PDF_BASE_FONTS).join(' / ')}）。`)
  }
  const size = Array.isArray(pageSize) ? pageSize : PDF_PAGE_SIZES[pageSize]
  if (!size) {
    throw new OfficeError('INVALID_REQUEST', `不支持的纸张：${String(pageSize)}（可用 ${Object.keys(PDF_PAGE_SIZES).join(' / ')}，或传 [宽, 高]）。`)
  }
  if (orientation !== 'portrait' && orientation !== 'landscape') {
    throw new OfficeError('INVALID_REQUEST', `方向只能是 portrait 或 landscape，收到 ${orientation}。`)
  }
  if (!Number.isFinite(fontSizePt) || fontSizePt <= 0) throw new OfficeError('INVALID_REQUEST', 'fontSizePt 必须是正数。')
  if (!Number.isFinite(marginPt) || marginPt < 0) throw new OfficeError('INVALID_REQUEST', 'marginPt 不能为负数。')
  const [rawWidth, rawHeight] = size
  const pageWidth = orientation === 'landscape' ? rawHeight : rawWidth
  const pageHeight = orientation === 'landscape' ? rawWidth : rawHeight
  const usableWidth = pageWidth - marginPt * 2
  if (usableWidth <= fontSizePt) {
    throw new OfficeError('INVALID_REQUEST', `页边距 ${marginPt} pt 太大，可用宽度只剩 ${Math.round(usableWidth)} pt。`)
  }
  const lineHeight = lineHeightPt ?? Math.round(fontSizePt * 1.4 * 100) / 100
  const usableHeight = pageHeight - marginPt * 2
  const linesPerPage = Math.max(1, Math.floor(usableHeight / lineHeight))

  if (!Array.isArray(lines)) throw new OfficeError('INVALID_REQUEST', 'lines 必须是字符串数组。')
  for (const line of lines) {
    if (typeof line !== 'string') throw new OfficeError('INVALID_REQUEST', 'lines 里只能放字符串。')
    if (/[\n\r]/.test(line)) {
      throw new OfficeError('INVALID_REQUEST', '单个元素里不能含换行：请把换行拆成多个元素（`\\f` 前缀表示强制分页）。')
    }
  }
  // 无法用 WinAnsi 表示的字符直接拒绝（中文等），而不是写出一堆乱码。
  // 注意：行首的 `\f` 是「强制分页」标记，校验前必须先去掉。
  for (const line of lines) {
    const bad = findUnencodableChars(line.startsWith('\f') ? line.slice(1) : line)
    if (bad.length > 0) {
      throw new OfficeError(
        'UNSUPPORTED_FEATURE',
        `正文包含标准 14 字体无法表示的字符：${bad.join(' ')}。从零生成 PDF 只支持 WinAnsi（拉丁字母/数字/常见标点）；中文需要嵌入字体子集，属未实现能力。`,
        { chars: bad }
      )
    }
  }

  // 1) 折行 + 分页
  const physical = []
  for (const line of lines) {
    const forceBreak = line.startsWith('\f')
    const body = forceBreak ? line.slice(1) : line
    if (forceBreak && physical.length > 0) physical.push({ breakBefore: true })
    for (const wrapped of wrapLine(body, usableWidth, fontSizePt, fontFamily)) physical.push({ text: wrapped })
  }
  const pages = []
  let currentPage = []
  let used = 0
  for (const item of physical) {
    if (item.breakBefore === true) {
      if (currentPage.length > 0) pages.push(currentPage)
      currentPage = []
      used = 0
      continue
    }
    if (used >= linesPerPage) {
      pages.push(currentPage)
      currentPage = []
      used = 0
    }
    currentPage.push(item.text)
    used += 1
  }
  pages.push(currentPage)
  if (pages.length === 0) pages.push([])

  // 2) 对象表：1=Catalog, 2=Pages, 3=Font, 4=ToUnicode, 5=内容流…, 之后每页一个页面对象
  const fontNum = 3
  const toUnicodeNum = 4
  const firstContentNum = 5
  const pageObjectNums = pages.map((_, index) => firstContentNum + pages.length + index)
  const objects = []
  objects.push({
    num: 1,
    value: { Type: 'Catalog', Pages: { ref: `2 0` } }
  })
  objects.push({
    num: 2,
    value: { Type: 'Pages', Kids: pageObjectNums.map((num) => ({ ref: `${num} 0` })), Count: pages.length }
  })
  objects.push({
    num: fontNum,
    value: {
      Type: 'Font',
      Subtype: 'Type1',
      BaseFont: family.regular,
      Encoding: 'WinAnsiEncoding',
      ToUnicode: { ref: `${toUnicodeNum} 0` }
    }
  })
  const cmap = buildToUnicodeCMap()
  objects.push({ num: toUnicodeNum, value: { dict: {}, stream: Buffer.from(cmap, 'latin1') } })

  pages.forEach((pageLines, index) => {
    // 文本算子必须写在 BT…ET 文本对象里（PDF 规范要求，也是提取器识别文本的前提）
    const parts = ['BT']
    pageLines.forEach((text, lineIndex) => {
      const y = pageHeight - marginPt - (lineIndex + 1) * lineHeight + lineHeight * 0.25
      parts.push(`/F1 ${round2(fontSizePt)} Tf`)
      parts.push(`1 0 0 1 ${round2(marginPt)} ${round2(y)} Tm`)
      parts.push(`${literalString(text)} Tj`)
    })
    parts.push('ET')
    const content = parts.join('\n')
    const streamBytes = Buffer.from(content, 'latin1')
    objects.push({
      num: firstContentNum + index,
      value: {
        dict: compress ? { Filter: 'FlateDecode' } : {},
        stream: compress ? deflateSync(streamBytes) : streamBytes
      }
    })
    objects.push({
      num: pageObjectNums[index],
      value: {
        Type: 'Page',
        Parent: { ref: '2 0' },
        MediaBox: [0, 0, round2(pageWidth), round2(pageHeight)],
        Resources: { Font: { F1: { ref: `${fontNum} 0` } } },
        Contents: { ref: `${firstContentNum + index} 0` }
      }
    })
  })

  const info = buildInfoDict(metadata)
  const result = serializePdfFile({
    objects: objects.sort((a, b) => a.num - b.num),
    rootNum: 1,
    inlineInfo: info,
    version: '1.7'
  })
  return { buffer: result.buffer, pages: pages.length, lines: physical.length }
}

/**
 * 构造 `/Info` 字典（非 ASCII 值按 UTF-16BE + BOM 写，PDF 规范要求）。
 * @param {object} metadata - 元数据。
 * @returns {object|null} 字典或 null。
 */
function buildInfoDict(metadata) {
  const dict = {}
  const text = (value) => {
    const str = String(value)
    // eslint-disable-next-line no-control-regex
    if (/^[\x20-\x7e]*$/.test(str)) return Buffer.from(str, 'latin1')
    return Buffer.concat([Buffer.from([0xfe, 0xff]), Buffer.from(str, 'utf16le').swap16()])
  }
  if (metadata.title) dict.Title = text(metadata.title)
  if (metadata.author) dict.Author = text(metadata.author)
  if (metadata.subject) dict.Subject = text(metadata.subject)
  if (metadata.keywords) dict.Keywords = text(metadata.keywords)
  dict.Creator = text(metadata.creator ?? 'dsh-exp-office')
  dict.Producer = text(metadata.producer ?? 'dsh-exp-office pdf-writer')
  const now = new Date()
  const pad = (n) => String(n).padStart(2, '0')
  dict.CreationDate = Buffer.from(
    `D:${now.getUTCFullYear()}${pad(now.getUTCMonth() + 1)}${pad(now.getUTCDate())}${pad(now.getUTCHours())}${pad(now.getUTCMinutes())}${pad(now.getUTCSeconds())}Z`,
    'latin1'
  )
  return dict
}

/**
 * 保留两位小数（PDF 里没必要写更长的数字）。
 * @param {number} value - 数值。
 * @returns {number} 处理后的数值。
 */
function round2(value) {
  return Math.round(value * 100) / 100
}
