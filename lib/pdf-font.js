/**
 * PDF 字体嵌入（共用）：为一段文本准备一份**子集化的 TrueType 字体**与配套的 PDF 数据。
 *
 * 供两处使用：
 *   - `pdf-writer.js`：从零生成 PDF 时的正文字体；
 *   - `pdf.js`：给已有 PDF 叠加中文水印 / 页码时，叠加流也需要一份可用字体。
 *
 * 分工：本模块只产出**数据**（子集字节、CID→GID 映射、宽度、ToUnicode、度量），
 * 具体对象怎么编号、怎么接进页面资源由调用方决定（生成器与增量更新两条路的编号方式不同）。
 *
 * @module dsh-exp-office/pdf-font
 */

import { existsSync, readFileSync } from 'node:fs'
import { OfficeError } from './errors.js'
import { readCmap, readFont, subsetFont } from './font-subset.js'

/** WinAnsi（标准 14 字体）能表示的码点：超出这个范围就需要嵌入字体。 */
const WIN_ANSI_CODES = new Set()
WIN_ANSI_CODES.add(0x09)
for (let code = 0x20; code <= 0x7e; code += 1) WIN_ANSI_CODES.add(code)
for (let code = 0xa0; code <= 0xff; code += 1) WIN_ANSI_CODES.add(code)
for (const code of [
  0x20ac, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x017d, 0x2018, 0x2019, 0x201c,
  0x201d, 0x2022, 0x2013, 0x2014, 0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x017e, 0x0178
]) {
  WIN_ANSI_CODES.add(code)
}

/**
 * 可用于嵌入的中文字体候选（Windows 上常见；按「纯 TTF 优先」排序）。
 * 用到中文时按顺序找第一个存在的；都找不到就明确报错，而不是画乱码。
 */
export const CJK_FONT_CANDIDATES = Object.freeze([
  { path: 'C:\\Windows\\Fonts\\simhei.ttf', index: 0, name: 'SimHei' },
  { path: 'C:\\Windows\\Fonts\\Deng.ttf', index: 0, name: 'DengXian' },
  { path: 'C:\\Windows\\Fonts\\simfang.ttf', index: 0, name: 'FangSong' },
  { path: 'C:\\Windows\\Fonts\\simkai.ttf', index: 0, name: 'KaiTi' },
  { path: 'C:\\Windows\\Fonts\\msyh.ttc', index: 0, name: 'MicrosoftYaHei' },
  { path: 'C:\\Windows\\Fonts\\simsun.ttc', index: 0, name: 'SimSun' },
  { path: 'C:\\Windows\\Fonts\\NotoSansSC-VF.ttf', index: 0, name: 'NotoSansSC' }
])

/**
 * 找一个可用的中文字体。
 * @param {string} [explicitPath] - 调用方指定的字体路径。
 * @param {number} [index] - `.ttc` 里的下标。
 * @returns {{path: string, index: number, name: string}} 字体位置。
 */
export function resolveCjkFont(explicitPath, index = 0) {
  if (explicitPath) {
    if (!existsSync(explicitPath)) {
      throw new OfficeError('FILE_NOT_FOUND', `指定的字体文件不存在：${explicitPath}`)
    }
    return { path: explicitPath, index, name: explicitPath.replace(/^.*[\\/]/, '').replace(/\.[^.]+$/, '') }
  }
  for (const candidate of CJK_FONT_CANDIDATES) {
    if (existsSync(candidate.path)) return candidate
  }
  throw new OfficeError(
    'UNSUPPORTED_FEATURE',
    '文本含非 WinAnsi 字符，需要嵌入字体，但本机没找到可用的中文字体（试过 SimHei / 等线 / 仿宋 / 楷体 / 微软雅黑 / 宋体 / Noto Sans SC）。可传 cjkFontPath 指定字体文件。',
    { tried: CJK_FONT_CANDIDATES.map((c) => c.path) }
  )
}

/**
 * 判断文本里是否有 WinAnsi 表示不了的字符（需要嵌入字体）。
 * @param {string} text - 文本。
 * @returns {boolean} 是否需要嵌入字体。
 */
export function needsEmbeddedFont(text) {
  for (const ch of String(text)) {
    if (!WIN_ANSI_CODES.has(ch.codePointAt(0))) return true
  }
  return false
}

/**
 * 找出文本里无法用 WinAnsi 表示的字符（去重，最多 10 个）。
 * @param {string} text - 文本。
 * @returns {string[]} 无法编码的字符。
 */
export function findUnencodableChars(text) {
  const bad = new Set()
  for (const ch of String(text)) {
    if (!WIN_ANSI_CODES.has(ch.codePointAt(0))) bad.add(ch)
    if (bad.size >= 10) break
  }
  return [...bad]
}

/**
 * 为一批文本构建「子集字体 + CID 映射 + 宽度 + ToUnicode」。
 *
 * CID 分配：按字符首次出现的顺序从 1 开始编号（0 保留），
 * `/CIDToGIDMap` 把这个 CID 映射回原字体里的字形编号，因此子集不需要重排字形。
 *
 * @param {object} args - 参数。
 * @param {string[]} args.texts - 会写进 PDF 的全部文本（用于收集字符）。
 * @param {string} [args.fontPath] - 指定的字体文件。
 * @param {number} [args.fontIndex] - `.ttc` 下标。
 * @returns {object} 嵌入所需的全部数据。
 */
export function buildPdfFont({ texts, fontPath = null, fontIndex = 0 }) {
  const location = resolveCjkFont(fontPath, fontIndex)
  const font = readFont(readFileSync(location.path), location.index)
  const lookup = readCmap(font.tableBuffer(0x636d6170))

  const cids = new Map()
  const codepoints = []
  const missing = new Set()
  for (const line of texts) {
    for (const ch of String(line)) {
      const code = ch.codePointAt(0)
      if (cids.has(code)) continue
      const gid = lookup(code)
      if (gid === 0 && !WIN_ANSI_CODES.has(code)) {
        missing.add(ch)
        continue
      }
      cids.set(code, cids.size + 1)
      codepoints.push({ code, gid, cid: cids.get(code) })
    }
  }
  if (missing.size > 0) {
    throw new OfficeError('UNSUPPORTED_FEATURE', `字体「${location.name}」缺少这些字符的字形：${[...missing].slice(0, 10).join(' ')}。`, {
      chars: [...missing].slice(0, 10),
      font: location.path
    })
  }
  if (codepoints.length === 0) {
    throw new OfficeError('INVALID_REQUEST', '没有需要嵌入的字符（文本为空或全在 WinAnsi 范围内）。')
  }

  const { buffer: subset, glyphCount } = subsetFont(font, codepoints.map((item) => item.gid), { name: location.name })
  const cidToGid = Buffer.alloc(2 * (codepoints.length + 1))
  // `/W` 必须是**真正的 PDF 数组**：`[cid [width] cid [width] …]`。
  // （拼成字符串会被序列化器当名称写，字体字典就非法了 —— 实测 Word 会因此读出 0 个字。）
  const widths = []
  for (const item of codepoints) {
    cidToGid.writeUInt16BE(item.gid, item.cid * 2)
    widths.push(item.cid, [Math.round((font.advance(item.gid) * 1000) / font.unitsPerEm)])
  }

  const toUnicodeLines = [
    '/CIDInit /ProcSet findresource begin',
    '12 dict begin',
    'begincmap',
    '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
    '/CMapName /Adobe-Identity-UCS def',
    '/CMapType 2 def',
    '1 begincodespacerange',
    '<0000> <FFFF>',
    'endcodespacerange'
  ]
  for (let i = 0; i < codepoints.length; i += 100) {
    const chunk = codepoints.slice(i, i + 100)
    toUnicodeLines.push(`${chunk.length} beginbfchar`)
    for (const item of chunk) toUnicodeLines.push(`<${item.cid.toString(16).padStart(4, '0')}> <${utf16Hex(item.code)}>`)
    toUnicodeLines.push('endbfchar')
  }
  toUnicodeLines.push('endcmap', 'CMapName currentdict /CMap defineresource pop', 'end', 'end')

  const scale = (value) => Math.round((value * 1000) / font.unitsPerEm)
  return {
    fontName: location.name,
    fontPath: location.path,
    // 子集前缀：PDF 规范建议嵌入子集用 6 个大写字母 + `+` 前缀
    pdfName: `${subsetPrefix(subset)}+${location.name}`,
    subset,
    glyphCount,
    charCount: codepoints.length,
    cidToGid,
    widths,
    toUnicode: toUnicodeLines.join('\n'),
    fontBBox: font.bbox.map(scale),
    ascent: scale(font.ascender),
    descent: scale(font.descender),
    unitsPerEm: font.unitsPerEm,
    /** 码点 → CID。 */
    cidOf(code) {
      return cids.has(code) ? cids.get(code) : null
    },
    /** 把文本编码成 Identity-H 的十六进制字符串（每个字符 2 字节 CID）。 */
    encode(text) {
      let hex = ''
      for (const ch of String(text)) {
        const cid = cids.get(ch.codePointAt(0))
        if (cid === undefined) {
          throw new OfficeError('UNSUPPORTED_FEATURE', `嵌入的字体「${location.name}」里没有字符「${ch}」的字形。`, {
            char: ch,
            code_point: `U+${ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')}`
          })
        }
        hex += cid.toString(16).padStart(4, '0')
      }
      return `<${hex}>`
    },
    /** 用真实字宽测量一段文本的宽度（点）。 */
    widthOf(text, fontSize) {
      let total = 0
      for (const ch of String(text)) {
        const cid = cids.get(ch.codePointAt(0))
        if (cid === undefined) {
          total += fontSize * 0.5
          continue
        }
        const gid = codepoints[cid - 1].gid
        total += (font.advance(gid) * fontSize) / font.unitsPerEm
      }
      return total
    }
  }
}

/**
 * 构造一份 Type0 字体所需的全部 PDF 对象描述。
 *
 * 返回的是 `{num, value}` 列表，调用方自己决定编号起点与引用方式。
 *
 * @param {object} font - {@link buildPdfFont} 的结果。
 * @param {number} startNum - 起始对象号。
 * @param {object} [options] - 选项。
 * @param {boolean} [options.compress] - 流是否 Flate 压缩，默认 true。
 * @param {(buffer: Buffer) => Buffer} [options.deflate] - 压缩函数（由调用方注入，避免重复依赖）。
 * @returns {{objects: object[], type0Num: number, cidFontNum: number, descriptorNum: number, fontFileNum: number, cidToGidNum: number, toUnicodeNum: number, nextNum: number}} 对象描述。
 */
export function buildType0FontObjects(font, startNum, { compress = true, deflate = (buffer) => buffer } = {}) {
  const type0Num = startNum
  const cidFontNum = startNum + 1
  const descriptorNum = startNum + 2
  const fontFileNum = startNum + 3
  const cidToGidNum = startNum + 4
  const toUnicodeNum = startNum + 5
  const stream = (buffer, extraDict = {}) => ({
    dict: compress ? { Filter: 'FlateDecode', ...extraDict } : { ...extraDict },
    stream: compress ? deflate(buffer) : buffer
  })
  const objects = [
    {
      num: type0Num,
      value: {
        Type: 'Font',
        Subtype: 'Type0',
        BaseFont: font.pdfName,
        Encoding: 'Identity-H',
        DescendantFonts: [{ ref: `${cidFontNum} 0` }],
        ToUnicode: { ref: `${toUnicodeNum} 0` }
      }
    },
    {
      num: cidFontNum,
      value: {
        Type: 'Font',
        Subtype: 'CIDFontType2',
        BaseFont: font.pdfName,
        CIDSystemInfo: {
          Registry: Buffer.from('Adobe', 'latin1'),
          Ordering: Buffer.from('Identity', 'latin1'),
          Supplement: 0
        },
        FontDescriptor: { ref: `${descriptorNum} 0` },
        DW: 1000,
        W: font.widths,
        CIDToGIDMap: { ref: `${cidToGidNum} 0` }
      }
    },
    {
      num: descriptorNum,
      value: {
        Type: 'FontDescriptor',
        FontName: font.pdfName,
        Flags: 4,
        FontBBox: font.fontBBox,
        ItalicAngle: 0,
        Ascent: font.ascent,
        Descent: font.descent,
        CapHeight: font.ascent,
        StemV: 80,
        FontFile2: { ref: `${fontFileNum} 0` }
      }
    },
    { num: fontFileNum, value: stream(font.subset, { Length1: font.subset.length }) },
    { num: cidToGidNum, value: stream(font.cidToGid) },
    { num: toUnicodeNum, value: stream(Buffer.from(font.toUnicode, 'latin1')) }
  ]
  return { objects, type0Num, cidFontNum, descriptorNum, fontFileNum, cidToGidNum, toUnicodeNum, nextNum: startNum + 6 }
}

/**
 * 子集前缀：按规范用 6 个大写字母（这里由子集字节的哈希稳定生成，便于复现）。
 * @param {Buffer} subset - 子集字体字节。
 * @returns {string} 6 个大写字母。
 */
function subsetPrefix(subset) {
  let hash = 0
  for (let i = 0; i < subset.length; i += 97) hash = (hash * 31 + subset[i]) >>> 0
  let prefix = ''
  let value = hash
  for (let i = 0; i < 6; i += 1) {
    prefix += String.fromCharCode(65 + (value % 26))
    value = Math.floor(value / 26)
  }
  return prefix
}

/**
 * 把一个码点写成 UTF-16BE 十六进制（含代理对）。
 * @param {number} code - 码点。
 * @returns {string} 十六进制字符串。
 */
function utf16Hex(code) {
  if (code <= 0xffff) return code.toString(16).padStart(4, '0').toUpperCase()
  const value = code - 0x10000
  const high = 0xd800 + (value >> 10)
  const low = 0xdc00 + (value & 0x3ff)
  return `${high.toString(16).padStart(4, '0')}${low.toString(16).padStart(4, '0')}`.toUpperCase()
}
