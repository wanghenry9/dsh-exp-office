/**
 * TrueType 字体解析与**保留 GID 的子集化**（用于把中文写进我们自己生成的 PDF）。
 *
 * 为什么需要：PDF 标准 14 字体没有汉字字形，要写中文必须嵌入字体。整个中文字体有
 * 10–20 MB，直接整份嵌入会让每个 PDF 都胖到十几 MB，因此这里做子集化。
 *
 * 关键设计：**保留原始 GID 编号**。
 *   - PDF 的 `CIDFontType2` 用 `/CIDToGIDMap` 把 CID 映射到字形编号，因此只要子集里
 *     字形编号不变，就不需要重排编号（重排编号是子集化最容易出错的部分）；
 *   - 做法是保留 `loca` 的长度（numGlyphs + 1 项），未用到的字形写成**零长度**，
 *     这在 TrueType 里是合法的（就是所谓的空字形）；
 *   - 复合字形（带重音的拉丁字母、部分汉字）会被递归闭包进来，否则渲染时缺笔画。
 *
 * 支持：普通 `.ttf` 与 `.ttc`（字体集合，取指定索引）。
 * 不支持（会明确报错）：CFF/OpenType（`OTTO`）、没有 `glyf`/`loca` 的字体、可变字体的非默认实例。
 *
 * @module dsh-exp-office/font-subset
 */

import { OfficeError } from './errors.js'

/** sfnt 表标签 → 名称（仅用于报错信息）。 */
const TAG_GLYF = 0x676c7966
const TAG_LOCA = 0x6c6f6361
const TAG_HEAD = 0x68656164
const TAG_HHEA = 0x68686561
const TAG_MAXP = 0x6d617870
const TAG_HMTX = 0x686d7478
const TAG_CMAP = 0x636d6170
const TAG_OS2 = 0x4f532f32
const TAG_POST = 0x706f7374
const TAG_NAME = 0x6e616d65
const TAG_TTCF = 0x74746366
const TAG_OTTO = 0x4f54544f

/**
 * 读取一个 sfnt（.ttf 或 .ttc 中的一份字体）。
 *
 * @param {Buffer} buffer - 字体文件字节。
 * @param {number} [fontIndex] - `.ttc` 里的字体下标（默认 0）。
 * @returns {object} 解析结果：表目录与常用表。
 */
export function readFont(buffer, fontIndex = 0) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 12) {
    throw new OfficeError('UNSUPPORTED_FEATURE', '字体文件太小，不是有效的 TrueType 字体。')
  }
  let dirOffset = 0
  const tag = buffer.readUInt32BE(0)
  if (tag === TAG_TTCF) {
    const numFonts = buffer.readUInt32BE(8)
    if (fontIndex >= numFonts) {
      throw new OfficeError('INVALID_REQUEST', `字体集合只有 ${numFonts} 份字体，取不到第 ${fontIndex} 份。`)
    }
    dirOffset = buffer.readUInt32BE(12 + fontIndex * 4)
  } else if (tag === TAG_OTTO) {
    throw new OfficeError('UNSUPPORTED_FEATURE', '这是 OpenType/CFF 字体（OTTO），当前只支持 TrueType 轮廓（glyf/loca）。')
  } else if (tag !== 0x00010000 && tag !== 0x74727565) {
    throw new OfficeError('UNSUPPORTED_FEATURE', `无法识别的字体格式（sfnt 版本 0x${tag.toString(16)}）。`)
  }

  const numTables = buffer.readUInt16BE(dirOffset + 4)
  const tables = new Map()
  for (let i = 0; i < numTables; i += 1) {
    const record = dirOffset + 12 + i * 16
    if (record + 16 > buffer.length) break
    const tableTag = buffer.readUInt32BE(record)
    tables.set(tableTag, {
      offset: buffer.readUInt32BE(record + 8),
      length: buffer.readUInt32BE(record + 12)
    })
  }
  for (const required of [TAG_GLYF, TAG_LOCA, TAG_HEAD, TAG_HHEA, TAG_MAXP, TAG_HMTX]) {
    if (!tables.has(required)) {
      throw new OfficeError('UNSUPPORTED_FEATURE', `字体缺少必需的表 0x${required.toString(16)}，无法用于 PDF 嵌入。`)
    }
  }

  const tableBuffer = (tableTag) => {
    const entry = tables.get(tableTag)
    if (!entry) return null
    if (entry.offset + entry.length > buffer.length) {
      throw new OfficeError('CORRUPTED_DOCUMENT', `字体表 0x${tableTag.toString(16)} 越界。`)
    }
    return buffer.subarray(entry.offset, entry.offset + entry.length)
  }

  const head = tableBuffer(TAG_HEAD)
  const maxp = tableBuffer(TAG_MAXP)
  const hhea = tableBuffer(TAG_HHEA)
  const unitsPerEm = head.readUInt16BE(18)
  const indexToLocFormat = head.readInt16BE(50)
  const numGlyphs = maxp.readUInt16BE(4)
  const numberOfHMetrics = hhea.readUInt16BE(34)

  return {
    buffer,
    fontIndex,
    tables,
    tableBuffer,
    unitsPerEm,
    indexToLocFormat,
    numGlyphs,
    numberOfHMetrics,
    ascender: hhea.readInt16BE(4),
    descender: hhea.readInt16BE(6),
    bbox: [head.readInt16BE(36), head.readInt16BE(38), head.readInt16BE(40), head.readInt16BE(42)],
    /** 取字形数据（空字形返回长度 0 的切片）。 */
    glyph(gid) {
      const loca = readLoca(tableBuffer(TAG_LOCA), numGlyphs, indexToLocFormat)
      if (gid + 1 >= loca.length) return Buffer.alloc(0)
      const start = loca[gid]
      const end = loca[gid + 1]
      if (end <= start) return Buffer.alloc(0)
      const glyf = tableBuffer(TAG_GLYF)
      return glyf.subarray(start, end)
    },
    /** 字形推进宽度（字体单位）。 */
    advance(gid) {
      const hmtx = tableBuffer(TAG_HMTX)
      const index = gid < numberOfHMetrics ? gid : numberOfHMetrics - 1
      if (index < 0) return 0
      const at = index * 4
      return at + 2 <= hmtx.length ? hmtx.readUInt16BE(at) : 0
    }
  }
}

/**
 * 解析 `loca` 表。
 * @param {Buffer} loca - 表字节。
 * @param {number} numGlyphs - 字形数。
 * @param {number} format - `head.indexToLocFormat`（0=短格式，1=长格式）。
 * @returns {number[]} 长度 `numGlyphs + 1` 的偏移数组。
 */
function readLoca(loca, numGlyphs, format) {
  const out = new Array(numGlyphs + 1).fill(0)
  for (let i = 0; i <= numGlyphs; i += 1) {
    if (format === 0) {
      out[i] = 2 * (i * 2 + 2 <= loca.length ? loca.readUInt16BE(i * 2) : 0)
    } else {
      out[i] = i * 4 + 4 <= loca.length ? loca.readUInt32BE(i * 4) : 0
    }
  }
  return out
}

/**
 * 解析 `cmap`，返回「码点 → 字形编号」的查找函数。
 *
 * 优先 format 12（覆盖非 BMP），退回 format 4。
 *
 * @param {Buffer|null} cmap - `cmap` 表字节。
 * @returns {(codePoint: number) => number} 查找函数（找不到返回 0）。
 */
export function readCmap(cmap) {
  if (!cmap || cmap.length < 4) return () => 0
  const numTables = cmap.readUInt16BE(2)
  let best = null
  let bestScore = -1
  for (let i = 0; i < numTables; i += 1) {
    const record = 4 + i * 8
    if (record + 8 > cmap.length) break
    const platform = cmap.readUInt16BE(record)
    const encoding = cmap.readUInt16BE(record + 2)
    const offset = cmap.readUInt32BE(record + 4)
    if (offset + 2 > cmap.length) continue
    const format = cmap.readUInt16BE(offset)
    let score = -1
    if (format === 12) score = 3
    else if (format === 4) score = 2
    else if (format === 6) score = 1
    if (score < 0) continue
    // 偏好 Unicode 平台（0）与 Windows BMP/完整（3/1、3/10）
    if (platform === 3 && (encoding === 10 || encoding === 1)) score += 2
    if (platform === 0) score += 1
    if (score > bestScore) {
      bestScore = score
      best = { offset, format }
    }
  }
  if (!best) return () => 0

  if (best.format === 12) {
    const offset = best.offset
    const groups = cmap.readUInt32BE(offset + 12)
    return (codePoint) => {
      let lo = 0
      let hi = groups - 1
      while (lo <= hi) {
        const mid = (lo + hi) >> 1
        const at = offset + 16 + mid * 12
        if (at + 12 > cmap.length) break
        const start = cmap.readUInt32BE(at)
        const end = cmap.readUInt32BE(at + 4)
        if (codePoint < start) hi = mid - 1
        else if (codePoint > end) lo = mid + 1
        else return cmap.readUInt32BE(at + 8) + (codePoint - start)
      }
      return 0
    }
  }
  if (best.format === 4) {
    const offset = best.offset
    const segCountX2 = cmap.readUInt16BE(offset + 6)
    const segCount = segCountX2 / 2
    const endAt = offset + 14
    const startAt = endAt + segCountX2 + 2
    const deltaAt = startAt + segCountX2
    const rangeAt = deltaAt + segCountX2
    return (codePoint) => {
      if (codePoint > 0xffff) return 0
      for (let s = 0; s < segCount; s += 1) {
        const end = cmap.readUInt16BE(endAt + s * 2)
        if (codePoint > end) continue
        const start = cmap.readUInt16BE(startAt + s * 2)
        if (codePoint < start) return 0
        const delta = cmap.readInt16BE(deltaAt + s * 2)
        const rangeOffset = cmap.readUInt16BE(rangeAt + s * 2)
        if (rangeOffset === 0) return (codePoint + delta) & 0xffff
        const at = rangeAt + s * 2 + rangeOffset + (codePoint - start) * 2
        if (at + 2 > cmap.length) return 0
        const gid = cmap.readUInt16BE(at)
        return gid === 0 ? 0 : (gid + delta) & 0xffff
      }
      return 0
    }
  }
  if (best.format === 6) {
    const offset = best.offset
    const first = cmap.readUInt16BE(offset + 6)
    const count = cmap.readUInt16BE(offset + 8)
    return (codePoint) => {
      const index = codePoint - first
      if (index < 0 || index >= count) return 0
      const at = offset + 10 + index * 2
      return at + 2 <= cmap.length ? cmap.readUInt16BE(at) : 0
    }
  }
  return () => 0
}

/**
 * 收集一组字形编号的**闭包**（复合字形引用的子字形也要包含进来）。
 *
 * @param {object} font - {@link readFont} 的结果。
 * @param {Iterable<number>} gids - 初始字形编号。
 * @returns {Set<number>} 闭包后的字形集合。
 */
export function glyphClosure(font, gids) {
  const used = new Set()
  const queue = [...gids]
  while (queue.length > 0) {
    const gid = queue.pop()
    if (used.has(gid) || gid < 0 || gid >= font.numGlyphs) continue
    used.add(gid)
    const data = font.glyph(gid)
    if (data.length < 10) continue
    const contours = data.readInt16BE(0)
    if (contours >= 0) continue
    // 复合字形：逐个组件取 glyphIndex
    let at = 10
    while (at + 4 <= data.length) {
      const flags = data.readUInt16BE(at)
      const componentGid = data.readUInt16BE(at + 2)
      queue.push(componentGid)
      at += 4
      at += flags & 0x0001 ? 4 : 2 // ARG_1_AND_2_ARE_WORDS
      if (flags & 0x0008) at += 2 // WE_HAVE_A_SCALE
      else if (flags & 0x0040) at += 4 // X_AND_Y_SCALE
      else if (flags & 0x0080) at += 8 // TWO_BY_TWO
      if (!(flags & 0x0020)) break // MORE_COMPONENTS
    }
  }
  return used
}

/**
 * 生成一份**保留原始 GID 编号**的 TrueType 子集。
 *
 * @param {object} font - {@link readFont} 的结果。
 * @param {Iterable<number>} gids - 需要保留的字形编号（会做复合闭包）。
 * @param {object} [options] - 选项。
 * @param {string} [options.name] - 子集字体名（写入 `name` 表）。
 * @returns {{buffer: Buffer, glyphCount: number, byteLength: number}} 子集字体。
 */
export function subsetFont(font, gids, { name = 'DshSubset' } = {}) {
  const used = glyphClosure(font, gids)
  // 字形编号必须保留，但**编号高于最大用到字形**的部分可以整段砍掉：
  // 没有引用会指向它们（loca/hmtx 跟着变短，子集体积才可控）。
  let highest = 0
  for (const gid of used) highest = Math.max(highest, gid)
  const numGlyphs = Math.max(1, Math.min(font.numGlyphs, highest + 1))

  // 1) 重建 glyf / loca（未用字形为零长度）
  const glyfParts = []
  const loca = new Array(numGlyphs + 1).fill(0)
  let offset = 0
  for (let gid = 0; gid < numGlyphs; gid += 1) {
    loca[gid] = offset
    if (!used.has(gid)) continue
    const source = font.glyph(gid)
    if (source.length === 0) continue
    // glyf 里的每个字形都要 4 字节对齐（并补零）
    const padded = Buffer.alloc(Math.ceil(source.length / 4) * 4)
    source.copy(padded)
    glyfParts.push(padded)
    offset += padded.length
  }
  loca[numGlyphs] = offset
  const glyf = Buffer.concat(glyfParts)
  // 短格式偏移是「真实偏移 / 2」，能省一半：glyf 小于 128 KB 时用短格式
  const shortLoca = offset < 0x20000
  const locaBuffer = Buffer.alloc((numGlyphs + 1) * (shortLoca ? 2 : 4))
  for (let i = 0; i <= numGlyphs; i += 1) {
    if (shortLoca) locaBuffer.writeUInt16BE(loca[i] / 2, i * 2)
    else locaBuffer.writeUInt32BE(loca[i], i * 4)
  }

  // 2) head：按 loca 格式改写
  const head = Buffer.from(font.tableBuffer(TAG_HEAD))
  head.writeInt16BE(shortLoca ? 0 : 1, 50) // indexToLocFormat
  head.writeUInt32BE(0, 8) // checkSumAdjustment 稍后统一算

  // 3) maxp / hhea / hmtx：字形数变短，度量表同步裁剪
  const maxp = Buffer.from(font.tableBuffer(TAG_MAXP))
  maxp.writeUInt16BE(numGlyphs, 4)
  const numberOfHMetrics = Math.max(1, Math.min(font.numberOfHMetrics, numGlyphs))
  const hhea = Buffer.from(font.tableBuffer(TAG_HHEA))
  hhea.writeUInt16BE(numberOfHMetrics, 34)
  const sourceHmtx = font.tableBuffer(TAG_HMTX)
  const hmtx = Buffer.alloc(numberOfHMetrics * 4 + (numGlyphs - numberOfHMetrics) * 2)
  for (let i = 0; i < numberOfHMetrics; i += 1) {
    const at = i * 4
    if (at + 4 > sourceHmtx.length) break
    hmtx.writeUInt16BE(sourceHmtx.readUInt16BE(at), at)
    hmtx.writeInt16BE(sourceHmtx.readInt16BE(at + 2), at + 2)
  }
  for (let i = numberOfHMetrics; i < numGlyphs; i += 1) {
    const at = numberOfHMetrics * 4 + (i - numberOfHMetrics) * 2
    if (at + 2 > sourceHmtx.length) break
    hmtx.writeInt16BE(sourceHmtx.readInt16BE(at), numberOfHMetrics * 4 + (i - numberOfHMetrics) * 2)
  }

  // 4) cmap：为用到的码点重建一个最小 format 4（PDF 用 CIDToGIDMap，不依赖它，
  //    但留一个合法的 cmap 能让别的工具也认这份字体）
  const cmapTable = buildMinimalCmap(font, used)

  // 5) post 3.0（不含字形名）+ 最小 name
  const post = Buffer.alloc(32)
  post.writeUInt32BE(0x00030000, 0)
  const nameTable = buildMinimalName(name)

  const tables = [
    { tag: TAG_HEAD, data: head },
    { tag: TAG_HHEA, data: hhea },
    { tag: TAG_MAXP, data: maxp },
    { tag: TAG_HMTX, data: hmtx },
    { tag: TAG_LOCA, data: locaBuffer },
    { tag: TAG_GLYF, data: glyf },
    { tag: TAG_CMAP, data: cmapTable },
    { tag: TAG_POST, data: post },
    { tag: TAG_NAME, data: nameTable }
  ]
  const os2 = font.tableBuffer(TAG_OS2)
  if (os2) tables.push({ tag: TAG_OS2, data: Buffer.from(os2) })

  const buffer = assembleSfnt(tables)
  return { buffer, glyphCount: used.size, glyphTotal: numGlyphs, byteLength: buffer.length }
}

/**
 * 组装一个 sfnt 文件（表目录 + 表数据 + head 校验和修正）。
 * @param {object[]} tables - `{tag, data}` 列表。
 * @returns {Buffer} 文件字节。
 */
function assembleSfnt(tables) {
  const count = tables.length
  const searchRange = 2 ** Math.floor(Math.log2(count)) * 16
  const entrySelector = Math.floor(Math.log2(count))
  const rangeShift = count * 16 - searchRange

  const header = Buffer.alloc(12)
  header.writeUInt32BE(0x00010000, 0)
  header.writeUInt16BE(count, 4)
  header.writeUInt16BE(searchRange, 6)
  header.writeUInt16BE(entrySelector, 8)
  header.writeUInt16BE(rangeShift, 10)

  const directory = Buffer.alloc(count * 16)
  const chunks = []
  let offset = 12 + count * 16
  const sorted = [...tables].sort((a, b) => a.tag - b.tag)
  sorted.forEach((table, index) => {
    const padded = Buffer.alloc(Math.ceil(table.data.length / 4) * 4)
    table.data.copy(padded)
    const at = index * 16
    directory.writeUInt32BE(table.tag, at)
    directory.writeUInt32BE(checksum(padded), at + 4)
    directory.writeUInt32BE(offset, at + 8)
    directory.writeUInt32BE(table.data.length, at + 12)
    chunks.push(padded)
    table.__offset = offset
    offset += padded.length
  })

  const body = Buffer.concat([header, directory, ...chunks])
  const headEntry = sorted.find((table) => table.tag === 0x68656164)
  if (headEntry) {
    // head.checkSumAdjustment = 0xB1B0AFBA - 整个文件的校验和（此时该字段按 0 计算）
    const adjustment = (0xb1b0afba - checksum(body)) >>> 0
    body.writeUInt32BE(adjustment, headEntry.__offset + 8)
  }
  return body
}

/**
 * 计算 sfnt 校验和（按 4 字节大端求和）。
 * @param {Buffer} data - 数据。
 * @returns {number} 32 位校验和。
 */
function checksum(data) {
  let sum = 0
  for (let i = 0; i < data.length; i += 4) {
    const word = ((data[i] ?? 0) << 24) | ((data[i + 1] ?? 0) << 16) | ((data[i + 2] ?? 0) << 8) | (data[i + 3] ?? 0)
    sum = (sum + word) >>> 0
  }
  return sum >>> 0
}

/**
 * 为一个字形集合构造最小 `cmap`（format 4，BMP 码点）。
 * @param {object} font - 字体。
 * @param {Set<number>} used - 保留的字形集合。
 * @returns {Buffer} `cmap` 表。
 */
function buildMinimalCmap(font, used) {
  const cmap = font.tableBuffer(TAG_CMAP)
  const mapping = []
  if (cmap) {
    const lookup = readCmap(cmap)
    // 只反查「用到的字形」对应的码点：遍历常用区间即可（BMP + 基本扩展区）
    const ranges = [
      [0x20, 0x2fff],
      [0x3000, 0x9fff],
      [0xf900, 0xfaff],
      [0xff00, 0xffef]
    ]
    for (const [from, to] of ranges) {
      for (let code = from; code <= to; code += 1) {
        const gid = lookup(code)
        if (gid !== 0 && used.has(gid)) mapping.push([code, gid])
      }
    }
  }
  // format 4 段：每段 [startCode, endCode, idDelta]
  const segments = []
  let current = null
  for (const [code, gid] of mapping) {
    const delta = (gid - code) & 0xffff
    if (current && code === current.end + 1 && delta === current.delta) {
      current.end = code
      continue
    }
    current = { start: code, end: code, delta }
    segments.push(current)
  }
  segments.push({ start: 0xffff, end: 0xffff, delta: 1 }) // 必需的结束段
  const segCount = segments.length
  const length = 16 + segCount * 8
  const table = Buffer.alloc(12 + length) // 12 = cmap 头（版本 + 一张编码记录）
  table.writeUInt16BE(0, 0) // version
  table.writeUInt16BE(1, 2) // numTables
  table.writeUInt16BE(3, 4) // platformID = Windows
  table.writeUInt16BE(1, 6) // encodingID = BMP
  table.writeUInt32BE(12, 8) // offset
  const base = 12
  table.writeUInt16BE(4, base)
  table.writeUInt16BE(length, base + 2)
  table.writeUInt16BE(0, base + 4)
  table.writeUInt16BE(segCount * 2, base + 6)
  const searchRange = 2 ** Math.floor(Math.log2(segCount)) * 2
  table.writeUInt16BE(searchRange, base + 8)
  table.writeUInt16BE(Math.floor(Math.log2(segCount)), base + 10)
  table.writeUInt16BE(segCount * 2 - searchRange, base + 12)
  const endAt = base + 14
  const startAt = endAt + segCount * 2 + 2
  const deltaAt = startAt + segCount * 2
  const rangeAt = deltaAt + segCount * 2
  segments.forEach((segment, index) => {
    table.writeUInt16BE(segment.end, endAt + index * 2)
    table.writeUInt16BE(segment.start, startAt + index * 2)
    table.writeUInt16BE(segment.delta, deltaAt + index * 2)
    table.writeUInt16BE(0, rangeAt + index * 2)
  })
  return table.subarray(0, rangeAt + segCount * 2)
}

/**
 * 构造最小 `name` 表（只写字体名，满足「有 name 表」的最低要求）。
 * @param {string} familyName - 字体名。
 * @returns {Buffer} `name` 表。
 */
function buildMinimalName(familyName) {
  const ascii = Buffer.from(familyName, 'latin1')
  const records = [
    { id: 1, value: ascii }, // Family
    { id: 2, value: ascii }, // Subfamily
    { id: 4, value: ascii }, // Full name
    { id: 6, value: ascii } // PostScript name
  ]
  const count = records.length
  const stringOffset = 6 + count * 12
  const header = Buffer.alloc(6)
  header.writeUInt16BE(0, 0)
  header.writeUInt16BE(count, 2)
  header.writeUInt16BE(stringOffset, 4)
  const table = Buffer.alloc(stringOffset + ascii.length * count)
  header.copy(table, 0)
  records.forEach((record, index) => {
    const at = 6 + index * 12
    table.writeUInt16BE(3, at) // Windows
    table.writeUInt16BE(1, at + 2) // BMP
    table.writeUInt16BE(0x409, at + 4) // en-US
    table.writeUInt16BE(record.id, at + 6)
    table.writeUInt16BE(record.value.length, at + 8)
    table.writeUInt16BE(index * ascii.length, at + 10)
  })
  for (let i = 0; i < count; i += 1) ascii.copy(table, stringOffset + i * ascii.length)
  return table
}
