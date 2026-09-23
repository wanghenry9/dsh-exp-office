/**
 * OOXML 底层引擎：ZIP 容器读写 + 支持字节级最小修改的 XML 解析器。
 *
 * 设计要点（对齐开发要求 §六.4 / §十二.1 / §十八.3）：
 *   - 最小修改模式：解析器为每个节点记录源码偏移量，编辑只产生「区间补丁」，
 *     未被触碰的字节原样输出。因此修改单元格不会破坏图表、宏、外部链接、
 *     母版、批注等任何无关部件。
 *   - 零运行时依赖：ZIP 容器用 node:zlib 实现，XML 用自写扫描器。
 *   - 安全：拒绝 DOCTYPE / 外部实体（XML 实体攻击），ZIP 限制条目数、单条大小、
 *     总解压大小与压缩比（ZIP 炸弹），并限制 XML 字节数、节点数与嵌套深度。
 *
 * @module dsh-exp-office/ooxml
 */

import zlib from 'node:zlib'
import { OfficeError } from './errors.js'

const SIG_LOCAL = 0x04034b50
const SIG_CENTRAL = 0x02014b50
const SIG_EOCD = 0x06054b50
const SIG_EOCD64 = 0x06064b50
const SIG_EOCD64_LOCATOR = 0x07064b50

/** ZIP 安全上限（docs §十二.1 压缩炸弹与大小限制）。 */
export const DEFAULT_ZIP_LIMITS = Object.freeze({
  maxEntries: 8192,
  maxEntryBytes: 256 * 1024 * 1024,
  maxTotalBytes: 1024 * 1024 * 1024,
  maxRatio: 200
})

/** XML 安全上限（docs §十二.1 限制 XML 节点数量和递归深度）。 */
export const DEFAULT_XML_LIMITS = Object.freeze({
  maxBytes: 64 * 1024 * 1024,
  maxNodes: 2_000_000,
  maxDepth: 256
})

/** 仅允许 XML 预定义实体；DTD 与外部实体一律拒绝（防实体展开攻击）。 */
const NAMED_ENTITIES = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" }

/**
 * 解码 XML 文本中的实体引用。只处理预定义实体与数字字符引用，
 * 未知实体原样保留而绝不展开 —— 因此不存在实体展开攻击面。
 * @param {string} text - 原始文本。
 * @returns {string} 解码后的文本。
 */
export function decodeEntities(text) {
  if (!text.includes('&')) return text
  return text.replace(/&(#x?[0-9A-Fa-f]+|[A-Za-z][\w.-]*);/g, (match, body) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' || body[1] === 'X' ? Number.parseInt(body.slice(2), 16) : Number.parseInt(body.slice(1), 10)
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return match
      return String.fromCodePoint(code)
    }
    return Object.hasOwn(NAMED_ENTITIES, body) ? NAMED_ENTITIES[body] : match
  })
}

/**
 * 转义 XML 文本内容。
 * @param {string} text - 待转义文本。
 * @returns {string} 可安全置入元素体的文本。
 */
export function escapeXmlText(text) {
  return String(text).replace(/[&<>]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : '&gt;'))
}

/**
 * 转义 XML 属性值。
 * @param {string} text - 待转义文本。
 * @returns {string} 可安全置入双引号属性的文本。
 */
export function escapeXmlAttr(text) {
  return String(text).replace(/[&<>"]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&quot;'))
}

/**
 * 读取 32 位小端无符号整数。
 * @param {Buffer} buf - 数据缓冲。
 * @param {number} off - 偏移。
 * @returns {number} 数值。
 */
const u16 = (buf, off) => buf.readUInt16LE(off)
const u32 = (buf, off) => buf.readUInt32LE(off)
const u64 = (buf, off) => Number(buf.readBigUInt64LE(off))

/**
 * 一个已打开的 OPC（ZIP）包。条目按需解压并缓存，写入时未修改的条目
 * 直接复用原有压缩字节，保证除目标部件外逐字节不变。
 */
export class ZipPackage {
  #buffer
  #entries
  #limits
  #cache = new Map()
  #deleted = new Set()

  /**
   * @param {Buffer} buffer - 完整 ZIP 字节。
   * @param {Map<string, object>} entries - 中央目录条目。
   * @param {object} limits - 安全上限。
   */
  constructor(buffer, entries, limits) {
    this.#buffer = buffer
    this.#entries = entries
    this.#limits = limits
  }

  /**
   * 打开一个 ZIP 包。
   * @param {Buffer} buffer - ZIP 字节。
   * @param {Partial<typeof DEFAULT_ZIP_LIMITS>} [limitOverrides] - 上限覆盖。
   * @returns {ZipPackage} 包对象。
   */
  static open(buffer, limitOverrides = {}) {
    const limits = { ...DEFAULT_ZIP_LIMITS, ...limitOverrides }
    if (!Buffer.isBuffer(buffer) || buffer.length < 22) {
      throw new OfficeError('CORRUPTED_DOCUMENT', '文件不是有效的 ZIP 包（长度不足）。')
    }
    const eocd = ZipPackage.#findEocd(buffer)
    let entryCount = u16(buffer, eocd + 10)
    let cdOffset = u32(buffer, eocd + 16)
    let cdSize = u32(buffer, eocd + 12)

    // ZIP64：条目数或偏移量溢出 32 位时改读 ZIP64 记录。
    const needsZip64 = entryCount === 0xffff || cdOffset === 0xffffffff || cdSize === 0xffffffff
    if (needsZip64) {
      const zip64 = ZipPackage.#readZip64(buffer, eocd)
      if (zip64) {
        entryCount = zip64.entryCount
        cdOffset = zip64.cdOffset
        cdSize = zip64.cdSize
      }
    }
    if (entryCount > limits.maxEntries) {
      throw new OfficeError('MEMORY_LIMIT', `ZIP 条目数 ${entryCount} 超过上限 ${limits.maxEntries}。`, { entry_count: entryCount })
    }
    if (cdOffset + cdSize > buffer.length) {
      throw new OfficeError('CORRUPTED_DOCUMENT', 'ZIP 中央目录越界，文件已损坏。')
    }

    const entries = new Map()
    let cursor = cdOffset
    let totalUncompressed = 0
    for (let i = 0; i < entryCount; i += 1) {
      if (cursor + 46 > buffer.length || u32(buffer, cursor) !== SIG_CENTRAL) {
        throw new OfficeError('CORRUPTED_DOCUMENT', `ZIP 中央目录第 ${i} 项签名无效。`, { index: i })
      }
      const flags = u16(buffer, cursor + 8)
      const method = u16(buffer, cursor + 10)
      const modTime = u16(buffer, cursor + 12)
      const modDate = u16(buffer, cursor + 14)
      const crc = u32(buffer, cursor + 16)
      let compSize = u32(buffer, cursor + 20)
      let uncompSize = u32(buffer, cursor + 24)
      const nameLen = u16(buffer, cursor + 28)
      const extraLen = u16(buffer, cursor + 30)
      const commentLen = u16(buffer, cursor + 32)
      let localOffset = u32(buffer, cursor + 42)
      const name = buffer.toString('utf8', cursor + 46, cursor + 46 + nameLen)

      // ZIP64 扩展字段（0x0001）：补齐被置为 0xFFFFFFFF 的字段。
      if (uncompSize === 0xffffffff || compSize === 0xffffffff || localOffset === 0xffffffff) {
        const extraStart = cursor + 46 + nameLen
        let p = extraStart
        const extraEnd = extraStart + extraLen
        while (p + 4 <= extraEnd) {
          const id = u16(buffer, p)
          const size = u16(buffer, p + 2)
          if (id === 0x0001) {
            let q = p + 4
            if (uncompSize === 0xffffffff) { uncompSize = u64(buffer, q); q += 8 }
            if (compSize === 0xffffffff) { compSize = u64(buffer, q); q += 8 }
            if (localOffset === 0xffffffff) { localOffset = u64(buffer, q); q += 8 }
            break
          }
          p += 4 + size
        }
      }

      if ((flags & 0x0001) !== 0) {
        entries.set(name, { name, method, flags, crc, compSize, uncompSize, localOffset, modTime, modDate, encrypted: true })
      } else {
        if (method !== 0 && method !== 8) {
          throw new OfficeError('UNSUPPORTED_FEATURE', `ZIP 使用了不支持的压缩方法 ${method}。`, { entry: name, method })
        }
        if (uncompSize > limits.maxEntryBytes) {
          throw new OfficeError('MEMORY_LIMIT', `条 ${name} 解压后 ${uncompSize} 字节超过单条上限。`, { entry: name })
        }
        totalUncompressed += uncompSize
        entries.set(name, { name, method, flags, crc, compSize, uncompSize, localOffset, modTime, modDate, encrypted: false })
      }
      cursor += 46 + nameLen + extraLen + commentLen
    }
    if (totalUncompressed > limits.maxTotalBytes) {
      throw new OfficeError('MEMORY_LIMIT', `ZIP 解压总量 ${totalUncompressed} 字节超过上限 ${limits.maxTotalBytes}。`, {
        total_bytes: totalUncompressed
      })
    }
    return new ZipPackage(buffer, entries, limits)
  }

  /**
   * 从文件尾部定位 EOCD 记录。
   * @param {Buffer} buffer - ZIP 字节。
   * @returns {number} EOCD 起始偏移。
   */
  static #findEocd(buffer) {
    const maxComment = 0xffff
    const start = Math.max(0, buffer.length - maxComment - 22)
    for (let i = buffer.length - 22; i >= start; i -= 1) {
      if (u32(buffer, i) === SIG_EOCD) return i
    }
    throw new OfficeError('CORRUPTED_DOCUMENT', '未找到 ZIP 结束记录（EOCD），文件可能被截断。')
  }

  /**
   * 读取 ZIP64 结束记录。
   * @param {Buffer} buffer - ZIP 字节。
   * @param {number} eocd - EOCD 偏移。
   * @returns {{entryCount: number, cdOffset: number, cdSize: number}|null} ZIP64 信息。
   */
  static #readZip64(buffer, eocd) {
    const locator = eocd - 20
    if (locator < 0 || u32(buffer, locator) !== SIG_EOCD64_LOCATOR) return null
    const record = u64(buffer, locator + 8)
    if (record + 56 > buffer.length || u32(buffer, record) !== SIG_EOCD64) return null
    return {
      entryCount: u64(buffer, record + 32),
      cdSize: u64(buffer, record + 40),
      cdOffset: u64(buffer, record + 48)
    }
  }

  /** @returns {string[]} 包内所有条目名。 */
  names() {
    return [...this.#entries.keys()]
  }

  /**
   * 判断条目是否存在。
   * @param {string} name - 条目名。
   * @returns {boolean} 是否存在。
   */
  has(name) {
    return this.#entries.has(name) && !this.#deleted.has(name)
  }

  /**
   * 取出条目原始（未解压）字节，用于 `toBuffer()` 时的透传。
   * @param {string} name - 条目名。
   * @returns {Buffer} 压缩后的字节片段。
   */
  #rawOf(name) {
    const entry = this.#entries.get(name)
    const header = entry.localOffset
    if (u32(this.#buffer, header) !== SIG_LOCAL) {
      throw new OfficeError('CORRUPTED_DOCUMENT', `条目 ${name} 的本地文件头签名无效。`)
    }
    const nameLen = u16(this.#buffer, header + 26)
    const extraLen = u16(this.#buffer, header + 28)
    const dataStart = header + 30 + nameLen + extraLen
    return this.#buffer.subarray(dataStart, dataStart + entry.compSize)
  }

  /**
   * 读取并解压一个条目。
   * @param {string} name - 条目名。
   * @returns {Buffer} 解压后的字节。
   */
  read(name) {
    if (this.#deleted.has(name)) throw new OfficeError('FILE_NOT_FOUND', `包内不存在条目 ${name}。`, { entry: name })
    const cached = this.#cache.get(name)
    if (cached) return cached
    const entry = this.#entries.get(name)
    if (!entry) throw new OfficeError('FILE_NOT_FOUND', `包内不存在条目 ${name}。`, { entry: name })
    if (entry.encrypted) throw new OfficeError('PASSWORD_REQUIRED', `条目 ${name} 已加密，无法读取。`, { entry: name })

    const raw = this.#rawOf(name)
    let data
    if (entry.method === 0) {
      data = Buffer.from(raw)
    } else {
      if (entry.compSize > 0 && entry.uncompSize / entry.compSize > this.#limits.maxRatio && entry.uncompSize > 1024 * 1024) {
        throw new OfficeError('MEMORY_LIMIT', `条目 ${name} 压缩比异常（疑似压缩炸弹），已拒绝解压。`, {
          entry: name,
          ratio: entry.uncompSize / entry.compSize
        })
      }
      try {
        data = zlib.inflateRawSync(raw, { maxOutputLength: Math.min(entry.uncompSize + 1024, this.#limits.maxEntryBytes) })
      } catch (err) {
        throw new OfficeError('CORRUPTED_DOCUMENT', `条目 ${name} 解压失败：${err.message}`, { entry: name })
      }
    }
    if (data.length !== entry.uncompSize) {
      throw new OfficeError('CORRUPTED_DOCUMENT', `条目 ${name} 解压长度不符（期望 ${entry.uncompSize}，实际 ${data.length}）。`, { entry: name })
    }
    if (this.#cache.size < 256) this.#cache.set(name, data)
    return data
  }

  /**
   * 读取一个条目并按 UTF-8 解码为文本。去掉 BOM。
   * @param {string} name - 条目名。
   * @returns {string} 文本内容。
   */
  readText(name) {
    const buf = this.read(name)
    const text = buf.toString('utf8')
    return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text
  }

  /**
   * 标记条目为已修改内容。
   * @param {string} name - 条目名。
   * @param {Buffer|string} content - 新内容。
   * @returns {void}
   */
  write(name, content) {
    const buf = Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf8')
    if (buf.length > this.#limits.maxEntryBytes) {
      throw new OfficeError('MEMORY_LIMIT', `写入条目 ${name} 超过单条大小上限。`, { entry: name })
    }
    const previous = this.#entries.get(name)
    this.#entries.set(name, {
      name,
      method: previous?.method === 0 ? 0 : 8,
      flags: 0,
      crc: zlib.crc32(buf),
      uncompSize: buf.length,
      modTime: previous?.modTime ?? 0,
      modDate: previous?.modDate ?? 0x21, // 1980-01-01
      encrypted: false,
      replaced: true,
      data: buf
    })
    this.#deleted.delete(name)
    this.#cache.set(name, buf)
  }

  /**
   * 删除条目。
   * @param {string} name - 条目名。
   * @returns {boolean} 是否删除成功。
   */
  delete(name) {
    if (!this.#entries.has(name)) return false
    this.#entries.delete(name)
    this.#deleted.add(name)
    this.#cache.delete(name)
    return true
  }

  /**
   * 序列化为完整 ZIP 字节。未修改条目复用原压缩字节。
   * @returns {Buffer} ZIP 字节。
   */
  toBuffer() {
    const chunks = []
    const central = []
    let offset = 0

    for (const [name, entry] of this.#entries) {
      const nameBuf = Buffer.from(name, 'utf8')
      const needsUtf8 = nameBuf.length !== name.length
      let data
      let method = entry.method
      if (entry.replaced) {
        data = entry.data
        if (method === 8) {
          const deflated = zlib.deflateRawSync(data, { level: 6 })
          if (deflated.length < data.length) {
            data = deflated
          } else {
            method = 0
          }
        }
      } else {
        data = this.#rawOf(name)
      }
      const crc = entry.replaced ? entry.crc : entry.crc
      const uncompSize = entry.replaced ? entry.uncompSize : entry.uncompSize
      const flags = needsUtf8 ? 0x0800 : 0

      const local = Buffer.alloc(30)
      local.writeUInt32LE(SIG_LOCAL, 0)
      local.writeUInt16LE(20, 4)
      local.writeUInt16LE(flags, 6)
      local.writeUInt16LE(method, 8)
      local.writeUInt16LE(entry.modTime ?? 0, 10)
      local.writeUInt16LE(entry.modDate ?? 0x21, 12)
      local.writeUInt32LE(crc >>> 0, 14)
      local.writeUInt32LE(data.length, 18)
      local.writeUInt32LE(uncompSize, 22)
      local.writeUInt16LE(nameBuf.length, 26)
      local.writeUInt16LE(0, 28)
      chunks.push(local, nameBuf, data)

      const cen = Buffer.alloc(46)
      cen.writeUInt32LE(SIG_CENTRAL, 0)
      cen.writeUInt16LE(20, 4)
      cen.writeUInt16LE(20, 6)
      cen.writeUInt16LE(flags, 8)
      cen.writeUInt16LE(method, 10)
      cen.writeUInt16LE(entry.modTime ?? 0, 12)
      cen.writeUInt16LE(entry.modDate ?? 0x21, 14)
      cen.writeUInt32LE(crc >>> 0, 16)
      cen.writeUInt32LE(data.length, 20)
      cen.writeUInt32LE(uncompSize, 24)
      cen.writeUInt16LE(nameBuf.length, 28)
      cen.writeUInt16LE(0, 30)
      cen.writeUInt16LE(0, 32)
      cen.writeUInt16LE(0, 34)
      cen.writeUInt16LE(0, 36)
      cen.writeUInt32LE(0, 38)
      cen.writeUInt32LE(offset, 42)
      central.push(cen, nameBuf)

      offset += local.length + nameBuf.length + data.length
    }

    const centralBuf = Buffer.concat(central)
    const eocd = Buffer.alloc(22)
    eocd.writeUInt32LE(SIG_EOCD, 0)
    eocd.writeUInt16LE(0, 4)
    eocd.writeUInt16LE(0, 6)
    eocd.writeUInt16LE(this.#entries.size, 8)
    eocd.writeUInt16LE(this.#entries.size, 10)
    eocd.writeUInt32LE(centralBuf.length, 12)
    eocd.writeUInt32LE(offset, 16)
    eocd.writeUInt16LE(0, 20)

    return Buffer.concat([...chunks, centralBuf, eocd])
  }
}

/**
 * 支持字节级最小修改的 XML 文档。
 *
 * 解析阶段为每个元素记录源码偏移；编辑阶段只登记「区间补丁」；
 * 序列化时按偏移量拼接，未修改的字节逐字节原样保留。
 */
export class XmlDoc {
  #source
  #patches = []
  #limits
  #nodeCount = 0

  /**
   * @param {string} source - XML 文本。
   * @param {object} limits - 安全上限。
   */
  constructor(source, limits) {
    this.#source = source
    this.#limits = limits
  }

  /**
   * 解析 XML 文本。
   * @param {string} source - XML 文本。
   * @param {Partial<typeof DEFAULT_XML_LIMITS>} [limitOverrides] - 上限覆盖。
   * @returns {XmlDoc} 文档对象。
   */
  static parse(source, limitOverrides = {}) {
    const limits = { ...DEFAULT_XML_LIMITS, ...limitOverrides }
    if (source.length > limits.maxBytes) {
      throw new OfficeError('MEMORY_LIMIT', `XML 文本 ${source.length} 字节超过上限。`)
    }
    if (/<!DOCTYPE/i.test(source)) {
      throw new OfficeError('EXTERNAL_RESOURCE_BLOCKED', 'XML 含 DOCTYPE 声明，已按安全策略拒绝解析（防实体攻击）。')
    }
    const doc = new XmlDoc(source, limits)
    doc.root = doc.#parse()
    return doc
  }

  /**
   * 执行扫描解析。
   * @returns {object} 文档根节点。
   */
  #parse() {
    const src = this.#source
    const len = src.length
    const root = { type: 'root', name: '#document', children: [], start: 0, end: len, parent: null }
    const stack = [root]
    let i = 0
    let textStart = 0

    /** 把累积的文本作为一个文本节点挂到当前元素。 */
    const flushText = (until) => {
      if (until <= textStart) return
      const parent = stack.at(-1)
      if (parent === root) return
      const raw = src.slice(textStart, until)
      if (raw.trim() === '' && parent.children.length > 0) return
      parent.children.push({ type: 'text', name: '#text', raw, start: textStart, end: until, parent })
    }

    while (i < len) {
      const lt = src.indexOf('<', i)
      if (lt === -1) break
      if (lt > i) {
        flushText(lt)
        i = lt
      }
      if (src.startsWith('<!--', i)) {
        const end = src.indexOf('-->', i + 4)
        if (end === -1) throw new OfficeError('CORRUPTED_DOCUMENT', 'XML 注释未闭合。')
        i = end + 3
        textStart = i
        continue
      }
      if (src.startsWith('<![CDATA[', i)) {
        const end = src.indexOf(']]>', i + 9)
        if (end === -1) throw new OfficeError('CORRUPTED_DOCUMENT', 'XML CDATA 未闭合。')
        const parent = stack.at(-1)
        parent.children.push({ type: 'cdata', name: '#cdata', raw: src.slice(i + 9, end), start: i, end: end + 3, parent })
        i = end + 3
        textStart = i
        continue
      }
      if (src.startsWith('<?', i)) {
        const end = src.indexOf('?>', i + 2)
        if (end === -1) throw new OfficeError('CORRUPTED_DOCUMENT', 'XML 处理指令未闭合。')
        const parent = stack.at(-1)
        parent.children.push({ type: 'pi', name: '#pi', raw: src.slice(i, end + 2), start: i, end: end + 2, parent })
        i = end + 2
        textStart = i
        continue
      }
      if (src.startsWith('</', i)) {
        const end = src.indexOf('>', i + 2)
        if (end === -1) throw new OfficeError('CORRUPTED_DOCUMENT', 'XML 结束标签未闭合。')
        const name = src.slice(i + 2, end).trim()
        const open = stack.pop()
        if (!open || open === root) throw new OfficeError('CORRUPTED_DOCUMENT', `XML 出现多余的结束标签 </${name}>。`)
        if (open.name !== name) {
          throw new OfficeError('CORRUPTED_DOCUMENT', `XML 标签不匹配：<${open.name}> 与 </${name}>。`)
        }
        open.end = end + 1
        open.contentEnd = i
        i = end + 1
        textStart = i
        continue
      }
      if (src.startsWith('<!', i)) {
        throw new OfficeError('EXTERNAL_RESOURCE_BLOCKED', 'XML 含未支持的声明（DTD/实体），已拒绝解析。')
      }

      // 起始标签
      const parsed = this.#parseStartTag(i)
      this.#nodeCount += 1
      if (this.#nodeCount > this.#limits.maxNodes) {
        throw new OfficeError('MEMORY_LIMIT', `XML 节点数超过上限 ${this.#limits.maxNodes}。`)
      }
      if (stack.length > this.#limits.maxDepth) {
        throw new OfficeError('MEMORY_LIMIT', `XML 嵌套深度超过上限 ${this.#limits.maxDepth}。`)
      }
      const parent = stack.at(-1)
      parsed.parent = parent
      parsed.children = []
      parsed.attrs = parsed.attrs ?? new Map()
      parent.children.push(parsed)
      if (!parsed.selfClosing) stack.push(parsed)
      i = parsed.startTagEnd
      textStart = i
    }
    if (stack.length !== 1) {
      throw new OfficeError('CORRUPTED_DOCUMENT', `XML 有未闭合的标签 <${stack.at(-1)?.name}>。`)
    }
    return root
  }

  /**
   * 解析一个起始标签，记录属性值偏移量以便定点改写。
   * @param {number} start - `<` 的偏移。
   * @returns {object} 元素节点。
   */
  #parseStartTag(start) {
    const src = this.#source
    const len = src.length
    let i = start + 1
    const nameStart = i
    while (i < len && !/[\s/>]/.test(src[i])) i += 1
    const name = src.slice(nameStart, i)
    if (name === '') throw new OfficeError('CORRUPTED_DOCUMENT', `偏移 ${start} 处的标签名为空。`)
    const attrs = new Map()
    const attrOrder = []
    let selfClosing = false

    for (;;) {
      while (i < len && /\s/.test(src[i])) i += 1
      if (i >= len) throw new OfficeError('CORRUPTED_DOCUMENT', `标签 <${name}> 未闭合。`)
      if (src[i] === '>') { i += 1; break }
      if (src[i] === '/') {
        if (src[i + 1] !== '>') throw new OfficeError('CORRUPTED_DOCUMENT', `标签 <${name}> 的 '/' 后缺少 '>'。`)
        selfClosing = true
        i += 2
        break
      }
      const aStart = i
      while (i < len && !/[\s=/>]/.test(src[i])) i += 1
      const attrName = src.slice(aStart, i)
      while (i < len && /\s/.test(src[i])) i += 1
      if (src[i] !== '=') throw new OfficeError('CORRUPTED_DOCUMENT', `属性 ${attrName} 缺少 '='。`)
      i += 1
      while (i < len && /\s/.test(src[i])) i += 1
      const quote = src[i]
      if (quote !== '"' && quote !== "'") throw new OfficeError('CORRUPTED_DOCUMENT', `属性 ${attrName} 的值未加引号。`)
      i += 1
      const vStart = i
      const vEnd = src.indexOf(quote, i)
      if (vEnd === -1) throw new OfficeError('CORRUPTED_DOCUMENT', `属性 ${attrName} 的值未闭合。`)
      attrs.set(attrName, { raw: src.slice(vStart, vEnd), value: decodeEntities(src.slice(vStart, vEnd)), valueStart: vStart, valueEnd: vEnd, quote })
      attrOrder.push(attrName)
      i = vEnd + 1
    }
    return { type: 'element', name, start, startTagEnd: i, end: i, selfClosing, attrs, attrOrder, contentEnd: i }
  }

  /**
   * 登记一个区间补丁。区间不得与既有补丁重叠。
   *
   * 补丁按 `start` 保持有序，并用二分定位插入点，只与相邻补丁比较：
   * 有序区间集合中任何重叠必然发生在相邻项之间。若改成「每个补丁全扫已有补丁」，
   * 批量写入会退化成 O(n²)（实测 20 万单元格写入需 29 秒）。
   *
   * @param {number} start - 起始偏移（含）。
   * @param {number} end - 结束偏移（不含）。
   * @param {string} replacement - 替换文本。
   * @returns {object} 补丁对象，可在后续调用中就地改写其 `replacement`。
   */
  patch(start, end, replacement) {
    const patches = this.#patches
    let lo = 0
    let hi = patches.length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (patches[mid].start <= start) lo = mid + 1
      else hi = mid
    }
    const prev = patches[lo - 1]
    if (prev && start < prev.end && end > prev.start) {
      throw new OfficeError('INTERNAL_ERROR', `XML 补丁区间重叠：[${start},${end}) 与 [${prev.start},${prev.end})。`)
    }
    const next = patches[lo]
    if (next && next.start < end && start < next.end) {
      throw new OfficeError('INTERNAL_ERROR', `XML 补丁区间重叠：[${start},${end}) 与 [${next.start},${next.end})。`)
    }
    const entry = { start, end, replacement }
    patches.splice(lo, 0, entry)
    return entry
  }

  /**
   * 设置元素属性值；不存在则在起始标签内追加。
   * @param {object} node - 元素节点。
   * @param {string} name - 属性名。
   * @param {string} value - 属性值。
   * @returns {void}
   */
  /**
   * 设置元素属性值；不存在则在起始标签内追加。
   *
   * 每个属性只保留一个补丁并在重复赋值时就地改写，因此「同一属性写多次」
   * 不会产生区间重叠。属性原本不存在时，补丁是零长的插入补丁。
   *
   * @param {object} node - 元素节点。
   * @param {string} name - 属性名。
   * @param {string} value - 属性值。
   * @returns {void}
   */
  setAttr(node, name, value) {
    const existing = node.attrs.get(name)
    const escaped = escapeXmlAttr(value)
    if (existing) {
      if (existing.patch) {
        existing.patch.replacement = existing.inserted === true ? ` ${name}="${escaped}"` : escaped
      } else {
        existing.patch = this.patch(existing.valueStart, existing.valueEnd, escaped)
      }
      existing.value = String(value)
      existing.raw = escaped
      return
    }
    const insertAt = node.selfClosing ? node.startTagEnd - 2 : node.startTagEnd - 1
    const patch = this.patch(insertAt, insertAt, ` ${name}="${escaped}"`)
    node.attrs.set(name, {
      value: String(value),
      raw: escaped,
      valueStart: insertAt,
      valueEnd: insertAt,
      quote: '"',
      inserted: true,
      patch
    })
    node.attrOrder.push(name)
  }

  /**
   * 删除元素属性。
   * @param {object} node - 元素节点。
   * @param {string} name - 属性名。
   * @returns {boolean} 是否删除。
   */
  removeAttr(node, name) {
    const existing = node.attrs.get(name)
    if (!existing) return false
    if (existing.patch && existing.inserted === true) {
      // 属性是本会话新增的：直接撤销那个插入补丁。
      this.#patches = this.#patches.filter((p) => p !== existing.patch)
      node.attrs.delete(name)
      return true
    }
    const src = this.#source
    let from = existing.valueStart
    while (from > node.start && /\s/.test(src[from - 1])) from -= 1
    const to = existing.valueEnd + 1
    this.patch(from, to, '')
    node.attrs.delete(name)
    return true
  }

  /**
   * 读取元素全部文本内容（含子元素文本），已解码实体。
   * @param {object} node - 元素节点。
   * @returns {string} 文本内容。
   */
  text(node) {
    if (node.selfClosing || node.contentEnd === undefined || node.contentEnd < node.startTagEnd) return ''
    const raw = this.#source.slice(node.startTagEnd, node.contentEnd)
    if (!raw.includes('<')) return decodeEntities(raw)
    return decodeEntities(raw.replace(/<[^>]*>/g, ''))
  }

  /**
   * 替换元素的全部内容为纯文本。
   * @param {object} node - 元素节点。
   * @param {string} text - 新文本。
   * @returns {void}
   */
  setText(node, text) {
    if (node.selfClosing) {
      throw new OfficeError('INTERNAL_ERROR', `不能给自闭合元素 <${node.name}/> 设置文本。`)
    }
    this.patch(node.startTagEnd, node.contentEnd, text === '' ? '' : escapeXmlText(text))
  }

  /**
   * 用一段 XML 整体替换节点。
   * @param {object} node - 元素节点。
   * @param {string} xml - 替换用的 XML 文本。
   * @returns {void}
   */
  replace(node, xml) {
    this.patch(node.start, node.end, xml)
  }

  /**
   * 删除节点。
   * @param {object} node - 元素节点。
   * @returns {void}
   */
  remove(node) {
    this.patch(node.start, node.end, '')
  }

  /**
   * 在节点之后插入 XML 文本。
   * @param {object} node - 参照节点。
   * @param {string} xml - 待插入 XML。
   * @returns {void}
   */
  insertAfter(node, xml) {
    this.patch(node.end, node.end, xml)
  }

  /**
   * 在节点之前插入 XML 文本。
   * @param {object} node - 参照节点。
   * @param {string} xml - 待插入 XML。
   * @returns {void}
   */
  insertBefore(node, xml) {
    this.patch(node.start, node.start, xml)
  }

  /**
   * 在父元素内容末尾追加 XML 文本。自闭合父元素会被就地展开。
   * @param {object} parent - 父元素。
   * @param {string} xml - 待插入 XML。
   * @returns {void}
   */
  appendChild(parent, xml) {
    if (parent.selfClosing) {
      // 把 `<x/>` 就地展开成 `<x>…</x>`：先去掉 `/`，再补子节点与结束标签。
      this.patch(parent.startTagEnd - 2, parent.startTagEnd - 1, '>')
      this.patch(parent.end, parent.end, `${xml}</${parent.name}>`)
      parent.selfClosing = false
      parent.contentEnd = parent.end
      return
    }
    this.patch(parent.contentEnd, parent.contentEnd, xml)
  }

  /**
   * 用当前文本重新解析，使本次会话插入的节点也能被后续查询看到。
   *
   * 补丁模型下新增节点只存在于输出文本里，不在已解析的树上；
   * 需要「看到自己刚插入的内容」时（如重复性检查）先调用本方法。
   * 调用后此前持有的节点引用全部失效。
   *
   * @returns {object} 新的根节点。
   */
  rescan() {
    const text = this.toString()
    const fresh = XmlDoc.parse(text, this.#limits)
    this.#source = text
    this.#patches = []
    this.root = fresh.root
    return this.root
  }

  /**
   * 取节点在**原始源码**中的字节片段（不含本会话补丁）。
   * 用于与既有内容做语义比对，例如判断某个 `<font>` 是否已存在。
   *
   * @param {object} node - 元素节点。
   * @returns {string} 原始 XML 片段。
   */
  raw(node) {
    return this.#source.slice(node.start, node.end)
  }

  /** @returns {string} 应用全部补丁后的 XML 文本。 */
  toString() {
    if (this.#patches.length === 0) return this.#source
    // 补丁在 patch() 中已按 start 保持有序，这里直接顺序拼接。
    let out = ''
    let cursor = 0
    for (const p of this.#patches) {
      out += this.#source.slice(cursor, p.start) + p.replacement
      cursor = p.end
    }
    return out + this.#source.slice(cursor)
  }
}

/**
 * 递归收集一个元素的全部文本内容（已解码实体）。
 *
 * 供各格式适配器共用；WordprocessingML / PresentationML 与 SpreadsheetML
 * 的取值需求一致，因此不重复实现。
 *
 * @param {XmlDoc} doc - 文档对象。
 * @param {object} node - 元素节点。
 * @returns {string} 文本内容。
 */
export function nodeText(doc, node) {
  let text = ''
  for (const child of node.children ?? []) {
    if (child.type === 'text') text += decodeEntities(child.raw)
    else if (child.type === 'cdata') text += child.raw
    else if (child.type === 'element') text += nodeText(doc, child)
  }
  return text
}

/** 每像素对应的 EMU 数（96 DPI 下 1 px = 9525 EMU）。 */
export const EMU_PER_PX = 9525

/** 支持的图片类型及其 MIME。 */
export const IMAGE_TYPES = Object.freeze({
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf'
})

/**
 * 取图片扩展名对应的 MIME 类型。
 * @param {string} extension - 扩展名。
 * @returns {string} MIME。
 */
export function imageContentType(extension) {
  const mime = IMAGE_TYPES[extension.toLowerCase()]
  if (!mime) throw new OfficeError('UNSUPPORTED_FILE_TYPE', `不支持的图片扩展名：${extension}`)
  return mime
}

/**
 * 按真实字节判定图片类型。
 *
 * 与文件类型检测同一原则：不信扩展名。只看魔数，
 * 因此把 .png 改名成 .jpg 也不会写出错误的内容类型声明。
 *
 * @param {Buffer} data - 图片字节。
 * @returns {string} 扩展名（png/jpg/gif/bmp）。
 */
export function detectImageType(data) {
  if (data.length >= 8 && data.readUInt32BE(0) === 0x89504e47) return 'png'
  if (data.length >= 3 && data[0] === 0xff && data[1] === 0xd8 && data[2] === 0xff) return 'jpg'
  if (data.length >= 6 && data.subarray(0, 3).toString('latin1') === 'GIF') return 'gif'
  if (data.length >= 2 && data.subarray(0, 2).toString('latin1') === 'BM') return 'bmp'
  throw new OfficeError('UNSUPPORTED_FILE_TYPE', '无法识别的图片格式（支持 PNG / JPEG / GIF / BMP）。', {
    solution: '请先转换为 PNG 或 JPEG 再插入。'
  })
}

/**
 * 读取图片的像素尺寸。
 *
 * 只解析文件头，不解码像素，因此对任意大小的图片都是常数开销。
 * 无法判定时返回 null，由调用方回退到默认尺寸。
 *
 * @param {Buffer} data - 图片字节。
 * @param {string} extension - 扩展名。
 * @returns {{width: number, height: number}|null} 像素尺寸。
 */
export function readImageSize(data, extension) {
  try {
    const ext = extension.toLowerCase()
    if (ext === 'png' && data.length >= 24) {
      return { width: data.readUInt32BE(16), height: data.readUInt32BE(20) }
    }
    if (ext === 'gif' && data.length >= 10) {
      return { width: data.readUInt16LE(6), height: data.readUInt16LE(8) }
    }
    if (ext === 'bmp' && data.length >= 26) {
      return { width: Math.abs(data.readInt32LE(18)), height: Math.abs(data.readInt32LE(22)) }
    }
    if ((ext === 'jpg' || ext === 'jpeg') && data.length >= 4) {
      let offset = 2
      while (offset + 9 < data.length) {
        if (data[offset] !== 0xff) {
          offset += 1
          continue
        }
        const marker = data[offset + 1]
        // SOF0–SOF15（除 DHT/JPG/DAC）携带帧尺寸
        if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
          return { height: data.readUInt16BE(offset + 5), width: data.readUInt16BE(offset + 7) }
        }
        const length = data.readUInt16BE(offset + 2)
        if (length < 2) return null
        offset += 2 + length
      }
    }
  } catch {
    return null
  }
  return null
}

/**
 * 计算插入图片的显示尺寸（EMU）。
 *
 * 只给宽度时按原始比例缩放；都不给时用原始像素尺寸（96 DPI）；
 * 原始尺寸未知且未指定时回退到 4 英寸宽。DOCX 与 PPTX 适配器共用。
 *
 * @param {object} args - `{natural, widthPx, heightPx}`。
 * @returns {{cx: number, cy: number}} EMU 尺寸。
 */
export function resolveImageExtent({ natural, widthPx, heightPx }) {
  let w = widthPx
  let h = heightPx
  if (w === undefined && h === undefined) {
    if (natural) {
      w = natural.width
      h = natural.height
    } else {
      w = 384
      h = 216
    }
  } else if (w !== undefined && h === undefined) {
    h = natural && natural.width > 0 ? Math.round((w * natural.height) / natural.width) : w
  } else if (w === undefined && h !== undefined) {
    w = natural && natural.height > 0 ? Math.round((h * natural.width) / natural.height) : h
  }
  if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) {
    throw new OfficeError('INVALID_REQUEST', `图片显示尺寸非法：${w}×${h}`)
  }
  return { cx: Math.round(w * EMU_PER_PX), cy: Math.round(h * EMU_PER_PX) }
}

/**
 * 确保 `[Content_Types].xml` 声明了某个扩展名的 Default。
 *
 * OPC 层的通用能力，三种格式适配器共用。这类操作如果在每个适配器里各写一份，
 * 「同一件事有两种实现」迟早分叉 —— PPTX 适配器就踩过一次「读写各写一套遍历」
 * 导致下标错位的坑，同样的道理适用于容器层操作。
 *
 * @param {ZipPackage} pkg - OPC 包。
 * @param {string} extension - 扩展名（不带点）。
 * @param {string} contentType - MIME 类型。
 * @returns {void}
 */
export function ensureContentTypeDefault(pkg, extension, contentType) {
  const doc = XmlDoc.parse(pkg.readText('[Content_Types].xml'))
  const root = doc.root.children.find((c) => c.type === 'element')
  for (const node of findAll(root, 'Default')) {
    if ((attr(node, 'Extension') ?? '').toLowerCase() === extension.toLowerCase()) return
  }
  doc.appendChild(root, `<Default Extension="${escapeXmlAttr(extension)}" ContentType="${contentType}"/>`)
  pkg.write('[Content_Types].xml', doc.toString())
}

/**
 * 确保 `[Content_Types].xml` 声明了某个部件的 Override。
 * @param {ZipPackage} pkg - OPC 包。
 * @param {string} partName - 以 `/` 开头的部件名。
 * @param {string} contentType - MIME 类型。
 * @returns {void}
 */
export function ensureContentTypeOverride(pkg, partName, contentType) {
  const doc = XmlDoc.parse(pkg.readText('[Content_Types].xml'))
  const root = doc.root.children.find((c) => c.type === 'element')
  for (const node of findAll(root, 'Override')) {
    if (attr(node, 'PartName') === partName) return
  }
  doc.appendChild(root, `<Override PartName="${escapeXmlAttr(partName)}" ContentType="${contentType}"/>`)
  pkg.write('[Content_Types].xml', doc.toString())
}

/**
 * 移除 `[Content_Types].xml` 中某个部件的 Override。
 * @param {ZipPackage} pkg - OPC 包。
 * @param {string} partName - 以 `/` 开头的部件名。
 * @returns {void}
 */
export function removeContentTypeOverride(pkg, partName) {
  const doc = XmlDoc.parse(pkg.readText('[Content_Types].xml'))
  for (const node of findAll(doc.root, 'Override')) {
    if (attr(node, 'PartName') === partName) doc.remove(node)
  }
  pkg.write('[Content_Types].xml', doc.toString())
}

/**
 * 向任意关系部件追加一条关系，返回新的关系 ID。
 *
 * ID 取既有最大值 +1，而不是从 1 开始试 —— 猜一个已被占用的 ID
 * 会让文档里的引用指向错误的部件。
 *
 * @param {ZipPackage} pkg - OPC 包。
 * @param {string} relsPart - 关系部件路径，如 `word/_rels/document.xml.rels`。
 * @param {string} type - 关系类型 URI。
 * @param {string} target - 关系目标（相对该部件所在目录）。
 * @returns {string} 新的关系 ID。
 */
export function addRelationshipTo(pkg, relsPart, type, target) {
  const xml = pkg.has(relsPart)
    ? pkg.readText(relsPart)
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`
  const doc = XmlDoc.parse(xml)
  const root = doc.root.children.find((c) => c.type === 'element')
  let max = 0
  for (const rel of findAll(root, 'Relationship')) {
    const match = /^rId(\d+)$/.exec(attr(rel, 'Id') ?? '')
    if (match) max = Math.max(max, Number(match[1]))
  }
  const rid = `rId${max + 1}`
  doc.appendChild(root, `<Relationship Id="${rid}" Type="${type}" Target="${escapeXmlAttr(target)}"/>`)
  pkg.write(relsPart, doc.toString())
  return rid
}

/**
 * 从关系部件中删除一条关系。
 * @param {ZipPackage} pkg - OPC 包。
 * @param {string} relsPart - 关系部件路径。
 * @param {string} rid - 关系 ID。
 * @returns {void}
 */
export function removeRelationshipFrom(pkg, relsPart, rid) {
  if (!pkg.has(relsPart)) return
  const doc = XmlDoc.parse(pkg.readText(relsPart))
  for (const rel of findAll(doc.root, 'Relationship')) {
    if (attr(rel, 'Id') === rid) doc.remove(rel)
  }
  pkg.write(relsPart, doc.toString())
}

/**
 * 深度优先查找全部匹配元素（按本地名匹配，忽略命名空间前缀）。
 * @param {object} node - 起始节点。
 * @param {string} localName - 目标本地名。
 * @param {object[]} [out] - 结果累积数组。
 * @returns {object[]} 匹配到的元素节点。
 */
export function findAll(node, localName, out = []) {
  for (const child of node.children ?? []) {
    if (child.type !== 'element') continue
    const local = child.name.includes(':') ? child.name.slice(child.name.indexOf(':') + 1) : child.name
    if (local === localName) out.push(child)
    findAll(child, localName, out)
  }
  return out
}

/**
 * 查找第一个匹配元素。
 * @param {object} node - 起始节点。
 * @param {string} localName - 目标本地名。
 * @returns {object|undefined} 匹配到的元素节点。
 */
export function find(node, localName) {
  return findAll(node, localName, [])[0]
}

/**
 * 读取元素属性值（忽略命名空间前缀）。
 * @param {object} node - 元素节点。
 * @param {string} localName - 属性本地名。
 * @returns {string|undefined} 属性值。
 */
export function attr(node, localName) {
  const direct = node.attrs?.get(localName)
  if (direct) return direct.value
  for (const [key, entry] of node.attrs ?? []) {
    const local = key.includes(':') ? key.slice(key.indexOf(':') + 1) : key
    if (local === localName) return entry.value
  }
  return undefined
}
