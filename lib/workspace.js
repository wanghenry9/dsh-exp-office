/**
 * 文件管理层：路径安全、真实类型检测、哈希、临时目录、文件锁与事务式写入。
 *
 * 对齐开发要求 §二.4（文件管理层）、§十二.1（文件安全）与 §十一（事务式流程）：
 *   1. 不直接覆盖原文件      2. 创建临时副本     3. 在副本上修改
 *   4. 保存后重新读取验证    5. 验证成功才产出   6. 验证失败保留原文件
 *   7. 清理无效输出与临时资源 8. 记录变更与性能
 *
 * @module dsh-exp-office/workspace
 */

import { createHash, randomUUID } from 'node:crypto'
import { open, mkdir, readFile, rename, rm, stat, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { dirname, isAbsolute, join, normalize, resolve, sep } from 'node:path'
import { OfficeError } from './errors.js'

/** 允许处理的文件类型及其真实结构特征（docs §十二.1：不能只依赖扩展名）。 */
export const SUPPORTED_TYPES = Object.freeze({
  xlsx: { ext: 'xlsx', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', kind: 'zip' },
  xlsm: { ext: 'xlsm', mime: 'application/vnd.ms-excel.sheet.macroEnabled.12', kind: 'zip', macro: true },
  xltx: { ext: 'xltx', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template', kind: 'zip' },
  docx: { ext: 'docx', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', kind: 'zip' },
  docm: { ext: 'docm', mime: 'application/vnd.ms-word.document.macroEnabled.12', kind: 'zip', macro: true },
  pptx: { ext: 'pptx', mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', kind: 'zip' },
  pptm: { ext: 'pptm', mime: 'application/vnd.ms-powerpoint.presentation.macroEnabled.12', kind: 'zip', macro: true },
  pdf: { ext: 'pdf', mime: 'application/pdf', kind: 'pdf' },
  csv: { ext: 'csv', mime: 'text/csv', kind: 'text' }
})

/** 单文件大小上限（docs §十二.1 限制文件大小）。 */
export const MAX_FILE_BYTES = 512 * 1024 * 1024

/** 文件句柄与锁的默认存活时间。 */
const LOCK_STALE_MS = 10 * 60 * 1000

/**
 * 校验并归一化用户提供的路径。
 *
 * 拒绝：空字节、UNC/设备路径、相对穿越到工作区之外。
 *
 * @param {string} input - 用户提供的路径。
 * @param {object} options - 选项。
 * @param {string} options.workspaceRoot - 允许访问的根目录。
 * @param {boolean} [options.mustExist] - 是否要求文件已存在。
 * @returns {string} 绝对路径。
 */
export function resolveSafePath(input, { workspaceRoot, mustExist = false }) {
  if (typeof input !== 'string' || input.trim() === '') {
    throw new OfficeError('INVALID_REQUEST', '路径不能为空。')
  }
  if (input.includes('\0')) {
    throw new OfficeError('INVALID_REQUEST', '路径包含空字节，已拒绝。')
  }
  if (/^\\\\[.?]\\/.test(input) || input.startsWith('\\\\?\\') || input.startsWith('\\\\.\\')) {
    throw new OfficeError('PERMISSION_DENIED', '不允许访问设备路径或 UNC 扩展路径。', { path: input })
  }
  const root = resolve(workspaceRoot)
  const absolute = isAbsolute(input) ? normalize(input) : resolve(root, input)
  const normalizedRoot = root.endsWith(sep) ? root : root + sep
  if (absolute !== root && !absolute.startsWith(normalizedRoot)) {
    throw new OfficeError('PERMISSION_DENIED', '路径超出允许访问的工作区范围。', { path: input, workspace_root: root })
  }
  if (mustExist && !existsSync(absolute)) {
    throw new OfficeError('FILE_NOT_FOUND', `文件不存在：${input}`, { path: absolute })
  }
  return absolute
}

/**
 * 依据真实字节结构识别文件类型，并与扩展名交叉校验。
 *
 * ZIP 容器进一步读取包内标志部件，以区分 xlsx / docx / pptx 与宏文件，
 * 因此把 .docx 改名成 .xlsx 也会被识破。
 *
 * @param {Buffer} buffer - 文件字节。
 * @param {string} filename - 文件名（用于扩展名交叉校验）。
 * @returns {object} `{ext, mime, kind, macro, mismatch}`。
 */
export function detectFileType(buffer, filename = '') {
  const declared = filename.includes('.') ? filename.slice(filename.lastIndexOf('.') + 1).toLowerCase() : ''
  const declaredSpec = SUPPORTED_TYPES[declared]

  if (buffer.length >= 5 && buffer.subarray(0, 5).toString('latin1') === '%PDF-') {
    return finalize('pdf', declared)
  }
  if (buffer.length >= 8 && buffer.readUInt32LE(0) === 0xe011cfd0) {
    // OLE 复合文档：加密的 OOXML，或 .xls/.doc/.ppt 旧版二进制格式
    const ole = declaredSpec && ['xlsx', 'docx', 'pptx'].includes(declared) ? declared : null
    if (ole) {
      throw new OfficeError('PASSWORD_REQUIRED', `文件是加密的 OOXML（OLE 容器），需要密码或解密副本。`, { filename })
    }
    throw new OfficeError('UNSUPPORTED_FILE_TYPE', '检测到 OLE 复合文档（旧版 Office 二进制格式或加密文件），当前不支持。', { filename })
  }
  if (buffer.length >= 4 && buffer.readUInt32LE(0) === 0x04034b50) {
    const kind = sniffOoxml(buffer)
    if (!kind) {
      throw new OfficeError('UNSUPPORTED_FILE_TYPE', 'ZIP 容器内没有 Office 标志部件，不是 Office 文档。', { filename })
    }
    return finalize(kind, declared)
  }
  if (buffer.length === 0) {
    throw new OfficeError('CORRUPTED_DOCUMENT', '文件为空。', { filename })
  }
  throw new OfficeError('UNSUPPORTED_FILE_TYPE', `无法识别的文件结构（扩展名 .${declared || '无'}）。`, { filename })
}

/**
 * 组装类型检测结果并标记扩展名不一致。
 *
 * `mismatch` 必须比较「按字节识别出的类型」与「扩展名声明的类型」。
 * 早先调用方传进来的是**已识别类型**的规格，于是两者永远相等、
 * `mismatch` 恒为 false —— 把一个 .docx 改名成 .xlsx 也检测不出来。
 * 因此这里只接收扩展名，由本函数自己去查声明规格。
 *
 * @param {string} ext - 按字节识别出的真实扩展名。
 * @param {string} declared - 文件名声明的扩展名。
 * @returns {object} 检测结果。
 */
function finalize(ext, declared) {
  const spec = SUPPORTED_TYPES[ext]
  const declaredSpec = SUPPORTED_TYPES[declared]
  const mismatch = declaredSpec !== undefined && declaredSpec.ext !== ext
  return { ext, mime: spec.mime, kind: spec.kind, macro: spec.macro === true, mismatch, declared_ext: declared }
}

/**
 * 通过包内标志部件判断 OOXML 具体类型。
 * @param {Buffer} buffer - ZIP 字节。
 * @returns {string|null} 识别出的类型键。
 */
function sniffOoxml(buffer) {
  const text = buffer.toString('latin1')
  const hasMacro = text.includes('vbaProject.bin')
  if (text.includes('xl/workbook.xml')) return hasMacro ? 'xlsm' : 'xlsx'
  if (text.includes('word/document.xml')) return hasMacro ? 'docm' : 'docx'
  if (text.includes('ppt/presentation.xml')) return hasMacro ? 'pptm' : 'pptx'
  return null
}

/**
 * 计算字节的 SHA-256。
 * @param {Buffer|string} data - 数据。
 * @returns {string} 十六进制摘要。
 */
export function sha256(data) {
  return createHash('sha256').update(data).digest('hex')
}

/** 为一次任务分配的临时目录，任务结束必须清理。 */
export class TempWorkspace {
  #dir
  #cleaned = false

  /**
   * @param {string} dir - 临时目录绝对路径。
   */
  constructor(dir) {
    this.#dir = dir
  }

  /**
   * 创建一个随机命名的临时目录（docs §十二.1：临时文件使用随机目录）。
   * @param {string} baseDir - 父目录。
   * @param {string} [tag] - 可选标签，仅用于排查。
   * @returns {Promise<TempWorkspace>} 临时目录对象。
   */
  static async create(baseDir, tag = 'task') {
    const dir = join(baseDir, 'tmp', `${tag}-${randomUUID()}`)
    await mkdir(dir, { recursive: true })
    return new TempWorkspace(dir)
  }

  /** @returns {string} 临时目录路径。 */
  get path() {
    return this.#dir
  }

  /**
   * 在临时目录内分配一个文件路径。
   * @param {string} name - 文件名。
   * @returns {string} 绝对路径。
   */
  file(name) {
    return join(this.#dir, name)
  }

  /**
   * 删除临时目录及其全部内容。可重复调用。
   * @returns {Promise<void>} 完成信号。
   */
  async dispose() {
    if (this.#cleaned) return
    this.#cleaned = true
    await rm(this.#dir, { recursive: true, force: true }).catch(() => {})
  }
}

/** 基于独占文件句柄的跨进程文件锁（docs §十二.4：避免多个任务同时修改同一文件）。 */
export class FileLock {
  #handle
  #lockPath

  /**
   * @param {string} lockPath - 锁文件路径。
   * @param {object} handle - 已独占打开的句柄。
   */
  constructor(lockPath, handle) {
    this.#lockPath = lockPath
    this.#handle = handle
  }

  /**
   * 尝试获取目标文件的锁。
   * @param {string} targetPath - 目标文件绝对路径。
   * @param {object} [options] - 选项。
   * @param {number} [options.timeoutMs] - 等待上限。
   * @returns {Promise<FileLock>} 锁对象。
   */
  static async acquire(targetPath, { timeoutMs = 5000 } = {}) {
    const lockPath = `${targetPath}.dsh-office.lock`
    const deadline = Date.now() + timeoutMs
    for (;;) {
      try {
        const handle = await open(lockPath, 'wx')
        await handle.writeFile(JSON.stringify({ pid: process.pid, at: new Date().toISOString() }))
        return new FileLock(lockPath, handle)
      } catch (err) {
        if (err.code !== 'EEXIST') throw err
        // 残留锁（进程异常退出）超过存活时间即视为失效
        const info = await stat(lockPath).catch(() => null)
        if (info && Date.now() - info.mtimeMs > LOCK_STALE_MS) {
          await rm(lockPath, { force: true }).catch(() => {})
          continue
        }
        if (Date.now() >= deadline) {
          throw new OfficeError('FILE_LOCKED', '文件正被其它任务修改，请稍后重试。', { path: targetPath })
        }
        await new Promise((r) => setTimeout(r, 100))
      }
    }
  }

  /**
   * 释放锁。异常路径也必须调用（docs §十八.3：异常时必须释放资源）。
   * @returns {Promise<void>} 完成信号。
   */
  async release() {
    if (!this.#handle) return
    const handle = this.#handle
    this.#handle = null
    await handle.close().catch(() => {})
    await rm(this.#lockPath, { force: true }).catch(() => {})
  }
}

/**
 * 事务式文件写入。
 *
 * 流程：写临时文件 → 可选校验 → 原子替换目标。校验失败时目标文件保持原样，
 * 并清理临时产物。
 */
export class Transaction {
  #target
  #tempPath
  #warnings = []
  #committed = false

  /**
   * @param {string} target - 目标文件绝对路径。
   * @param {string} tempPath - 临时文件路径。
   */
  constructor(target, tempPath) {
    this.#target = target
    this.#tempPath = tempPath
  }

  /**
   * 开启事务。
   * @param {string} target - 目标文件绝对路径。
   * @param {string} tempDir - 临时目录。
   * @returns {Promise<Transaction>} 事务对象。
   */
  static async begin(target, tempDir) {
    const tempPath = join(tempDir, `staging-${randomUUID()}`)
    return new Transaction(target, tempPath)
  }

  /** @returns {string} 暂存文件路径，供适配器写入。 */
  get stagingPath() {
    return this.#tempPath
  }

  /**
   * 登记一条警告。
   * @param {object} warning - `{code, message}`。
   * @returns {void}
   */
  warn(warning) {
    this.#warnings.push(warning)
  }

  /** @returns {object[]} 已登记警告。 */
  get warnings() {
    return this.#warnings
  }

  /**
   * 写入暂存文件并执行可选校验，全部通过后原子替换目标。
   *
   * @param {Buffer} bytes - 输出字节。
   * @param {object} [options] - 选项。
   * @param {(bytes: Buffer) => void|Promise<void>} [options.verify] - 校验回调，抛错即回滚。
   * @param {boolean} [options.overwrite] - 是否允许覆盖已存在的目标文件。
   * @returns {Promise<object>} `{path, size, sha256, warnings}`。
   */
  async commit(bytes, { verify = null, overwrite = true } = {}) {
    if (!existsSync(dirname(this.#tempPath))) {
      await mkdir(dirname(this.#tempPath), { recursive: true })
    }
    await writeFile(this.#tempPath, bytes)
    try {
      if (verify) await verify(bytes)
    } catch (err) {
      await rm(this.#tempPath, { force: true }).catch(() => {})
      throw err instanceof OfficeError
        ? err
        : new OfficeError('OUTPUT_VALIDATION_FAILED', `输出校验失败，原文件已保留：${err.message}`, { needsRollback: true })
    }
    if (!overwrite && existsSync(this.#target)) {
      await rm(this.#tempPath, { force: true }).catch(() => {})
      throw new OfficeError('PERMISSION_DENIED', '目标文件已存在且未允许覆盖，已保留原文件。', { path: this.#target })
    }
    await rename(this.#tempPath, this.#target)
    this.#committed = true
    return { path: this.#target, size: bytes.length, sha256: sha256(bytes), warnings: this.#warnings }
  }

  /**
   * 放弃事务并清理暂存文件。失败路径必须调用。
   * @returns {Promise<void>} 完成信号。
   */
  async rollback() {
    if (this.#committed) return
    await rm(this.#tempPath, { force: true }).catch(() => {})
  }
}

/** 受管目录：输出、临时、备份、审计日志都收在插件自己的目录下。 */
export class ManagedStore {
  #root

  /**
   * @param {string} root - 受管根目录。
   */
  constructor(root) {
    this.#root = root
  }

  /**
   * 打开（并创建）受管目录。
   * @param {string} workspaceRoot - 会话工作区根目录。
   * @returns {Promise<ManagedStore>} 受管目录对象。
   */
  static async open(workspaceRoot) {
    const root = join(workspaceRoot, '.dsh-exp-office')
    await mkdir(join(root, 'tmp'), { recursive: true })
    await mkdir(join(root, 'out'), { recursive: true })
    return new ManagedStore(root)
  }

  /** @returns {string} 受管根目录。 */
  get root() {
    return this.#root
  }

  /**
   * 生成一个受管输出文件路径，避免覆盖既有文件。
   * @param {string} name - 期望的文件名。
   * @returns {string} 绝对路径。
   */
  outputPath(name) {
    const safe = name.replace(/[\\/:*?"<>|]/g, '_')
    return join(this.#root, 'out', safe)
  }

  /**
   * 为一个任务创建临时目录。
   * @param {string} tag - 任务标签。
   * @returns {Promise<TempWorkspace>} 临时目录。
   */
  async temp(tag) {
    return TempWorkspace.create(this.#root, tag)
  }

  /**
   * 读取一个受管文件。
   * @param {string} name - 文件名。
   * @returns {Promise<Buffer>} 文件字节。
   */
  async read(name) {
    const path = this.outputPath(name)
    if (!existsSync(path)) throw new OfficeError('FILE_NOT_FOUND', `受管目录中不存在文件 ${name}。`, { name })
    return readFile(path)
  }
}

/**
 * 读取工作区文件并做大小与类型校验。
 * @param {string} absolutePath - 绝对路径。
 * @param {object} [options] - 选项。
 * @param {number} [options.maxBytes] - 大小上限。
 * @returns {Promise<{buffer: Buffer, type: object, path: string}>} 读取结果。
 */
export async function readDocument(absolutePath, { maxBytes = MAX_FILE_BYTES } = {}) {
  const info = await stat(absolutePath).catch(() => null)
  if (!info) throw new OfficeError('FILE_NOT_FOUND', '文件不存在。', { path: absolutePath })
  if (!info.isFile()) throw new OfficeError('INVALID_REQUEST', '目标不是普通文件。', { path: absolutePath })
  if (info.size > maxBytes) {
    throw new OfficeError('MEMORY_LIMIT', `文件 ${info.size} 字节超过上限 ${maxBytes}。`, { size: info.size })
  }
  const buffer = await readFile(absolutePath)
  const type = detectFileType(buffer, absolutePath)
  return { buffer, type, path: absolutePath }
}
