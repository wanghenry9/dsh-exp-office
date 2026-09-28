/**
 * 统一操作审计：每个工具调用落**一行 JSONL**（谁在用、调了什么、耗时、成败、目标文件）。
 *
 * 设计取舍：
 *   - **默认关闭**（与本地联动同一原则：不写任何本地日志除非显式开启）；
 *   - **只记元数据，不记内容**：不写文档正文、不写单元格值、不写路径全称；
 *     路径默认只留**文件名**（`basename`），需要定位时可以切到 `relative`（相对工作区）；
 *   - **永不因为审计失败而让工具失败**：写盘异常只记在内存里，工具照常返回；
 *   - **不会无限增长**：超过 `maxBytes`（默认 8 MB）就轮转成一个 `.1` 备份文件。
 *
 * @module dsh-exp-office/audit
 */

import { appendFileSync, mkdirSync, renameSync, statSync } from 'node:fs'
import { basename, dirname, isAbsolute, relative, resolve } from 'node:path'

/** 默认单文件大小上限：8 MB。 */
export const DEFAULT_AUDIT_MAX_BYTES = 8 * 1024 * 1024

/** 允许的路径记录方式。 */
export const AUDIT_PATH_MODES = Object.freeze(['basename', 'relative'])

/**
 * 工具审计日志。
 */
export class ToolAuditLog {
  #file

  #mode

  #workspaceRoot

  #maxBytes

  #writeErrors = 0

  #lastError = null

  /**
   * @param {object} options - 选项。
   * @param {string} options.file - JSONL 文件绝对路径。
   * @param {string} [options.pathMode] - `basename`（默认）或 `relative`。
   * @param {string} options.workspaceRoot - 工作区根目录（relative 模式的基准）。
   * @param {number} [options.maxBytes] - 单文件上限，超过则轮转。
   */
  constructor({ file, pathMode = 'basename', workspaceRoot, maxBytes = DEFAULT_AUDIT_MAX_BYTES }) {
    this.#file = file
    this.#mode = AUDIT_PATH_MODES.includes(pathMode) ? pathMode : 'basename'
    this.#workspaceRoot = workspaceRoot
    this.#maxBytes = Number.isFinite(maxBytes) && maxBytes > 0 ? maxBytes : DEFAULT_AUDIT_MAX_BYTES
  }

  /** 日志文件路径。 @returns {string} 绝对路径。 */
  get file() {
    return this.#file
  }

  /** 写失败次数（工具不会因此失败，但调用方可以查）。 @returns {number} 次数。 */
  get writeErrors() {
    return this.#writeErrors
  }

  /** 最近一次写失败原因。 @returns {string|null} 原因。 */
  get lastError() {
    return this.#lastError
  }

  /**
   * 把一个路径按配置的方式记录成可写进日志的字符串。
   *
   * 非绝对路径原样记录（本来就不是本机路径）；绝对路径按模式取 `basename` 或工作区相对路径，
   * **工作区之外的绝对路径只留文件名** —— 审计日志不应该泄漏本机目录结构。
   *
   * @param {unknown} value - 原始路径。
   * @returns {string|null} 记录用的字符串。
   */
  describePath(value) {
    if (typeof value !== 'string' || value === '') return null
    if (!isAbsolute(value)) return value
    if (this.#mode === 'relative') {
      const rel = relative(resolve(this.#workspaceRoot), resolve(value))
      // 跳出工作区（以 .. 开头）时退回文件名，避免泄漏外部目录
      if (rel !== '' && !rel.startsWith('..') && !isAbsolute(rel)) return rel
    }
    return basename(value)
  }

  /**
   * 追加一条记录。**永不抛异常**。
   * @param {object} entry - 记录内容（会被 JSON 序列化并压成一行）。
   * @returns {boolean} 是否写入成功。
   */
  record(entry) {
    try {
      this.#rotateIfNeeded()
      mkdirSync(dirname(this.#file), { recursive: true })
      appendFileSync(this.#file, `${JSON.stringify(entry)}\n`, 'utf8')
      return true
    } catch (err) {
      this.#writeErrors += 1
      this.#lastError = err?.message ?? String(err)
      return false
    }
  }

  /**
   * 超过上限时把现有文件轮转成 `<file>.1`（覆盖旧备份）。
   *
   * 用 `statSync(..., { throwIfNoEntry: false })` 一次系统调用同时完成「存在性 + 大小」判断，
   * 所以可以每次写入都查 —— 行为可预期（一超限就轮转），代价是一次 stat。
   * @returns {void}
   */
  #rotateIfNeeded() {
    try {
      const stats = statSync(this.#file, { throwIfNoEntry: false })
      if (!stats || stats.size < this.#maxBytes) return
      renameSync(this.#file, `${this.#file}.1`)
    } catch {
      // 轮转失败不影响这次记录（下一次写还会再试）
    }
  }
}

/**
 * 从工具参数里挑出「这次调用动了哪个文件」。
 * @param {object} args - 工具参数。
 * @returns {unknown} 路径或空。
 */
function targetOf(args) {
  if (!args || typeof args !== 'object') return null
  for (const key of ['path', 'directory', 'file', 'target_path', 'output_path']) {
    const value = args[key]
    if (typeof value === 'string' && value !== '') return value
  }
  return null
}

/**
 * 给一个工具定义套上审计中间件。
 *
 * 包装后的 `execute` 与原实现行为完全一致（返回值、异常都不变），只在返回前记一行日志。
 *
 * @param {object} definition - 原始工具定义。
 * @param {ToolAuditLog} audit - 审计日志。
 * @param {object} [options] - 选项。
 * @param {string} [options.pluginVersion] - 写进日志的插件版本。
 * @returns {object} 包装后的定义。
 */
export function withToolAudit(definition, audit, { pluginVersion = null } = {}) {
  const original = definition.execute
  return {
    ...definition,
    async execute(args, exec) {
      const started = Date.now()
      let result
      let thrown = null
      try {
        result = await original(args, exec)
        return result
      } catch (err) {
        thrown = err
        throw err
      } finally {
        const entry = {
          time: new Date().toISOString(),
          source: 'tool',
          tool: definition.name,
          plugin_version: pluginVersion,
          duration_ms: Date.now() - started,
          target: audit.describePath(targetOf(args))
        }
        // 工具按约定不抛裸异常，失败走信封；真抛了也要能审计到
        if (thrown) {
          entry.ok = false
          entry.error_code = thrown.code ?? 'THROWN'
          entry.error_message = String(thrown.message ?? '').slice(0, 200)
        } else {
          entry.ok = result?.success === true
          if (result?.success !== true) entry.error_code = result?.error?.code ?? null
        }
        entry.request_id = result?.request_id ?? null
        if (Array.isArray(result?.changes)) entry.changes_count = result.changes.length
        if (typeof result?.output_file?.size === 'number') entry.output_bytes = result.output_file.size
        audit.record(entry)
      }
    }
  }
}

/**
 * 批量套上审计中间件。
 * @param {object[]} definitions - 工具定义列表。
 * @param {ToolAuditLog} audit - 审计日志。
 * @param {object} [options] - 选项。
 * @returns {object[]} 包装后的定义列表。
 */
export function withToolAuditAll(definitions, audit, options = {}) {
  return definitions.map((definition) => withToolAudit(definition, audit, options))
}
