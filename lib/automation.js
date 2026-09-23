/**
 * 本地 Office/WPS 联动（阶段 6）：Node 侧封装。
 *
 * 分工：
 *   - 真正的 COM 调用全在 `office-automation.ps1`（固定模板、纯 ASCII、参数化）；
 *   - 这里负责开关、超时、进程回收兜底、结果解析与审计日志。
 *
 * 安全约定（开发要求 §十二.4）：
 *   - **默认关闭**：只有插件配置 `allowLocalAutomation: true` 才会启动本机 Office；
 *   - 独立进程：引擎在它自己的进程里跑，插件只通过结果 JSON 读结论；
 *   - 禁止宏/外链/外模板/弹窗：由脚本里的 AutomationSecurity=3、AskToUpdateLinks=false、
 *     DisplayAlerts=false、-NonInteractive 保证；
 *   - 超时：到点先杀 PowerShell 进程树，再按状态文件补一次 cleanup（只杀本次启动的引擎 PID）；
 *   - 审计：每次调用追加一行 JSONL（只记动作、引擎、文件名、字节数、耗时与错误，**不记文档内容**）；
 *   - **永不改动传入的文件**：输入只读打开，产物写到调用方指定的新路径。
 *
 * @module dsh-exp-office/automation
 */

import { spawn } from 'node:child_process'
import { appendFileSync, existsSync, mkdirSync, readFileSync, rmSync } from 'node:fs'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficeError } from './errors.js'

/** 支持自动化的文件类型。 */
export const AUTOMATION_KINDS = Object.freeze(['xlsx', 'xlsm', 'docx', 'docm', 'pptx', 'pptm'])

/** 可指定的引擎。 */
export const AUTOMATION_ENGINES = Object.freeze(['auto', 'excel', 'word', 'powerpoint', 'wps'])

/** 与文件类型对应的默认组件（用于校验引擎选择是否合理）。 */
const ENGINE_FOR_KIND = Object.freeze({
  xlsx: ['auto', 'excel', 'wps'],
  xlsm: ['auto', 'excel', 'wps'],
  docx: ['auto', 'word', 'wps'],
  docm: ['auto', 'word', 'wps'],
  pptx: ['auto', 'powerpoint', 'wps'],
  pptm: ['auto', 'powerpoint', 'wps']
})

const SCRIPT_PATH = join(dirname(fileURLToPath(import.meta.url)), 'office-automation.ps1')

/** 默认超时：2 分钟。启动 Office 本身可能十几秒，重算大表会更久。 */
export const DEFAULT_AUTOMATION_TIMEOUT_MS = 120000

/**
 * 本地引擎联动。
 */
export class OfficeAutomation {
  #enabled
  #powershell
  #scriptPath
  #scratchDir
  #auditPath
  #timeoutMs

  /**
   * @param {object} options - 选项。
   * @param {boolean} [options.enabled] - 是否允许启动本机 Office/WPS（默认 false）。
   * @param {string} [options.powershell] - PowerShell 可执行文件（默认 Windows 自带的 powershell.exe）。
   * @param {string} [options.scriptPath] - 自动化脚本路径（测试可替换成桩脚本）。
   * @param {string} [options.scratchDir] - 临时目录（每次调用在其下建子目录）。
   * @param {string} [options.auditPath] - 审计日志路径（JSONL）。
   * @param {number} [options.timeoutMs] - 默认超时。
   */
  constructor({
    enabled = false,
    powershell = process.env.DSH_OFFICE_POWERSHELL ?? 'powershell.exe',
    scriptPath = SCRIPT_PATH,
    scratchDir = null,
    auditPath = null,
    timeoutMs = DEFAULT_AUTOMATION_TIMEOUT_MS
  } = {}) {
    this.#enabled = enabled === true
    this.#powershell = powershell
    this.#scriptPath = scriptPath
    this.#scratchDir = scratchDir
    this.#auditPath = auditPath
    this.#timeoutMs = timeoutMs
  }

  /** @returns {boolean} 是否已允许本地联动。 */
  get enabled() {
    return this.#enabled
  }

  /** @returns {number} 默认超时（毫秒）。 */
  get timeoutMs() {
    return this.#timeoutMs
  }

  /**
   * 检测本机引擎（只读注册表；`probeCom` 时才真的创建 COM 实例）。
   * @param {object} [options] - 选项。
   * @param {boolean} [options.probeCom] - 是否实际创建 COM 实例验证可用性。
   * @returns {Promise<object>} `{engines, microsoft_office, wps}`。
   */
  async detectEngines({ probeCom = false } = {}) {
    // 只读注册表的检测不需要授权（它不启动任何程序）；真的去创建 COM 实例才需要。
    const result = await this.#invoke({ action: 'detect', probeCom, requiresEnabled: probeCom === true })
    const engines = Array.isArray(result.engines) ? result.engines : []
    return {
      engines,
      microsoft_office: engines.some((e) => e.id !== 'wps' && e.installed),
      wps: engines.some((e) => e.id === 'wps' && e.installed),
      probed_com: probeCom === true,
      started_pids: result.started_pids ?? [],
      leftover_pids: result.leftover_pids ?? [],
      elapsed_ms: result.elapsed_ms ?? null
    }
  }

  /**
   * 用真实引擎**重新计算**并另存为新文件（Excel 的公式缓存、Word 的域都会更新）。
   *
   * 输入文件只读打开，**不会被改动**；产物写到 `outputPath`。
   *
   * @param {object} args - 参数。
   * @param {string} args.inputPath - 输入文件（绝对路径）。
   * @param {string} args.outputPath - 输出文件（绝对路径，覆盖已存在文件需调用方自行确认）。
   * @param {string} args.kind - 文件类型（xlsx/xlsm/docx/docm/pptx/pptm）。
   * @param {string} [args.engine] - auto/excel/word/powerpoint/wps。
   * @param {number} [args.timeoutMs] - 本次超时。
   * @returns {Promise<object>} 执行结果。
   */
  async recalculate({ inputPath, outputPath, kind, engine = 'auto', timeoutMs = this.#timeoutMs }) {
    return this.#runOnCopy({ action: 'recalc', inputPath, outputPath, kind, engine, timeoutMs })
  }

  /**
   * 用真实引擎**重新渲染**（打开并用引擎自己重写一份，等价于「重新保存」）。
   * @param {object} args - 参数，同 {@link OfficeAutomation#recalculate}。
   * @returns {Promise<object>} 执行结果。
   */
  async rerender({ inputPath, outputPath, kind, engine = 'auto', timeoutMs = this.#timeoutMs }) {
    return this.#runOnCopy({ action: 'rerender', inputPath, outputPath, kind, engine, timeoutMs })
  }

  /**
   * 执行一次「引擎打开 → 处理 → 另存」。
   * @param {object} args - 参数。
   * @returns {Promise<object>} 执行结果。
   */
  async #runOnCopy({ action, inputPath, outputPath, kind, engine, timeoutMs }) {
    if (!AUTOMATION_KINDS.includes(kind)) {
      throw new OfficeError('UNSUPPORTED_FILE_TYPE', `本地联动只支持 ${AUTOMATION_KINDS.join(' / ')}，收到 ${kind}。`, {
        supported: [...AUTOMATION_KINDS]
      })
    }
    if (!AUTOMATION_ENGINES.includes(engine)) {
      throw new OfficeError('INVALID_REQUEST', `engine 只能是 ${AUTOMATION_ENGINES.join(' / ')}，收到 ${engine}。`)
    }
    if (!ENGINE_FOR_KIND[kind].includes(engine)) {
      throw new OfficeError('INVALID_REQUEST', `${kind} 不能用 ${engine} 打开（可用：${ENGINE_FOR_KIND[kind].join(' / ')}）。`)
    }
    if (typeof inputPath !== 'string' || typeof outputPath !== 'string' || inputPath === '' || outputPath === '') {
      throw new OfficeError('INVALID_REQUEST', 'inputPath 与 outputPath 都必须是非空路径。')
    }
    const result = await this.#invoke({ action, kind, engine, inputPath, outputPath, timeoutMs })
    return {
      action,
      engine: result.engine,
      prog_id: result.prog_id ?? null,
      engine_version: result.version ?? null,
      input_path: inputPath,
      output_path: outputPath,
      output_size: result.output_size ?? null,
      elapsed_ms: result.elapsed_ms ?? null,
      warnings: result.warnings ?? [],
      killed_pids: result.killed_pids ?? [],
      leftover_pids: result.leftover_pids ?? []
    }
  }

  /**
   * 调用 PowerShell 脚本并解析结果 JSON。
   *
   * 结果走**文件**而不是 stdout：中文路径与 PowerShell 的输出编码在很多机器上会打架，
   * 而且不捕获管道输出也避开了某些受限环境下 `stdio: 'pipe'` 的限制。
   *
   * @param {object} args - 参数。
   * @param {string} args.action - 脚本动作。
   * @param {string} [args.kind] - 文件类型。
   * @param {string} [args.engine] - 引擎。
   * @param {string} [args.inputPath] - 输入路径。
   * @param {string} [args.outputPath] - 输出路径。
   * @param {boolean} [args.probeCom] - 是否探测 COM。
   * @param {number} [args.timeoutMs] - 超时。
   * @param {boolean} [args.requiresEnabled] - 是否要求已开启本地联动（默认 true；只读注册表的检测传 false）。
   * @returns {Promise<object>} 脚本返回的结果对象。
   */
  async #invoke({ action, kind, engine, inputPath, outputPath, probeCom = false, timeoutMs = this.#timeoutMs, requiresEnabled = true }) {
    if (process.platform !== 'win32') {
      throw new OfficeError('OFFICE_NOT_INSTALLED', '本地 Office/WPS 联动只在 Windows 上可用。', { platform: process.platform })
    }
    if (requiresEnabled && !this.#enabled) {
      throw new OfficeError(
        'AUTOMATION_DISABLED',
        '本地 Office/WPS 联动默认关闭：它会在本机启动 Office/WPS 进程。需要插件配置 allowLocalAutomation=true 才允许。'
      )
    }
    if (!existsSync(this.#scriptPath)) {
      throw new OfficeError('INTERNAL_ERROR', `自动化脚本不存在：${this.#scriptPath}`)
    }

    const scratch = join(this.#scratchDir ?? process.cwd(), `automation-${globalThis.crypto.randomUUID()}`)
    mkdirSync(scratch, { recursive: true })
    const resultPath = join(scratch, 'result.json')
    const statePath = join(scratch, 'state.json')
    const args = [
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      this.#scriptPath,
      '-Action',
      action,
      '-ResultPath',
      resultPath,
      '-StatePath',
      statePath
    ]
    if (kind) args.push('-Kind', kind)
    if (engine) args.push('-Engine', engine)
    if (inputPath) args.push('-InputPath', inputPath)
    if (outputPath) args.push('-OutputPath', outputPath)
    if (probeCom) args.push('-ProbeCom')

    const started = Date.now()
    let timedOut = false
    let exitCode = null
    let child = null
    try {
      child = spawn(this.#powershell, args, { stdio: 'ignore', windowsHide: true })
      const outcome = await new Promise((resolve) => {
        const timer = setTimeout(() => {
          timedOut = true
          resolve('timeout')
        }, Math.max(5000, timeoutMs))
        child.once('error', (err) => {
          clearTimeout(timer)
          resolve(`error:${err.message}`)
        })
        child.once('exit', (code) => {
          clearTimeout(timer)
          exitCode = code
          resolve('exit')
        })
      })

      if (outcome === 'timeout') {
        await this.#killTree(child.pid)
        await this.#cleanupAfterTimeout(statePath)
        this.#audit({ action, engine, ok: false, error: 'TIMEOUT', elapsed_ms: Date.now() - started })
        throw new OfficeError('TIMEOUT', `本地引擎执行超过 ${timeoutMs} ms，已终止并清理本次启动的进程。`, {
          action,
          timeout_ms: timeoutMs
        })
      }
      if (String(outcome).startsWith('error:')) {
        this.#audit({ action, engine, ok: false, error: String(outcome).slice(7), elapsed_ms: Date.now() - started })
        throw new OfficeError('INTERNAL_ERROR', `无法启动 PowerShell：${String(outcome).slice(7)}`)
      }

      if (!existsSync(resultPath)) {
        this.#audit({ action, engine, ok: false, error: 'NO_RESULT', elapsed_ms: Date.now() - started })
        throw new OfficeError('AUTOMATION_FAILED', `本地引擎没有返回结果（退出码 ${exitCode}），已清理它启动的进程。`, {
          action,
          exit_code: exitCode
        })
      }
      const result = JSON.parse(readFileSync(resultPath, 'utf8'))
      this.#audit({
        action,
        engine: result.engine ?? engine ?? null,
        // 审计只记**文件名**，不记绝对路径，更不记文档内容（开发要求 §12.3）
        input_name: inputPath ? basename(inputPath) : null,
        output_name: outputPath ? basename(outputPath) : null,
        ok: result.ok === true,
        error: result.ok === true ? null : (result.error ?? 'unknown'),
        elapsed_ms: result.elapsed_ms ?? Date.now() - started,
        output_size: result.output_size ?? null,
        leftover_pids: result.leftover_pids ?? []
      })
      if (result.ok !== true) {
        const message = String(result.error ?? '未知错误')
        const code = /could be created|不可用|not found|no engine/i.test(message)
          ? result.engine === 'wps'
            ? 'WPS_NOT_INSTALLED'
            : 'OFFICE_NOT_INSTALLED'
          : 'AUTOMATION_FAILED'
        throw new OfficeError(code, `本地引擎执行失败：${message}`, {
          action,
          engine: result.engine ?? null,
          warnings: result.warnings ?? []
        })
      }
      return result
    } finally {
      rmSync(scratch, { recursive: true, force: true })
    }
  }

  /**
   * 杀掉一个进程树（只用于本次自己启动的 PowerShell）。
   * @param {number} pid - 进程号。
   * @returns {Promise<void>} 完成后 resolve。
   */
  async #killTree(pid) {
    if (!pid) return
    await new Promise((resolve) => {
      const killer = spawn('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore', windowsHide: true })
      killer.once('exit', () => resolve())
      killer.once('error', () => resolve())
    })
  }

  /**
   * 超时后的兜底清理：按状态文件里的 PID 收尾。
   *
   * 状态文件由脚本在开始时写入，记录**本次开始前已存在的引擎 PID**；
   * 清理动作只杀「不在这个集合里」的引擎进程，因此不会碰用户自己开着的 Office/WPS。
   *
   * @param {string} statePath - 状态文件路径。
   * @returns {Promise<void>} 完成后 resolve。
   */
  async #cleanupAfterTimeout(statePath) {
    try {
      if (!existsSync(statePath)) return
      const scratch = dirname(statePath)
      const resultPath = join(scratch, 'cleanup.json')
      await new Promise((resolve) => {
        const child = spawn(
          this.#powershell,
          ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', this.#scriptPath, '-Action', 'cleanup', '-ResultPath', resultPath, '-StatePath', statePath],
          { stdio: 'ignore', windowsHide: true }
        )
        const timer = setTimeout(() => {
          try {
            child.kill()
          } catch {
            /* 已经退出 */
          }
          resolve()
        }, 30000)
        child.once('exit', () => {
          clearTimeout(timer)
          resolve()
        })
        child.once('error', () => {
          clearTimeout(timer)
          resolve()
        })
      })
    } catch {
      // 兜底清理本身失败不覆盖主错误
    }
  }

  /**
   * 追加一条审计记录（不含文档内容）。
   * @param {object} entry - 记录。
   * @returns {void}
   */
  #audit(entry) {
    if (!this.#auditPath) return
    try {
      mkdirSync(dirname(this.#auditPath), { recursive: true })
      appendFileSync(this.#auditPath, `${JSON.stringify({ time: new Date().toISOString(), ...entry })}\n`, 'utf8')
    } catch {
      // 审计写不进去不能影响主流程
    }
  }
}
