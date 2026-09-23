/**
 * 文档转 PDF：接入宿主内置的 LibreOffice 引擎。
 *
 * 设计取舍：
 *   - **不引入依赖**。`@deepseek-ai/libreoffice-kit` 由宿主提供，这里延迟解析并按需加载；
 *     解析不到时返回明确的 `OFFICE_NOT_INSTALLED`，而不是让整个插件加载失败。
 *   - **不自己起进程**。引擎自带沙箱、超时、取消、字节上限与输出路径保护，
 *     自己再包一层只会更差。
 *   - `render()` 返回的 `missingFonts` 直接补上了 `validate_docx` 里
 *     「字体替换」这一项 —— 那是读 XML 无法判定、必须靠渲染引擎回答的问题。
 *
 * @module dsh-exp-office/convert
 */

import { existsSync } from 'node:fs'
import { stat } from 'node:fs/promises'
import { basename } from 'node:path'
import { OfficeError } from './errors.js'
import { resolveHostModule } from './define-tool.js'

/** LibreOffice 引擎能接受的输入后缀。 */
export const CONVERTIBLE_EXTENSIONS = Object.freeze(['docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt'])

/** 转换任务的默认超时（含引擎冷启动与字体装载）。 */
const DEFAULT_TIMEOUT_MS = 180000

/** 已创建的转换器缓存：引擎启动昂贵，同进程内复用。 */
let converterPromise = null

/**
 * 取得（并缓存）一个转换器实例。
 *
 * 引擎按「串行队列」工作，并发渲染会排队而不是并行占内存，因此单实例是安全的。
 *
 * @param {object} [options] - 引擎选项。
 * @returns {Promise<object>} 转换器实例。
 */
async function getConverter(options = {}) {
  converterPromise ??= (async () => {
    let mod
    try {
      mod = await resolveHostModule('@deepseek-ai/libreoffice-kit')
    } catch (err) {
      throw new OfficeError('OFFICE_NOT_INSTALLED', `未找到宿主内置的 LibreOffice 引擎：${err.message}`, {
        solution: '确认 @deepseek-ai/libreoffice-kit 已随 DSH 安装；否则请改用插件自身的独立处理能力。'
      })
    }
    if (typeof mod.createConverter !== 'function') {
      throw new OfficeError('OFFICE_NOT_INSTALLED', '已解析到 libreoffice-kit，但不含 createConverter 导出。')
    }
    const converter = await mod.createConverter({ timeoutMs: DEFAULT_TIMEOUT_MS, ...options })
    return { converter, mod }
  })()
  try {
    return await converterPromise
  } catch (err) {
    converterPromise = null // 失败不缓存，允许后续重试
    throw err
  }
}

/**
 * 释放引擎。任务结束后调用可回收 LibreOffice 进程与临时资源。
 * @returns {Promise<void>} 完成信号。
 */
export async function disposeConverter() {
  if (!converterPromise) return
  const pending = converterPromise
  converterPromise = null
  try {
    const { converter } = await pending
    await converter.dispose()
  } catch {
    // 释放失败不应掩盖原始错误
  }
}

/**
 * 把一个办公文档转换为 PDF。
 *
 * @param {object} args - 参数。
 * @param {string} args.inputPath - 输入绝对路径。
 * @param {string} args.outputPath - 输出绝对路径（必须不存在，由引擎负责拒绝覆盖）。
 * @param {AbortSignal} [args.signal] - 取消信号。
 * @param {number} [args.timeoutMs] - 超时。
 * @returns {Promise<{backend: string, missingFonts: string[], size: number, pages: number|null}>} 转换结果。
 */
export async function convertToPdf({ inputPath, outputPath, signal, timeoutMs }) {
  if (!existsSync(inputPath)) {
    throw new OfficeError('FILE_NOT_FOUND', `输入文件不存在：${basename(inputPath)}`)
  }
  if (existsSync(outputPath)) {
    throw new OfficeError('PERMISSION_DENIED', `输出文件已存在：${basename(outputPath)}。转换不会覆盖既有文件。`, {
      needsConfirmation: true
    })
  }
  const { converter } = await getConverter(timeoutMs ? { timeoutMs } : {})
  let result
  try {
    result = await converter.render({ inputPath, outputPath }, signal)
  } catch (err) {
    if (err?.code === 'CONVERSION_TIMEOUT') throw new OfficeError('TIMEOUT', `转换超时：${err.message}`)
    if (err?.code === 'CONVERSION_CANCELLED') throw new OfficeError('TIMEOUT', '转换已取消。', { retryable: true })
    throw new OfficeError('INTERNAL_ERROR', `转换失败：${err?.message ?? err}`)
  }
  const info = await stat(outputPath).catch(() => null)
  const bytes = info ? await (await import('node:fs/promises')).readFile(outputPath) : Buffer.alloc(0)
  return {
    backend: result.backend,
    missingFonts: result.missingFonts ?? [],
    size: info?.size ?? 0,
    pages: countPdfPages(bytes)
  }
}

/**
 * 统计 PDF 页数。
 *
 * 只做轻量扫描：找到 `/Type /Page` 对象计数。够用且不需要引入 PDF 解析库；
 * 对压缩对象流（ObjStm）内的页面对象会漏计，因此调用方应把它当作下界。
 *
 * @param {Buffer} bytes - PDF 字节。
 * @returns {number|null} 页数；非 PDF 时返回 null。
 */
export function countPdfPages(bytes) {
  if (bytes.length < 5 || bytes.subarray(0, 5).toString('latin1') !== '%PDF-') return null
  const text = bytes.toString('latin1')
  const matches = text.match(/\/Type\s*\/Page[^s]/g)
  return matches ? matches.length : null
}
