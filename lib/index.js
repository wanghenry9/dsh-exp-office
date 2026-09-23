/**
 * dsh-exp-office：DeepSeek Harness 的 Office/WPS 办公文件插件入口。
 *
 * 这是一个 Cordis 插件：把办公文档工具注册进 `ctx.tools`，注册后 schema 自动
 * 进入系统提示词，宿主平面的 agent 均可调用。
 *
 * 兼容性策略（docs §十）：
 *   - 通过 `inject: ['tools']` 做能力检测，而不是依赖 Harness 版本号；
 *   - 工具 schema 只用协议级通用词汇，不硬编码任何 Harness 内部字段；
 *   - 通过 capabilities 模块对外声明插件版本、协议版本与能力清单。
 *
 * @module dsh-exp-office
 */

import { createOfficeTools } from './tools.js'
import { resolveDefineTool } from './define-tool.js'
import { toCapabilityManifest } from './capabilities.js'

const name = 'dsh-exp-office'
const inject = ['tools']

/**
 * 注册全部办公工具的 Cordis 插件体。
 *
 * @param {object} ctx - Cordis 上下文，需携带 `ctx.tools`。
 * @param {object} [config] - 插件配置。
 * @param {string} [config.workspaceRoot] - 允许读写的工作区根目录，默认进程工作目录。
 * @param {boolean} [config.allowOverwriteOriginals] - 是否允许覆盖原文件，默认 false。
 * @param {number} [config.maxFileBytes] - 单文件大小上限。
 * @param {number} [config.maxEditableCells] - 单张工作表可安全编辑的单元格数上限，默认 30 万（写入要把整表解析成 DOM，见 xlsx.js）。
 * @param {boolean} [config.allowLocalAutomation] - 是否允许本地 Office/WPS 联动（重算/重渲染），**默认 false**。
 * @param {number} [config.automationTimeoutMs] - 本地引擎单次调用超时，默认 120000。
 * @param {string} [config.powershellPath] - 指定 PowerShell 可执行文件（默认 Windows 自带的 powershell.exe）。
 * @param {Function} [config.defineTool] - 显式注入 defineTool（测试或特殊宿主用）。
 * @returns {Promise<void>} Cordis effect：resolve 后不得返回普通对象，否则报 Invalid effect。
 */
async function apply(ctx, config = {}) {
  if (!ctx?.tools?.register) {
    throw new Error('dsh-exp-office 需要宿主提供 ctx.tools 服务；请确认 @deepseek-ai/dsh-tools 已挂载。')
  }
  const { defineTool, source } = await resolveDefineTool({ injected: config.defineTool ?? ctx.__defineTool })
  const definitions = createOfficeTools({ defineTool, config })
  for (const definition of definitions) ctx.tools.register(definition)
  const manifest = toCapabilityManifest()
  ctx.logger?.info?.(
    `[dsh-exp-office] 已注册 ${definitions.length} 个办公工具（defineTool 来源：${source}），` +
      `插件 v${manifest.plugin_version}，能力 ${manifest.capabilities.length} 项`
  )
}

export { apply, name, inject, createOfficeTools, toCapabilityManifest }
export default { apply, name, inject }
