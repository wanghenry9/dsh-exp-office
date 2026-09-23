/**
 * `defineTool` 解析器。
 *
 * DSH 插件的惯例是把 `@deepseek-ai/dsh-tools` 声明为 peerDependency 并直接 import
 * （见 dsh-better-sidebar）。但插件被不同方式挂载时，模块解析根可能不同，因此这里
 * 按优先级分层解析，并在全部失败时给出可诊断的错误，而不是加载期崩溃。
 *
 * 优先级：
 *   1. 显式注入（config.defineTool / ctx.__defineTool）—— 测试与宿主注入用
 *   2. 常规 import('@deepseek-ai/dsh-tools')
 *   3. 从候选根目录用 createRequire 解析（DSH 安装目录、profile 目录、环境变量）
 *
 * @module dsh-exp-office/define-tool
 */

import { createRequire } from 'node:module'
import { existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { pathToFileURL } from 'node:url'

/**
 * 猜测可能的 DSH 安装根目录。
 * @returns {string[]} 候选目录（存在的才返回）。
 */
function candidateRoots() {
  const roots = []
  const envHome = process.env.DSH_HOME ?? process.env.DSH_PROFILE_DIR
  if (envHome) roots.push(envHome)
  // 插件自身位于 profile 的 node_modules 下时，profile 根就是上两级
  const here = dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1'))
  roots.push(join(here, '..', '..', '..'))
  // 全局 npm 安装的 DSH
  const npmGlobal = process.env.APPDATA ? join(process.env.APPDATA, 'npm', 'node_modules', '@deepseek-ai', 'dsh', 'node_modules', '@deepseek-ai') : null
  if (npmGlobal) roots.push(npmGlobal)
  return roots.filter((r) => r && existsSync(r))
}

/**
 * 解析 `defineTool` 实现。
 *
 * @param {object} [options] - 选项。
 * @param {Function} [options.injected] - 显式注入的实现。
 * @returns {Promise<{defineTool: Function, source: string}>} 解析结果与来源说明。
 */
/**
 * 解析宿主提供的模块。
 *
 * 先试常规 import，再逐个候选根目录用 createRequire 解析 —— 插件被不同方式
 * 挂载时模块解析根可能不同，硬绑一种方式会让插件在别的部署里加载失败。
 *
 * @param {string} specifier - 模块名。
 * @returns {Promise<object>} 模块命名空间。
 */
export async function resolveHostModule(specifier) {
  try {
    return await import(specifier)
  } catch {
    // 继续尝试候选根目录
  }
  for (const root of candidateRoots()) {
    try {
      const require = createRequire(join(root, 'noop.js'))
      const resolved = require.resolve(specifier)
      return await import(pathToFileURL(resolved).href)
    } catch {
      // 试下一个候选
    }
  }
  throw new Error(`无法解析宿主模块 ${specifier}`)
}

/**
 * 解析 `defineTool` 实现。
 * @param {object} [options] - 选项。
 * @param {Function} [options.injected] - 显式注入的实现。
 * @returns {Promise<{defineTool: Function, source: string}>} 解析结果与来源说明。
 */
export async function resolveDefineTool({ injected } = {}) {
  if (typeof injected === 'function') return { defineTool: injected, source: 'injected' }
  try {
    const mod = await resolveHostModule('@deepseek-ai/dsh-tools')
    if (typeof mod.defineTool === 'function') return { defineTool: mod.defineTool, source: 'host-module' }
  } catch {
    // 落到下面的统一报错
  }
  throw new Error(
    'dsh-exp-office 无法解析 @deepseek-ai/dsh-tools。请确认该包与插件在同一模块解析根下，' +
      '或通过 config.defineTool 显式注入。'
  )
}
