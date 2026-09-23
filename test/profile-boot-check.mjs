/**
 * Profile 启动自检：不开服务，按 DSH 的真实加载路径确认插件能被装载。
 *
 * 起因：2026-09-23 的 `dsh web` 启动失败 —— `apply` 返回了普通对象摘要，
 * Cordis 的 safeCollect 只接受「清理函数 / null / undefined」，于是整个插件树加载失败。
 * `dsh --profile web --dump-config` 只组合配置树、**不会 apply 插件**，因此当时也没拦住。
 *
 * 这个脚本做的三件事，正对应当前唯一还能出错的三个环节：
 *   1. 从 profile 目录**按包名解析**插件模块（真实解析路径，而不是工作区相对路径）
 *   2. 用 profile 里那份真实 Cordis 的 Context + provide('tools') 装载并 await
 *   3. 断言注册出的工具数与 dsh.plugin.json 声明一致
 *
 * 用法：node test/profile-boot-check.mjs [profileDir]
 * 默认 profile 目录：%USERPROFILE%\.dsh\profiles\web
 */
import { createRequire } from 'node:module'
import os from 'node:os'
import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const profileDir = process.argv[2] ?? `${os.homedir()}/.dsh/profiles/web`
if (!existsSync(profileDir)) {
  console.log(`⏭️  跳过：profile 目录不存在 ${profileDir}`)
  process.exit(0)
}

const require = createRequire(join(profileDir, 'noop.js'))

/** 依次尝试多个锚点解析模块，返回文件路径。 */
function resolveFirst(specifiers) {
  for (const specifier of specifiers) {
    try {
      return require.resolve(specifier)
    } catch {
      // 继续尝试下一个锚点
    }
  }
  return null
}

const pluginPath = resolveFirst(['dsh-exp-office'])
if (!pluginPath) {
  console.log(`❌ 从 ${profileDir} 解析不到 dsh-exp-office（profile 尚未安装该插件？）`)
  process.exit(1)
}
const cordisPath = resolveFirst([
  '@deepseek-ai/cordis',
  `${os.homedir()}/AppData/Roaming/npm/node_modules/@deepseek-ai/dsh/node_modules/@deepseek-ai/cordis/lib/index.js`
])
if (!cordisPath) {
  console.log('❌ 解析不到 @deepseek-ai/cordis')
  process.exit(1)
}

console.log(`插件模块：${pluginPath}`)
console.log(`Cordis  ：${cordisPath}`)

// 先查副本漂移：profile 里是 pnpm 复制出来的独立副本，改完源码不刷新就会「源码已修、线上仍旧」。
const installedRoot = join(pluginPath, '..', '..')
const workspaceRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const drift = []
for (const relative of ['dsh.plugin.json', 'cordis.patch.yml', 'package.json']) {
  const a = join(workspaceRoot, relative)
  const b = join(installedRoot, relative)
  if (!existsSync(a) || !existsSync(b)) continue
  if (readFileSync(a).compare(readFileSync(b)) !== 0) drift.push(relative)
}
const libDir = join(workspaceRoot, 'lib')
for (const name of readdirSync(libDir)) {
  if (!name.endsWith('.js')) continue
  const b = join(installedRoot, 'lib', name)
  if (!existsSync(b) || readFileSync(join(libDir, name)).compare(readFileSync(b)) !== 0) drift.push(`lib/${name}`)
}
if (drift.length > 0) {
  console.log(`❌ 安装副本与工作区不一致（${drift.length} 个文件）：${drift.join('、')}`)
  console.log('   修复：cd ' + profileDir + ' && pnpm install --force')
  process.exit(1)
}
console.log(`✅ 安装副本与工作区逐字节一致（${drift.length} 处漂移）`)

const plugin = await import(pathToFileURL(pluginPath).href)
const { Context } = await import(pathToFileURL(cordisPath).href)

const declared = JSON.parse(readFileSync(join(pluginPath, '..', '..', 'dsh.plugin.json'), 'utf8')).contributes.tools.length
const registered = new Map()
const ctx = new Context()
ctx.provide('tools', {
  register(definition) {
    registered.set(definition.name, definition)
  }
})

// 这一行就是 dsh 启动时做的事：apply 抛错或返回非法 effect 都会在这里被拒。
await ctx.plugin({ apply: plugin.apply, name: plugin.name, inject: plugin.inject }, {})
if (registered.size !== declared) {
  console.log(`❌ 注册工具数 ${registered.size} 与清单声明的 ${declared} 不一致`)
  process.exit(1)
}
console.log(`✅ 装载成功：注册 ${registered.size} 个工具，与 dsh.plugin.json 声明一致`)
console.log('✅ profile 启动路径自检通过（未启动任何服务）')
process.exit(0)
