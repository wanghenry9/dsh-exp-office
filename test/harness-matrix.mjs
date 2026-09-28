/**
 * Harness 版本兼容矩阵：把插件装进**每个 DSH 版本自带的** `cordis` + `dsh-tools` 组合里，
 * 验证「声明支持 >=0.1.0-rc.6」这句话在真实版本上是成立的。
 *
 * 为什么这样测就够了：本插件的 `lib/` **不 import 任何 @deepseek-ai/* 包** ——
 * `defineTool` 是宿主通过 `apply(ctx, { defineTool })` 注入进来的，`cordis` 的
 * `Context` 也由宿主提供。所以兼容面就是这两样东西，而它们在 npm 上按版本发布。
 *
 * 每个版本做四件事：
 *   1. 用那个版本的 `defineTool` 定义全部工具（值 schema DSL 能不能吃下我们的 90 个定义）
 *   2. 直接看 `apply()` 的返回值是不是合法 Cordis effect（undefined / null / disposer）
 *   3. 用那个版本的**真实 Cordis Context** 装载（`provide('tools')` + `ctx.plugin`）
 *   4. 跑三个真实调用：生成一份 PDF、列目录、以及一个必须结构化报错的坏输入
 *
 * 用法：
 *   node test/harness-matrix.mjs                      # 默认版本集
 *   node test/harness-matrix.mjs 0.1.7-rc.2 0.1.6-alpha.2
 *   node test/harness-matrix.mjs --json               # 机器可读输出
 *   node test/harness-matrix.mjs --keep               # 保留临时目录
 *
 * 退出码：0 = 全部通过；2 = 一个版本都没装上（通常是没网）；3 = 有版本不兼容。
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { createRequire } from 'node:module'
import { execFileSync } from 'node:child_process'
import { pathToFileURL } from 'node:url'

const ROOT = path.resolve(import.meta.dirname, '..')
const argv = process.argv.slice(2)
const asJson = argv.includes('--json')
const keep = argv.includes('--keep')
const DEFAULT_VERSIONS = ['0.1.2-rc.1', '0.1.5-rc.1', '0.1.5-rc.3', '0.1.6-alpha.2', '0.1.7-rc.2']
const versions = argv.filter((a) => !a.startsWith('--'))
const targets = versions.length > 0 ? versions : DEFAULT_VERSIONS

const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'dsh.plugin.json'), 'utf8'))
const declaredTools = manifest.contributes.tools.length

/**
 * 给嵌套 npm 一份干净环境（外层 `npm run` 注入的 `npm_config_*` 会改变子进程行为）。
 * @returns {NodeJS.ProcessEnv} 子进程环境。
 */
function cleanEnv() {
  const env = {}
  for (const [key, value] of Object.entries(process.env)) {
    if (key.toLowerCase().startsWith('npm_config_')) continue
    env[key] = value
  }
  return env
}

const npmCli = path.join(path.dirname(process.execPath), 'node_modules', 'npm', 'bin', 'npm-cli.js')

/**
 * 在指定目录里跑 npm。
 * @param {string[]} args - npm 参数。
 * @param {string} cwd - 工作目录。
 * @returns {string} 标准输出。
 */
function npm(args, cwd) {
  const command = fs.existsSync(npmCli) ? process.execPath : process.platform === 'win32' ? 'npm.cmd' : 'npm'
  const finalArgs = fs.existsSync(npmCli) ? [npmCli, ...args] : args
  return execFileSync(command, finalArgs, {
    cwd,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    env: cleanEnv(),
    shell: !fs.existsSync(npmCli) && process.platform === 'win32'
  })
}

/**
 * 解析一个包在给定 node_modules 里的 ESM 入口。
 * @param {string} dir - 含 node_modules 的目录。
 * @param {string} name - 包名。
 * @returns {string} 入口文件绝对路径。
 */
function resolveEntry(dir, name) {
  const pkgDir = path.join(dir, 'node_modules', ...name.split('/'))
  const pkg = JSON.parse(fs.readFileSync(path.join(pkgDir, 'package.json'), 'utf8'))
  const exportsField = pkg.exports ?? {}
  const root = typeof exportsField === 'string' ? exportsField : exportsField['.'] ?? {}
  const candidate =
    (typeof root === 'string' ? root : root.import ?? root.default ?? root.require) ?? pkg.module ?? pkg.main ?? 'index.js'
  const resolved = path.resolve(pkgDir, candidate)
  if (fs.existsSync(resolved)) return resolved
  // 有些包把 exports 写成条件对象的数组，退回到 main
  const fallback = path.resolve(pkgDir, pkg.main ?? 'index.js')
  if (fs.existsSync(fallback)) return fallback
  throw new Error(`找不到 ${name} 的入口（exports=${JSON.stringify(exportsField)}）`)
}

/**
 * 跑一个版本的完整检查。
 * @param {string} version - dsh-tools 版本。
 * @returns {Promise<object>} 结果行。
 */
async function checkVersion(version) {
  const work = fs.mkdtempSync(path.join(os.tmpdir(), `dsh-matrix-${version}-`))
  const row = { version, cordis: null, tools: null, defined: null, effect_ok: null, loaded: null, smoke: null, error: null }
  try {
    fs.writeFileSync(path.join(work, 'package.json'), JSON.stringify({ name: 'matrix-probe', private: true, version: '0.0.0' }, null, 2))
    npm(['install', '--no-audit', '--no-fund', '--ignore-scripts', `@deepseek-ai/dsh-tools@${version}`], work)

    const cordisEntry = resolveEntry(work, '@deepseek-ai/cordis')
    const toolsEntry = resolveEntry(work, '@deepseek-ai/dsh-tools')
    row.cordis = JSON.parse(fs.readFileSync(path.join(work, 'node_modules/@deepseek-ai/cordis/package.json'), 'utf8')).version
    row.tools = JSON.parse(fs.readFileSync(path.join(work, 'node_modules/@deepseek-ai/dsh-tools/package.json'), 'utf8')).version

    const { Context } = await import(pathToFileURL(cordisEntry).href)
    const { defineTool } = await import(pathToFileURL(toolsEntry).href)
    if (typeof defineTool !== 'function') throw new Error(`${version} 的 dsh-tools 没有导出 defineTool`)
    if (typeof Context !== 'function') throw new Error(`${version} 的 cordis 没有导出 Context`)

    const { apply } = await import(pathToFileURL(path.join(ROOT, 'lib', 'index.js')).href)
    const workspaceRoot = path.join(work, 'ws')
    fs.mkdirSync(workspaceRoot, { recursive: true })
    const config = { workspaceRoot, defineTool }

    // 1) effect 契约 + 注册数量（apply 内部用**那个版本**的 defineTool 定义全部工具）
    const probe = new Map()
    const effect = await apply(
      { tools: { register: (definition) => probe.set(definition.name, definition) }, logger: { info() {} } },
      config
    )
    row.effect_ok = effect === undefined || effect === null || typeof effect === 'function'
    row.registered = probe.size

    // 2) 定义形状：注册出来的定义必须带齐那个版本要求的字段
    //    （0.1.7-rc.2 起 `output.render` 变成必需 —— 只在这里能暴露出来）
    const malformed = []
    for (const [name, definition] of probe) {
      const problems = []
      if (typeof definition.name !== 'string' || definition.name === '') problems.push('name')
      if (typeof definition.description !== 'string') problems.push('description')
      if (definition.parameters === undefined) problems.push('parameters')
      if (definition.output?.schema === undefined) problems.push('output.schema')
      if (typeof definition.output?.render !== 'function') problems.push('output.render')
      if (typeof definition.execute !== 'function') problems.push('execute')
      if (problems.length > 0) malformed.push(`${name}: ${problems.join('+')}`)
    }
    row.defined = probe.size - malformed.length
    row.malformed = malformed.slice(0, 3)

    // 3) 真实 Cordis 装载
    const registered = new Map()
    const ctx = new Context()
    ctx.provide('tools', { register: (definition) => registered.set(definition.name, definition) })
    await ctx.plugin({ apply, name: 'dsh-exp-office', inject: ['tools'] }, config)
    row.loaded = registered.size

    // 4) 真实调用（零依赖、纯本地）+ 用那个版本的渲染契约渲染一次
    const call = async (name, args) => {
      const tool = registered.get(name)
      if (!tool) throw new Error(`工具 ${name} 未注册`)
      return tool.execute(args, { signal: undefined, agent: null })
    }
    const pdf = await call('office_create_pdf', { path: 'matrix-smoke.pdf', lines: ['矩阵冒烟', 'mixed 123'], title: '矩阵' })
    const listed = await call('office_list_files', {})
    const bad = await call('office_read_workbook', { path: 'not-here.xlsx' })
    const okEnvelope = (value) => value && typeof value === 'object' && 'success' in value && 'request_id' in value
    let rendered = null
    try {
      const output = registered.get('office_create_pdf').output.render({ path: 'matrix-smoke.pdf' }, pdf)
      rendered = Array.isArray(output) && output.length > 0
    } catch (err) {
      rendered = false
      row.render_error = err.message
    }
    row.smoke = {
      pdf_ok: pdf.success === true && fs.existsSync(path.join(workspaceRoot, 'matrix-smoke.pdf')),
      list_ok: Array.isArray(listed.data?.files),
      error_ok: bad.success === false && typeof bad.error?.code === 'string',
      envelope_ok: okEnvelope(pdf) && okEnvelope(listed) && okEnvelope(bad),
      rendered_ok: rendered,
      error_code: bad.error?.code ?? null
    }
  } catch (err) {
    row.error = (err.stderr || err.message || String(err)).toString().split('\n').filter(Boolean).pop().slice(0, 300)
  } finally {
    if (!keep) fs.rmSync(work, { recursive: true, force: true })
  }
  return row
}

/**
 * 判断一行是否通过。
 * @param {object} row - 结果行。
 * @returns {boolean} 是否通过。
 */
function passed(row) {
  return (
    row.error === null &&
    row.defined === declaredTools &&
    (row.malformed ?? []).length === 0 &&
    row.effect_ok === true &&
    row.registered === declaredTools &&
    row.loaded === declaredTools &&
    row.smoke?.pdf_ok === true &&
    row.smoke?.list_ok === true &&
    row.smoke?.error_ok === true &&
    row.smoke?.envelope_ok === true &&
    row.smoke?.rendered_ok === true
  )
}

const rows = []
for (const version of targets) {
  if (!asJson) console.log(`→ 正在验证 dsh-tools@${version} …`)
  rows.push(await checkVersion(version))
}

const installed = rows.filter((r) => r.tools !== null)
const failed = rows.filter((r) => r.tools !== null && !passed(r))

if (asJson) {
  console.log(JSON.stringify({ declared_tools: declaredTools, rows, passed: rows.filter(passed).length, failed: failed.length }, null, 2))
} else {
  console.log('\n=== Harness 版本兼容矩阵 ===')
  console.log(`插件声明: ${declaredTools} 个工具`)
  console.log('')
  const header = ['dsh-tools', 'cordis', '定义形状', 'effect', '装载', '调用+渲染', '结论']
  const table = rows.map((r) => [
    r.version,
    r.cordis ?? '-',
    r.defined === null ? '-' : `${r.defined === declaredTools ? '✓' : '✗'} ${r.defined}`,
    r.effect_ok === null ? '-' : r.effect_ok ? '✓' : '✗',
    r.loaded === null ? '-' : String(r.loaded),
    r.smoke === null ? '-' : r.smoke.pdf_ok && r.smoke.list_ok && r.smoke.error_ok && r.smoke.rendered_ok ? '✓' : '✗',
    r.tools === null ? `SKIP（${r.error ?? '未安装'}）` : passed(r) ? '✅ 兼容' : '❌ 不兼容'
  ])
  const widths = header.map((h, i) => Math.max(...[h, ...table.map((row) => row[i])].map((cell) => [...cell].length)))
  const line = (cells) => cells.map((cell, i) => cell + ' '.repeat(Math.max(0, widths[i] - [...cell].length))).join('  ')
  console.log(line(header))
  console.log(widths.map((w) => '-'.repeat(w)).join('  '))
  for (const row of table) console.log(line(row))
  console.log('')
  for (const row of rows) {
    if (row.error && row.tools === null) console.log(`  [skip] ${row.version}: ${row.error}`)
    else if (!passed(row)) console.log(`  [fail] ${row.version}: ${row.error ?? JSON.stringify(row.smoke)}`)
  }
  console.log(`\n结论: ${rows.filter(passed).length} 个版本通过，${failed.length} 个失败，${rows.length - installed.length} 个跳过`)
}

if (installed.length === 0) process.exit(2)
process.exit(failed.length === 0 ? 0 : 3)
