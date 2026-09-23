/**
 * 发布包自检：`npm pack` 出来的 tarball 必须「装得上、装得全、能装载」。
 *
 * 做三件事：
 *   1. 打包（`npm pack --json`），核对文件清单 —— `package.json` 的 `files` 白名单是否覆盖了
 *      运行必需的部件（lib/ 全部模块、dsh.plugin.json、cordis.patch.yml、README.md），
 *      以及有没有把不该进包的东西（.dsh-release、node_modules、测试夹具、任务清单）带进去；
 *   2. 在临时目录里真的 `npm install <tarball>`，确认安装成功且没有把 devDependency 拖进来；
 *   3. 用**安装后的那份代码**跑真实 Cordis 装载，断言注册的工具数与 `dsh.plugin.json` 声明一致
 *      —— 这一步等价于「用户装完能不能用」。
 *
 * 退出码：0 = 通过；3 = 有问题；2 = 用法/环境错误。
 *
 * 用法：node test/pack-check.mjs [--keep]（--keep 保留临时目录便于排查）
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execFileSync } from 'node:child_process'
import { pathToFileURL } from 'node:url'

const ROOT = path.resolve(import.meta.dirname, '..')
const keep = process.argv.includes('--keep')
const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'dsh-office-pack-'))

const run = (cmd, args, options = {}) =>
  execFileSync(cmd, args, { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], ...options })

/**
 * 调 npm 的稳妥方式：直接让当前 node 跑 npm 的 JS 入口。
 *
 * Windows 上 `npm` 是 `npm.cmd`/`npm.ps1`，`execFileSync('npm', …)` 会 ENOENT，
 * 而 `shell: true` 又要自己处理引号；直接用 `node <npm-cli.js>` 两种平台都干净。
 *
 * @param {string[]} args - npm 参数。
 * @param {object} [options] - execFileSync 选项。
 * @returns {string} 标准输出。
 */
function runNpm(args, options = {}) {
  const npmCli = path.join(path.dirname(process.execPath), 'node_modules', 'npm', 'bin', 'npm-cli.js')
  if (fs.existsSync(npmCli)) return run(process.execPath, [npmCli, ...args], options)
  return run(process.platform === 'win32' ? 'npm.cmd' : 'npm', args, { ...options, shell: process.platform === 'win32' })
}

const problems = []
console.log('=== 发布包自检 ===')
console.log(`工作目录: ${workDir}`)

// 1) 打包
let packed
try {
  const out = runNpm(['pack', '--json', '--pack-destination', workDir])
  const meta = JSON.parse(out)
  packed = { file: path.join(workDir, meta[0].filename), files: meta[0].files.map((f) => f.path), size: meta[0].size }
} catch (err) {
  console.error(`打包失败：${err.stderr || err.message}`)
  process.exit(2)
}
console.log(`tarball: ${path.basename(packed.file)}（${packed.size} 字节，${packed.files.length} 个文件）`)

// 2) 文件清单核对
const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'dsh.plugin.json'), 'utf8'))
const declaredTools = manifest.contributes.tools.length
const required = ['package.json', 'dsh.plugin.json', 'cordis.patch.yml', 'README.md', 'lib/index.js', 'lib/tools.js', 'lib/capabilities.js']
for (const file of required) {
  if (!packed.files.includes(file)) problems.push(`缺少必需文件：${file}`)
}
// lib 下的每个模块都必须进包（加新模块时最容易忘记）
for (const file of fs.readdirSync(path.join(ROOT, 'lib'))) {
  if (!file.endsWith('.js')) continue
  if (!packed.files.includes(`lib/${file}`)) problems.push(`lib/${file} 没进包（检查 package.json 的 files 白名单）`)
}
for (const pattern of ['.dsh-release/', 'node_modules/', 'test/', '任务清单.md', 'docs/']) {
  const strays = packed.files.filter((f) => f.startsWith(pattern))
  if (strays.length > 0) problems.push(`包里混入了不该发布的内容：${pattern}（${strays.length} 个）`)
}

// 3) 临时安装
try {
  fs.writeFileSync(path.join(workDir, 'package.json'), JSON.stringify({ name: 'pack-smoke', private: true, version: '0.0.0' }, null, 2))
  runNpm(['install', '--no-audit', '--no-fund', packed.file], { cwd: workDir })
  const installedDir = path.join(workDir, 'node_modules', 'dsh-exp-office')
  if (!fs.existsSync(path.join(installedDir, 'lib', 'index.js'))) problems.push('安装后找不到 lib/index.js')
  const installedModules = fs.existsSync(path.join(installedDir, 'node_modules')) ? fs.readdirSync(path.join(installedDir, 'node_modules')) : []
  if (installedModules.length > 0) problems.push(`安装后带了运行时依赖（应为零依赖）：${installedModules.join('、')}`)
  console.log(`安装成功: ${path.relative(workDir, installedDir)}（运行时依赖 ${installedModules.length} 个）`)

  // 4) 用安装后的代码跑真实 Cordis 装载
  const dshTools = path.join(os.homedir(), 'AppData/Roaming/npm/node_modules/@deepseek-ai/dsh/node_modules/@deepseek-ai/dsh-tools/lib/index.js')
  const cordisPath = path.join(path.dirname(path.dirname(path.dirname(dshTools))), 'cordis/lib/index.js')
  if (!fs.existsSync(dshTools) || !fs.existsSync(cordisPath)) {
    console.log('跳过装载自检：本机找不到 DSH 的 dsh-tools / cordis（非开发机属正常）')
  } else {
    const { Context } = await import(pathToFileURL(cordisPath).href)
    const { defineTool } = await import(pathToFileURL(dshTools).href)
    const { apply } = await import(pathToFileURL(path.join(installedDir, 'lib', 'index.js')).href)

    // (a) effect 契约：必须直接看 `apply()` 的返回值（`ctx.plugin()` 返回的是 fork 对象，不是 effect）
    const probe = new Map()
    const effect = await apply(
      { tools: { register: (definition) => probe.set(definition.name, definition) }, logger: { info() {} } },
      { workspaceRoot: workDir, defineTool }
    )
    if (!(effect === undefined || effect === null || typeof effect === 'function')) {
      problems.push('apply 返回值不是合法的 Cordis effect（undefined/null/disposer）')
    }
    if (probe.size !== declaredTools) {
      problems.push(`直接调用 apply 注册 ${probe.size} 个工具，与清单声明的 ${declaredTools} 个不一致`)
    }

    // (b) 真实 Cordis 装载（走 Context + provide，与 DSH 的加载路径一致）
    const registered = new Map()
    const ctx = new Context()
    ctx.provide('tools', { register: (definition) => registered.set(definition.name, definition) })
    await ctx.plugin({ apply, name: 'dsh-exp-office', inject: ['tools'] }, { workspaceRoot: workDir, defineTool })
    if (registered.size !== declaredTools) {
      problems.push(`装载后注册 ${registered.size} 个工具，与清单声明的 ${declaredTools} 个不一致`)
    } else {
      console.log(`装载自检: 注册 ${registered.size} 个工具，与 dsh.plugin.json 一致（effect 契约通过）`)
    }
  }
} catch (err) {
  problems.push(`安装或装载失败：${(err.stderr || err.message || '').toString().split('\n')[0]}`)
}

if (!keep) fs.rmSync(workDir, { recursive: true, force: true })

if (problems.length === 0) {
  console.log('结论: 发布包自检通过（文件齐全、零依赖、可装载）')
  process.exit(0)
}
console.log(`结论: ${problems.length} 个问题`)
for (const p of problems) console.log(`  [!] ${p}`)
process.exit(3)
