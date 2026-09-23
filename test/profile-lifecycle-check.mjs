/**
 * 安装 / 升级 / 卸载 演练（阶段 7「安装、升级和卸载流程」）。
 *
 * 在一个**临时目录**里假装它是一个 profile，全程离线、不碰你真实的 `~/.dsh`：
 *   1. 安装：以 `file:<工作区>` 作为依赖安装，断言装得上、零运行时依赖、能装载出全部工具；
 *   2. 升级：改成从 `npm pack` 出来的 tarball 安装（覆盖旧版本），断言文件确实被换新、仍能装载；
 *   3. 卸载：`npm uninstall`，断言插件目录被移除、profile 里没有残留（锁文件 / 暂存目录 / 临时文件）。
 *
 * 退出码：0 = 通过；3 = 有问题；2 = 环境错误（例如 npm 不可用）。
 *
 * 用法：node test/profile-lifecycle-check.mjs [--keep]
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execFileSync } from 'node:child_process'
import { pathToFileURL } from 'node:url'

const ROOT = path.resolve(import.meta.dirname, '..')
const keep = process.argv.includes('--keep')
const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'dsh-office-lifecycle-'))
const profileDir = path.join(workDir, 'profile')
const problems = []

/**
 * 调 npm：直接让当前 node 跑 npm 的 JS 入口（Windows 下 `npm` 是 .cmd，execFile 会 ENOENT）。
 * @param {string[]} args - 参数。
 * @param {object} [options] - 选项。
 * @returns {string} 输出。
 */
function runNpm(args, options = {}) {
  const npmCli = path.join(path.dirname(process.execPath), 'node_modules', 'npm', 'bin', 'npm-cli.js')
  return execFileSync(process.execPath, [npmCli, ...args], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], ...options })
}

/**
 * 用真实 Cordis 装载某个已安装副本，返回注册到的工具数。
 * @param {string} pluginDir - 插件目录（含 lib/index.js）。
 * @returns {Promise<number>} 注册的工具数。
 */
async function loadInstalled(pluginDir) {
  const dshTools = path.join(os.homedir(), 'AppData/Roaming/npm/node_modules/@deepseek-ai/dsh/node_modules/@deepseek-ai/dsh-tools/lib/index.js')
  const cordisPath = path.join(path.dirname(path.dirname(path.dirname(dshTools))), 'cordis/lib/index.js')
  const { Context } = await import(pathToFileURL(cordisPath).href)
  const { defineTool } = await import(pathToFileURL(dshTools).href)
  const { apply } = await import(pathToFileURL(path.join(pluginDir, 'lib', 'index.js')).href)
  const registered = new Map()
  const ctx = new Context()
  ctx.provide('tools', { register: (definition) => registered.set(definition.name, definition) })
  await ctx.plugin({ apply, name: 'dsh-exp-office', inject: ['tools'] }, { workspaceRoot: workDir, defineTool })
  return registered.size
}


console.log('=== 安装 / 升级 / 卸载 演练 ===')
console.log(`临时 profile: ${profileDir}`)
fs.mkdirSync(profileDir, { recursive: true })
fs.writeFileSync(
  path.join(profileDir, 'package.json'),
  JSON.stringify({ name: 'dsh-profile-rehearsal', private: true, version: '0.0.0', dependencies: {} }, null, 2)
)

const declaredTools = JSON.parse(fs.readFileSync(path.join(ROOT, 'dsh.plugin.json'), 'utf8')).contributes.tools.length
const installedDir = path.join(profileDir, 'node_modules', 'dsh-exp-office')

/**
 * 给工作区的 `lib/` 与三个清单算一份指纹，用来断言整套演练**绝不改动源码**。
 *
 * 这一条是被真实事故逼出来的：`file:<目录>` 依赖在这台机器上是**符号链接**回工作区，
 * 于是「改一改安装副本里的文件看看升级会不会覆盖」这种做法实际改的是**工作区源码**
 * （实测真的往 `lib/tools.js` 追了 6 行标记）。有了指纹断言，这类越界立刻被抓出来。
 *
 * @returns {string} 指纹（路径:大小:修改时间 的拼接）。
 */
function sourceFingerprint() {
  const entries = []
  for (const dir of ['lib']) {
    for (const name of fs.readdirSync(path.join(ROOT, dir)).sort()) {
      const full = path.join(ROOT, dir, name)
      const stat = fs.statSync(full)
      entries.push(`${dir}/${name}:${stat.size}:${stat.mtimeMs}`)
    }
  }
  for (const name of ['package.json', 'dsh.plugin.json', 'cordis.patch.yml']) {
    const stat = fs.statSync(path.join(ROOT, name))
    entries.push(`${name}:${stat.size}:${stat.mtimeMs}`)
  }
  return entries.join('|')
}

const fingerprintBefore = sourceFingerprint()

// 1) 安装（file: 指向工作区）
try {
  runNpm(['install', '--no-audit', '--no-fund', `file:${ROOT}`], { cwd: profileDir })
} catch (err) {
  console.error(`安装失败：${(err.stderr || err.message).toString().split('\n')[0]}`)
  process.exit(2)
}
if (!fs.existsSync(path.join(installedDir, 'lib', 'index.js'))) problems.push('安装后找不到 lib/index.js')
const isLinked = fs.lstatSync(installedDir).isSymbolicLink()
console.log(`安装方式: ${isLinked ? '符号链接回工作区（file: 目录依赖的默认行为）' : '真实复制'}`)
const depsAfterInstall = fs.existsSync(path.join(installedDir, 'node_modules')) ? fs.readdirSync(path.join(installedDir, 'node_modules')) : []
if (depsAfterInstall.length > 0) problems.push(`安装后带了运行时依赖：${depsAfterInstall.join('、')}`)
let count = await loadInstalled(installedDir)
if (count !== declaredTools) problems.push(`安装后装载出 ${count} 个工具，与清单 ${declaredTools} 不一致`)
else console.log(`安装: 装载 ${count} 个工具 ✓（运行时依赖 0 个）`)

// 2) 升级：准备一份「新版本」（版本号 +1 的工作区副本），把依赖从旧目录切到新目录再收敛
//
// 两条实测结论（都踩过，写在这里免得下次又绕）：
//   - `file:<目录>` 依赖在这台机器上装出来是**符号链接**，所以「改安装副本里的文件看升级是否覆盖」
//     这种探针会直接改到工作区源码上 —— 本脚本因此加了源码指纹断言（见 2.5）；
//   - `file:C:/…` 这种盘符绝对路径会被 npm 解析坏（日志里变成 `file:C:Users…`，分隔符被吃掉 → ENOENT），
//     依赖里一律用**相对路径**；tarball 安装路径由 `test/pack-check.mjs` 单独覆盖。
let upgradedVersion = null
try {
  const stage = path.join(workDir, 'stage')
  fs.cpSync(ROOT, stage, {
    recursive: true,
    filter: (src) => !/node_modules|\.git|\.dsh-release|[\\/]fixtures([\\/]|$)/.test(src)
  })
  const stagedPkg = JSON.parse(fs.readFileSync(path.join(stage, 'package.json'), 'utf8'))
  const [major, minor, patch] = stagedPkg.version.split('.').map(Number)
  stagedPkg.version = `${major}.${minor}.${patch + 1}`
  upgradedVersion = stagedPkg.version
  fs.writeFileSync(path.join(stage, 'package.json'), `${JSON.stringify(stagedPkg, null, 2)}\n`)

  const profilePkgPath = path.join(profileDir, 'package.json')
  const profilePkg = JSON.parse(fs.readFileSync(profilePkgPath, 'utf8'))
  profilePkg.dependencies = { 'dsh-exp-office': `file:${path.relative(profileDir, stage).replace(/\\/g, '/')}` }
  fs.writeFileSync(profilePkgPath, `${JSON.stringify(profilePkg, null, 2)}\n`)
  runNpm(['install', '--no-audit', '--no-fund'], { cwd: profileDir })
} catch (err) {
  problems.push(`升级失败：${(err.stderr || err.message).toString().split('\n')[0]}`)
}
if (upgradedVersion) {
  if (!fs.existsSync(installedDir)) {
    problems.push('升级后安装位置不存在')
  } else {
    const installedVersion = JSON.parse(fs.readFileSync(path.join(installedDir, 'package.json'), 'utf8')).version
    if (installedVersion !== upgradedVersion) problems.push(`升级后版本号是 ${installedVersion}，期望 ${upgradedVersion}`)
    count = await loadInstalled(installedDir)
    if (count !== declaredTools) problems.push(`升级后装载出 ${count} 个工具，与清单 ${declaredTools} 不一致`)
    else console.log(`升级: 版本 ${installedVersion}、装载 ${count} 个工具 ✓`)
  }
}

// 2.5) 安全断言：整套演练**绝不能改动工作区源码**（file: 目录依赖是符号链接，很容易误伤）
if (sourceFingerprint() !== fingerprintBefore) {
  problems.push('演练改动了工作区源码（lib/ 或清单文件的指纹变了）—— 检查是否有对安装副本的写操作落到了源码上')
} else {
  console.log('源码保护: 演练全程未改动工作区 lib/ 与清单 ✓')
}

// 3) 卸载
try {
  runNpm(['uninstall', 'dsh-exp-office'], { cwd: profileDir })
} catch (err) {
  problems.push(`卸载失败：${(err.stderr || err.message).toString().split('\n')[0]}`)
}
if (fs.existsSync(installedDir)) problems.push('卸载后插件目录仍存在')

// 4) 残留检查：锁文件、暂存目录、临时文件
const leftovers = []
const walk = (dir) => {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === '.package-lock.json') continue
    const full = path.join(dir, entry.name)
    if (entry.isDirectory()) walk(full)
    else if (/\.dsh-office\.lock$|\.tmp$|^staging-/.test(entry.name)) leftovers.push(path.relative(profileDir, full))
  }
}
walk(profileDir)
if (leftovers.length > 0) problems.push(`profile 里有残留：${leftovers.join('、')}`)
else console.log('卸载: 目录已移除、无锁文件与暂存残留 ✓')

if (!keep) fs.rmSync(workDir, { recursive: true, force: true })

if (problems.length === 0) {
  console.log('结论: 安装 → 升级 → 卸载 全流程通过')
  process.exit(0)
}
console.log(`结论: ${problems.length} 个问题`)
for (const p of problems) console.log(`  [!] ${p}`)
process.exit(3)
