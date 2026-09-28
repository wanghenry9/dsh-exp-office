/**
 * Harness **CLI 级**兼容矩阵：用每个 DSH 版本**自己的 CLI**走一遍真实用户的路径 ——
 * 建 profile → 从 GitHub 装插件 → 组合配置树，确认那个版本的加载器接受我们的清单与补丁层。
 *
 * 与 `test/harness-matrix.mjs` 的分工：
 *   - harness-matrix.mjs（快，几十秒）：验证那个版本的 `dsh-tools` / `cordis` 能定义并运行
 *     我们的 90 个工具（schema DSL + 调用 + 渲染契约）；
 *   - 本脚本（慢，每个版本约 1.5–2 分钟）：验证那个版本的 **CLI 与加载器**接受
 *     `dsh.plugin.json` / `cordis.patch.yml`，并真的把插件装进 profile。
 *
 * 全程隔离：每个版本用**自己的临时 DSH_HOME**，绝不碰用户正在用的 `~/.dsh`。
 * 装好的 DSH 应用缓存在系统临时目录（`dsh-matrix-cache-<版本>`），重跑不必再下载。
 *
 * 用法：
 *   node test/harness-cli-matrix.mjs                        # 默认：最早可装 + 宿主 + 最新
 *   node test/harness-cli-matrix.mjs 0.1.7-rc.2             # 指定版本
 *   node test/harness-cli-matrix.mjs --source file          # 用本地工作区代替 GitHub
 *   node test/harness-cli-matrix.mjs --json                  # 机器可读
 *
 * 退出码：0 = 全部通过；2 = 一个版本都没跑成（通常是没网）；3 = 有版本不兼容。
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { execFileSync } from 'node:child_process'

const ROOT = path.resolve(import.meta.dirname, '..')
const argv = process.argv.slice(2)
const asJson = argv.includes('--json')
const keep = argv.includes('--keep')
// 先把「带值的选项」摘出来，剩下的位置参数才是版本号（否则 --source 的取值会被当成版本）
const positional = []
let source = 'github'
for (let i = 0; i < argv.length; i += 1) {
  if (argv[i] === '--source') {
    source = argv[i + 1] ?? 'github'
    i += 1
    continue
  }
  if (argv[i] === '--install-timeout' || argv[i] === '--cli-timeout') { i += 1; continue }
  if (argv[i].startsWith('--')) continue
  positional.push(argv[i])
}
const DEFAULT_VERSIONS = ['0.1.6-alpha.2', '0.1.7-rc.2']
const targets = positional.length > 0 ? positional : DEFAULT_VERSIONS
// 旧版本的依赖树可能极慢（实测 dsh@0.1.1-rc.1 装 20 分钟无产出），
// 需要验证更早版本时显式传版本号，并可用 --install-timeout 放宽。
const installTimeoutMs = Number(
  argv.includes('--install-timeout') ? argv[argv.indexOf('--install-timeout') + 1] : 300000
)
const cliTimeoutMs = Number(argv.includes('--cli-timeout') ? argv[argv.indexOf('--cli-timeout') + 1] : 300000)

const packageJson = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'))

/**
 * 推导安装源：默认用 GitHub 仓库地址（与用户真实安装方式一致）。
 * @returns {string} pnpm 可识别的安装规格。
 */
function resolveSpec() {
  if (source === 'file') return ROOT
  if (source === 'npm') return `${packageJson.name}@${packageJson.version}`
  const url = packageJson.repository?.url ?? packageJson.homepage ?? ''
  const match = /github\.com[/:]([^/]+\/[^/]+?)(?:\.git)?$/.exec(url)
  if (!match) {
    throw new Error(`无法从 package.json 推导 GitHub 仓库地址（repository=${JSON.stringify(packageJson.repository)}）：请先补 repository，或用 --source file`)
  }
  return `github:${match[1]}`
}

const spec = resolveSpec()

const npmCli = path.join(path.dirname(process.execPath), 'node_modules', 'npm', 'bin', 'npm-cli.js')

/**
 * 干净环境（外层 npm run 注入的 npm_config_* 会改变子进程行为；DSH_HOME 每次显式指定）。
 * @param {object} [extra] - 追加的环境变量。
 * @returns {NodeJS.ProcessEnv} 环境。
 */
function cleanEnv(extra = {}) {
  const env = {}
  for (const [key, value] of Object.entries(process.env)) {
    if (key.toLowerCase().startsWith('npm_config_')) continue
    if (key === 'DSH_HOME') continue
    env[key] = value
  }
  return { ...env, ...extra }
}

/**
 * 跑命令并返回结果。
 *
 * 为什么把标准输出重定向到**文件**而不是管道：npm 会留下持有管道写端的孙进程，
 * 用管道时 `execFileSync` 会一直等 EOF —— 实测在装 `dsh@0.1.1-rc.1`（旧版本依赖树异常）
 * 时卡了 20 分钟不动。写文件就没有这个等待，配合超时能稳定收敛。
 *
 * @param {string} command - 可执行文件。
 * @param {string[]} args - 参数。
 * @param {object} [options] - 选项（cwd / env / timeout）。
 * @returns {{ok: boolean, output: string, status: number, timedOut: boolean}} 结果。
 */
let logSeq = 0
function run(command, args, options = {}) {
  const logFile = path.join(os.tmpdir(), `dsh-matrix-log-${process.pid}-${(logSeq += 1)}.txt`)
  const fd = fs.openSync(logFile, 'w')
  try {
    execFileSync(command, args, {
      cwd: options.cwd ?? ROOT,
      stdio: ['ignore', fd, fd],
      env: cleanEnv(options.env),
      timeout: options.timeout ?? 300000,
      killSignal: 'SIGKILL'
    })
    return { ok: true, output: fs.readFileSync(logFile, 'utf8'), status: 0, timedOut: false }
  } catch (err) {
    const output = fs.existsSync(logFile) ? fs.readFileSync(logFile, 'utf8') : ''
    return { ok: false, output, status: err.status ?? 1, timedOut: err.signal === 'SIGKILL' }
  } finally {
    fs.closeSync(fd)
    fs.rmSync(logFile, { force: true })
  }
}

/**
 * 确保某个版本的 DSH 应用已安装到缓存目录，返回其 CLI 入口。
 * @param {string} version - DSH 版本。
 * @returns {{ok: boolean, bin: string|null, note: string}} 结果。
 */
function ensureApp(version) {
  const cache = path.join(os.tmpdir(), `dsh-matrix-cache-${version}`)
  const bin = path.join(cache, 'node_modules', '@deepseek-ai', 'dsh', 'lib', 'bin.js')
  if (fs.existsSync(bin)) return { ok: true, bin, note: '缓存命中' }
  fs.mkdirSync(cache, { recursive: true })
  fs.writeFileSync(path.join(cache, 'package.json'), JSON.stringify({ name: 'matrix-cache', private: true, version: '0.0.0' }, null, 2))
  const install = run(process.execPath, [npmCli, 'install', '--no-audit', '--no-fund', '--ignore-scripts', `@deepseek-ai/dsh@${version}`], {
    cwd: cache,
    timeout: installTimeoutMs
  })
  if (!fs.existsSync(bin)) {
    return {
      ok: false,
      bin: null,
      note: install.timedOut
        ? `安装超时（>${Math.round(installTimeoutMs / 1000)}s，旧版本依赖树可能异常慢；可用 --install-timeout 放宽）`
        : `安装失败：${install.output.split('\n').filter(Boolean).pop()?.slice(0, 200) ?? '未知'}` 
    }
  }
  return { ok: true, bin, note: '本次下载安装' }
}

const rows = []
for (const version of targets) {
  if (!asJson) console.log(`→ 正在验证 DSH ${version} …`)
  const row = { version, cli_version: null, installed: null, dumped: null, bundle_in_tree: null, error: null, app: null }
  const app = ensureApp(version)
  row.app = app.note
  if (!app.ok) {
    row.error = app.note
    rows.push(row)
    continue
  }
  const home = fs.mkdtempSync(path.join(os.tmpdir(), `dsh-home-cli-${version}-`))
  try {
    const cliVersion = run(process.execPath, [app.bin, '--version'], { env: { DSH_HOME: home }, timeout: cliTimeoutMs })
    row.cli_version = cliVersion.output.trim().split('\n').pop() ?? null

    const add = run(process.execPath, [app.bin, 'plugin', '--profile', 'matrix', 'add', spec], { env: { DSH_HOME: home }, timeout: cliTimeoutMs })
    if (!add.ok) throw new Error(`plugin add 失败：${add.output.split('\n').filter(Boolean).pop()}`)
    // 装完必须真的落到 profile 的依赖里
    const profilePkg = path.join(home, 'profiles', 'matrix', 'package.json')
    const profile = JSON.parse(fs.readFileSync(profilePkg, 'utf8'))
    row.installed = Boolean(profile.dependencies?.[packageJson.name])
    row.bundles = Array.isArray(profile.dsh?.profile?.bundles) ? profile.dsh.profile.bundles.includes(packageJson.name) : false

    const dump = run(process.execPath, [app.bin, '--profile', 'matrix', '--dump-config'], { env: { DSH_HOME: home }, timeout: cliTimeoutMs })
    if (!dump.ok) throw new Error(`--dump-config 失败：${dump.output.split('\n').filter(Boolean).pop()}`)
    row.dumped = true
    // 组合树里必须出现我们的插件层（`# == <name>` 是 dump-config 的分层标题）
    row.bundle_in_tree = dump.output.includes(`# == ${packageJson.name}`) && dump.output.includes(`- id: ${packageJson.name}`)
    row.tree_excerpt = dump.output
      .split('\n')
      .filter((line) => line.includes(packageJson.name))
      .slice(0, 3)
      .map((line) => line.trim())
  } catch (err) {
    row.error = (err.message ?? String(err)).split('\n')[0].slice(0, 300)
  } finally {
    if (!keep) fs.rmSync(home, { recursive: true, force: true })
  }
  rows.push(row)
}

/**
 * 判定一行是否通过。
 * @param {object} row - 结果行。
 * @returns {boolean} 是否通过。
 */
const passed = (row) =>
  row.error === null && row.cli_version === row.version && row.installed === true && row.bundles === true && row.bundle_in_tree === true

const attempted = rows.filter((r) => r.dumped !== null)
const failed = attempted.filter((r) => !passed(r))

if (asJson) {
  console.log(JSON.stringify({ spec, rows, passed: rows.filter(passed).length, failed: failed.length }, null, 2))
} else {
  console.log('\n=== Harness CLI 级兼容矩阵 ===')
  console.log(`安装源: ${spec}`)
  console.log('')
  const header = ['DSH 版本', 'CLI 报版本', '装进 profile', '注册为 bundle', '组合树含插件层', '结论']
  const table = rows.map((r) => [
    r.version,
    r.cli_version ?? '-',
    r.installed === null ? '-' : r.installed ? '✓' : '✗',
    r.bundles === null ? '-' : r.bundles ? '✓' : '✗',
    r.bundle_in_tree === null ? '-' : r.bundle_in_tree ? '✓' : '✗',
    r.dumped === null ? `SKIP（${r.error ?? '未跑成'}）` : passed(r) ? '✅ 兼容' : '❌ 不兼容'
  ])
  const widths = header.map((h, i) => Math.max(...[h, ...table.map((row) => row[i])].map((cell) => [...cell].length)))
  const line = (cells) => cells.map((cell, i) => cell + ' '.repeat(Math.max(0, widths[i] - [...cell].length))).join('  ')
  console.log(line(header))
  console.log(widths.map((w) => '-'.repeat(w)).join('  '))
  for (const row of table) console.log(line(row))
  console.log('')
  for (const row of rows) {
    if (row.error) console.log(`  [${row.dumped === null ? 'skip' : 'fail'}] ${row.version}: ${row.error}`)
    if (passed(row) && row.tree_excerpt) console.log(`  [tree] ${row.version}: ${row.tree_excerpt.join(' | ')}`)
  }
  console.log(`\n结论: ${rows.filter(passed).length} 个版本通过，${failed.length} 个失败，${rows.length - attempted.length} 个跳过`)
}

if (attempted.length === 0) process.exit(2)
process.exit(failed.length === 0 ? 0 : 3)
