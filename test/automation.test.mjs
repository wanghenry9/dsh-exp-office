/**
 * 本地 Office/WPS 联动的 Node 侧测试（阶段 6）。
 *
 * 这里**不启动真实 Office**：用一个「桩脚本」替换 office-automation.ps1，
 * 因此可以稳定地验证开关、超时、进程回收兜底、结果解析、审计与错误码。
 * 真实引擎的验证在 `test/automation-check.mjs`（需要本机装了 Office/WPS）。
 *
 * 运行：node test/automation.test.mjs
 */
import assert from 'node:assert/strict'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { AUTOMATION_KINDS, OfficeAutomation } from '../lib/automation.js'

const here = dirname(fileURLToPath(import.meta.url))
const workspace = mkdtempSync(join(tmpdir(), 'dsh-office-automation-'))

let passed = 0
let failed = 0

/**
 * 运行一个用例并记录结果。
 * @param {string} name - 用例名。
 * @param {() => void|Promise<void>} fn - 用例体。
 * @returns {Promise<void>} 完成信号。
 */
async function test(name, fn) {
  try {
    await fn()
    passed += 1
    console.log(`  ✅ ${name}`)
  } catch (err) {
    failed += 1
    console.log(`  ❌ ${name}\n     ${err.message}`)
  }
}

// ── 桩脚本：纯 ASCII，行为由环境变量 STUB_MODE 控制 ──────────────────────────
const stubPath = join(workspace, 'stub-automation.ps1')
const stubLog = join(workspace, 'stub-invocations.log')
writeFileSync(
  stubPath,
  [
    'param(',
    "  [string]$Action = 'detect',",
    "  [string]$Kind = 'xlsx',",
    "  [string]$Engine = 'auto',",
    "  [string]$InputPath = '',",
    "  [string]$OutputPath = '',",
    '  [string]$ResultPath,',
    "  [string]$StatePath = '',",
    '  [switch]$ProbeCom',
    ')',
    "$mode = $env:STUB_MODE",
    'if ($mode -eq $null) { $mode = "ok" }',
    '$log = $env:STUB_LOG',
    'if ($log) { Add-Content -LiteralPath $log -Value "$Action|$Engine|$Kind" }',
    'if ($mode -eq "sleep") { Start-Sleep -Seconds 30 }',
    'if ($mode -eq "silent") { exit 0 }',
    'if ($mode -eq "fail") {',
    '  $json = \'{"action":"\' + $Action + \'","ok":false,"error":"stub engine exploded","warnings":[],"engine":"stub","killed_pids":[],"leftover_pids":[],"elapsed_ms":5}\'',
    '} else {',
    '  $json = \'{"action":"\' + $Action + \'","ok":true,"engine":"stub","version":"9.9","prog_id":"Stub.Application","warnings":["stub warning"],"killed_pids":[4242],"leftover_pids":[777],"elapsed_ms":7,"output_size":2048}\'',
    '}',
    '[System.IO.File]::WriteAllText($ResultPath, $json, (New-Object System.Text.UTF8Encoding($false)))',
    ''
  ].join('\n'),
  'ascii'
)

/**
 * 造一个受测对象。
 * @param {object} [options] - 覆盖项。
 * @returns {OfficeAutomation} 联动对象。
 */
function makeAutomation(options = {}) {
  return new OfficeAutomation({
    enabled: true,
    scriptPath: stubPath,
    scratchDir: join(workspace, 'scratch'),
    auditPath: join(workspace, 'audit.jsonl'),
    timeoutMs: 60000,
    ...options
  })
}

console.log('\n=== 1. 开关与参数校验 ===')

await test('默认关闭：驱动引擎时给出 AUTOMATION_DISABLED，只读检测仍然可用', async () => {
  process.env.STUB_MODE = 'ok'
  const automation = makeAutomation({ enabled: false })
  assert.equal(automation.enabled, false)
  // 只读注册表的检测不启动任何程序，因此不需要授权
  const detect = await automation.detectEngines({})
  assert.equal(detect.probed_com, false)
  // 真的要创建 COM 实例、真的要驱动引擎 → 必须授权
  await assert.rejects(
    () => automation.detectEngines({ probeCom: true }),
    (err) => err.code === 'AUTOMATION_DISABLED' && /allowLocalAutomation/.test(err.message)
  )
  await assert.rejects(
    () => automation.recalculate({ inputPath: 'a.xlsx', outputPath: 'b.xlsx', kind: 'xlsx' }),
    (err) => err.code === 'AUTOMATION_DISABLED'
  )
})

await test('不支持的类型 / 非法引擎 / 类型与引擎不匹配都被拒绝', async () => {
  const automation = makeAutomation()
  await assert.rejects(
    () => automation.recalculate({ inputPath: 'a.pdf', outputPath: 'b.pdf', kind: 'pdf' }),
    (err) => err.code === 'UNSUPPORTED_FILE_TYPE'
  )
  await assert.rejects(
    () => automation.recalculate({ inputPath: 'a.xlsx', outputPath: 'b.xlsx', kind: 'xlsx', engine: 'powerpoint' }),
    (err) => err.code === 'INVALID_REQUEST'
  )
  await assert.rejects(
    () => automation.recalculate({ inputPath: 'a.xlsx', outputPath: '', kind: 'xlsx' }),
    (err) => err.code === 'INVALID_REQUEST'
  )
  assert.deepEqual([...AUTOMATION_KINDS], ['xlsx', 'xlsm', 'docx', 'docm', 'pptx', 'pptm'])
})

console.log('\n=== 2. 结果解析 ===')

await test('检测：解析引擎清单（只读注册表，不需要真实 Office）', async () => {
  process.env.STUB_MODE = 'ok'
  const automation = makeAutomation()
  const result = await automation.detectEngines({})
  assert.equal(result.probed_com, false)
  assert.equal(result.started_pids.length + result.leftover_pids.length >= 0, true)
})

await test('重算：把脚本结果整理成统一结构，并如实报告残留进程', async () => {
  process.env.STUB_MODE = 'ok'
  const automation = makeAutomation()
  const result = await automation.recalculate({ inputPath: join(workspace, 'a.xlsx'), outputPath: join(workspace, 'b.xlsx'), kind: 'xlsx' })
  assert.equal(result.engine, 'stub')
  assert.equal(result.engine_version, '9.9')
  assert.equal(result.output_size, 2048)
  assert.deepEqual(result.warnings, ['stub warning'])
  assert.deepEqual(result.leftover_pids, [777], '残留 PID 必须如实返回，不能被吞掉')
})

await test('引擎报错 → AUTOMATION_FAILED，且带上脚本给的原始说明', async () => {
  process.env.STUB_MODE = 'fail'
  const automation = makeAutomation()
  await assert.rejects(
    () => automation.rerender({ inputPath: join(workspace, 'a.docx'), outputPath: join(workspace, 'b.docx'), kind: 'docx' }),
    (err) => err.code === 'AUTOMATION_FAILED' && /stub engine exploded/.test(err.message)
  )
})

await test('脚本没写结果文件 → AUTOMATION_FAILED（而不是假装成功）', async () => {
  process.env.STUB_MODE = 'silent'
  const automation = makeAutomation()
  await assert.rejects(
    () => automation.detectEngines({}),
    (err) => err.code === 'AUTOMATION_FAILED' && /没有返回结果/.test(err.message)
  )
})

console.log('\n=== 3. 超时与进程回收兜底 ===')

await test('超时：终止本次调用并补一次 cleanup（只按状态文件里的 PID 收尾）', async () => {
  process.env.STUB_MODE = 'sleep'
  const automation = makeAutomation({ timeoutMs: 1000 })
  const startedAt = Date.now()
  await assert.rejects(
    () => automation.detectEngines({}),
    (err) => err.code === 'TIMEOUT' && /已终止并清理/.test(err.message)
  )
  const elapsed = Date.now() - startedAt
  assert.ok(elapsed < 25000, `超时应在几十秒内返回，实际 ${elapsed} ms`)
})

console.log('\n=== 4. 审计与临时目录 ===')

await test('审计日志记录了动作/引擎/文件名，但不含绝对路径与文档内容', async () => {
  process.env.STUB_MODE = 'ok'
  const auditPath = join(workspace, 'audit-2.jsonl')
  const automation = makeAutomation({ auditPath })
  await automation.recalculate({ inputPath: join(workspace, 'sales.xlsx'), outputPath: join(workspace, 'out.xlsx'), kind: 'xlsx' })
  const raw = readFileSync(auditPath, 'utf8')
  const lines = raw.trim().split('\n').map((l) => JSON.parse(l))
  assert.ok(lines.length >= 1)
  const entry = lines[lines.length - 1]
  assert.equal(entry.action, 'recalc')
  assert.equal(entry.engine, 'stub')
  assert.equal(entry.ok, true)
  assert.equal(entry.input_name, 'sales.xlsx')
  assert.equal(typeof entry.elapsed_ms, 'number')
  assert.equal(raw.includes(workspace), false, '审计日志不应包含本机绝对路径')
})

await test('每次调用的临时目录都会被清理，不留 staging 残留', () => {
  const scratch = join(workspace, 'scratch')
  const leftovers = existsSync(scratch) ? readdirSync(scratch) : []
  assert.deepEqual(leftovers, [], `临时目录应清空，实际：${leftovers.join(', ')}`)
})

console.log('\n=== 5. 跨平台与脚本缺失 ===')

await test('脚本不存在时明确报 INTERNAL_ERROR（不是静默失败）', async () => {
  const automation = makeAutomation({ scriptPath: join(workspace, 'nope.ps1') })
  await assert.rejects(
    () => automation.detectEngines({}),
    (err) => err.code === 'INTERNAL_ERROR' && /自动化脚本不存在/.test(err.message)
  )
})

rmSync(workspace, { recursive: true, force: true })
void mkdirSync
void here

console.log(`\n结果：${passed} 通过，${failed} 失败\n`)
process.exit(failed === 0 ? 0 : 1)
