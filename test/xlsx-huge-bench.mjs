/**
 * 100 MB 级大文件的读取基准（阶段 7「大文件」欠账的验收脚本）。
 *
 * 每个场景跑在**独立子进程**里，用 `process.resourceUsage().maxRSS` 取真实峰值
 * （不受 GC 时机影响），再对照口径逐条判定：
 *
 * | 指标 | 改造前 | 目标 |
 * |---|---|---|
 * | 峰值 RSS（178 MB 部件，读任意区域） | 420 MB | ≤ 120 MB |
 * | 读开头 A1:J100 的耗时 | 1050 ms | ≤ 300 ms |
 * | 读尾部 A199500:J199600 的耗时 | 1545 ms | ≤ 1200 ms |
 *
 * 样本不存在时会自动用生成器造一份（2,000,000 个单元格 / 未压缩约 178 MB）。
 *
 * 用法：node test/xlsx-huge-bench.mjs [--rows 200000] [--cols 10] [--json]
 * 退出码：0 全部达标；3 有不达标；2 环境/样本问题。
 */
import { existsSync, readFileSync, statSync } from 'node:fs'
import { spawnSync } from 'node:child_process'
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const asJson = process.argv.includes('--json')
const arg = (name, fallback) => {
  const index = process.argv.indexOf(`--${name}`)
  return index === -1 ? fallback : process.argv[index + 1]
}
const rows = Number(arg('rows', '200000'))
const cols = Number(arg('cols', '10'))
const fixture = join(here, 'fixtures', `large-${rows}x${cols}-wide.xlsx`)

/** 口径：`maxRssMb` 与耗时上限。 */
const BUDGET = {
  rss_mb: 120,
  head_ms: 300,
  tail_ms: 1200
}
/** 改造前的实测值，用于对照展示。 */
const BASELINE = { rss_mb: 420, head_ms: 1050, tail_ms: 1545 }

if (!existsSync(fixture)) {
  console.log(`样本不存在，正在生成：${fixture}`)
  const generated = spawnSync(
    process.execPath,
    [join(here, 'make-xlsx-large-fixture.mjs'), '--rows', String(rows), '--cols', String(cols), '--wide'],
    { stdio: 'inherit' }
  )
  if (generated.status !== 0 || !existsSync(fixture)) {
    console.error('样本生成失败')
    process.exit(2)
  }
}

if (process.argv.includes('--child')) {
  // 子进程：跑一个场景并输出 JSON
  const scenario = process.argv[process.argv.indexOf('--scenario') + 1]
  const { Workbook } = await import('../lib/xlsx.js')
  const bytes = readFileSync(fixture)
  const wb = Workbook.open(bytes)
  const started = Date.now()
  let cells = null
  let mode = null
  let partBytes = null
  if (scenario === 'head') {
    const result = await wb.readRangeAsync('大数据', 'A1:J100')
    cells = result.rows[0].length
    mode = result.read_mode
    partBytes = result.part_bytes
  } else if (scenario === 'tail') {
    const result = await wb.readRangeAsync('大数据', `A${rows - 500}:J${rows - 499}`)
    cells = result.rows[0].length
    mode = result.read_mode
    partBytes = result.part_bytes
  }
  const elapsed = Date.now() - started
  console.log(
    JSON.stringify({
      scenario,
      elapsed_ms: elapsed,
      max_rss_mb: Math.round(process.resourceUsage().maxRSS / 1024),
      cells,
      read_mode: mode,
      part_bytes: partBytes
    })
  )
  process.exit(0)
}

const fileMb = statSync(fixture).size / 1048576
console.log('=== 100 MB 级大文件读取基准 ===')
console.log(`样本：${fixture.replace(here, 'test')}`)
console.log(`规模：${rows} 行 × ${cols} 列 = ${(rows * cols).toLocaleString('en-US')} 个单元格｜磁盘 ${fileMb.toFixed(1)} MB`)

const measurements = {}
for (const scenario of ['head', 'tail']) {
  const result = spawnSync(process.execPath, [import.meta.filename, '--child', '--scenario', scenario, '--rows', String(rows), '--cols', String(cols)], {
    encoding: 'utf8',
    maxBuffer: 16 * 1024 * 1024
  })
  const line = (result.stdout ?? '').trim().split('\n').filter(Boolean).pop() ?? ''
  try {
    measurements[scenario] = JSON.parse(line)
  } catch {
    console.error(`场景 ${scenario} 失败：${(result.stderr ?? '').trim().split('\n').slice(-2).join(' | ')}`)
    process.exit(2)
  }
}

const head = measurements.head
const tail = measurements.tail
const partMb = (head.part_bytes ?? 0) / 1048576

const checks = [
  ['峰值内存 (RSS)', 'rss_mb', Math.max(head.max_rss_mb, tail.max_rss_mb), BUDGET.rss_mb, 'MB', BASELINE.rss_mb],
  ['读开头 (A1:J100)', 'head_ms', head.elapsed_ms, BUDGET.head_ms, 'ms', BASELINE.head_ms],
  ['读尾部', 'tail_ms', tail.elapsed_ms, BUDGET.tail_ms, 'ms', BASELINE.tail_ms]
]

let failed = 0
const lines = []
for (const [label, , actual, limit, unit, before] of checks) {
  const ok = actual <= limit
  if (!ok) failed += 1
  lines.push({ label, actual, limit, unit, before, ok })
}

if (asJson) {
  console.log(JSON.stringify({ fixture, part_bytes: head.part_bytes, part_mb: partMb, measurements, checks: lines, failed }, null, 2))
} else {
  console.log(`部件未压缩：${partMb.toFixed(1)} MB｜读取路径：${head.read_mode}`)
  console.log('')
  console.log('指标                      改造前      实测      目标   结论')
  console.log('------------------------  ---------  --------  ------  ----')
  for (const row of lines) {
    const before = `${row.before} ${row.unit}`
    const actual = `${row.actual} ${row.unit}`
    console.log(`${row.label.padEnd(24)}  ${before.padStart(9)}  ${actual.padStart(8)}  ${String(row.limit).padStart(5)}  ${row.ok ? '✅' : '❌'}`)
  }
  console.log('')
  console.log(`结论：${failed === 0 ? `✅ ${lines.length} 项全部达标` : `❌ ${failed} 项未达标`}`)
}

process.exit(failed === 0 ? 0 : 3)
void pathToFileURL
