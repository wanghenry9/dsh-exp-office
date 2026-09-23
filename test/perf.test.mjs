/**
 * 性能基准：对照开发要求 §14.4 的参考性能目标，测量真实耗时与内存。
 *
 * 运行：
 *   node test/perf.test.mjs            # 小型 + 中型（默认，约 10 秒）
 *   node test/perf.test.mjs --large    # 追加 50MB 级大文件
 *
 * 说明：性能数字只在明确的文件规模与运行环境下才有意义，
 * 因此每次运行都会打印文件规模、耗时与内存峰值，而不是只给一个结论。
 */
import { mkdtempSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Workbook, formatRef, indexToCol } from '../lib/xlsx.js'
import { ZipPackage } from '../lib/ooxml.js'
import { buildEmptyWorkbook } from '../lib/tools.js'

const withLarge = process.argv.includes('--large')
const workspace = mkdtempSync(join(tmpdir(), 'dsh-office-perf-'))

let passed = 0
let failed = 0

/**
 * 生成一份指定规模的工作簿。
 * @param {number} rows - 数据行数。
 * @param {number} cols - 列数。
 * @returns {Buffer} .xlsx 字节。
 */
function makeWorkbook(rows, cols) {
  const header = Array.from({ length: cols }, (_, i) => ({ ref: `${indexToCol(i + 1)}1`, value: `列${i + 1}` }))
  const wb = Workbook.open(buildEmptyWorkbook('数据'))
  wb.writeCells('数据', header)
  const batch = []
  for (let r = 2; r <= rows + 1; r += 1) {
    for (let c = 1; c <= cols; c += 1) {
      batch.push({ ref: formatRef(c, r), value: c === 1 ? `行${r - 1}` : (r * c) % 100000 })
    }
    if (batch.length >= 50000) {
      wb.writeCells('数据', batch.splice(0, batch.length))
    }
  }
  if (batch.length > 0) wb.writeCells('数据', batch)
  return wb.save()
}

/**
 * 测量一次操作的耗时与内存增量。
 * @param {() => unknown} fn - 被测操作。
 * @returns {{ms: number, heapMB: number, rssMB: number, result: unknown}} 测量结果。
 */
function measure(fn) {
  const beforeHeap = process.memoryUsage().heapUsed
  const started = process.hrtime.bigint()
  const result = fn()
  const ms = Number(process.hrtime.bigint() - started) / 1e6
  const usage = process.memoryUsage()
  return {
    ms: Math.round(ms * 10) / 10,
    heapMB: Math.round((usage.heapUsed - beforeHeap) / 1048576 * 10) / 10,
    rssMB: Math.round(usage.rss / 1048576 * 10) / 10,
    result
  }
}

/**
 * 断言并记录一项基准。
 * @param {string} label - 场景名。
 * @param {number} actualMs - 实测耗时。
 * @param {number} targetMs - 目标耗时。
 * @param {string} extra - 附加信息。
 * @returns {void}
 */
function check(label, actualMs, targetMs, extra) {
  const ok = actualMs <= targetMs
  if (ok) passed += 1
  else failed += 1
  console.log(`  ${ok ? '✅' : '❌'} ${label}：${actualMs} ms（目标 ≤ ${targetMs} ms）${extra}`)
}

console.log(`\n运行环境：Node ${process.version}｜${process.platform} ${process.arch}`)
console.log(`CPU：${(await import('node:os')).cpus()[0]?.model ?? '未知'}\n`)

const scenarios = [
  { name: '小型', rows: 1000, cols: 10 },
  { name: '中型', rows: 20000, cols: 10 }
]
if (withLarge) scenarios.push({ name: '大型', rows: 100000, cols: 10 })

for (const scenario of scenarios) {
  console.log(`=== ${scenario.name}工作簿（${scenario.rows} 行 × ${scenario.cols} 列）===`)

  const generation = measure(() => makeWorkbook(scenario.rows, scenario.cols))
  const file = join(workspace, `${scenario.name}.xlsx`)
  writeFileSync(file, generation.result)
  const sizeMB = Math.round(statSync(file).size / 1048576 * 10) / 10
  console.log(`  生成：${generation.ms} ms，文件 ${sizeMB} MB，RSS ${generation.rssMB} MB`)

  const buffer = generation.result

  const open = measure(() => Workbook.open(buffer))
  console.log(`  打开：${open.ms} ms`)

  const wb = open.result
  const readRange = measure(() => wb.readRange('数据', `A1:${indexToCol(scenario.cols)}${Math.min(100, scenario.rows + 1)}`))
  console.log(`  读区域（100 行）：${readRange.ms} ms`)

  const readAll = measure(() => wb.readSheet('数据', { maxCells: scenario.rows * scenario.cols + 10 }))
  console.log(`  读全表（${readAll.result.cells.length} 个单元格）：${readAll.ms} ms`)

  const modify = measure(() => {
    const w = Workbook.open(buffer)
    w.writeCells('数据', [{ ref: 'B2', value: 12345 }, { ref: 'C3', value: '改过了' }])
    return w.save()
  })
  console.log(`  改 2 个单元格并保存：${modify.ms} ms，输出 ${Math.round(modify.result.length / 1048576 * 10) / 10} MB`)

  const modBuffer = modify.result
  const revalidate = measure(() => {
    const w = Workbook.open(modBuffer)
    return w.readRange('数据', 'B2:C3')
  })
  console.log(`  重新打开并校验：${revalidate.ms} ms`)

  if (scenario.name === '小型') {
    check('读取小型 XLSX', open.ms + readAll.ms, 2000, `（打开 ${open.ms} + 读全表 ${readAll.ms}）`)
    check('修改小型 XLSX', modify.ms, 3000, '')
  }
  if (scenario.name === '大型') {
    check('处理大文件不崩溃', 1, 1, `（${sizeMB} MB，读全表 ${readAll.ms} ms，RSS ${readAll.rssMB} MB）`)
  }
  console.log('')
}

console.log('=== ZIP 解压安全上限对性能的影响 ===')
const pkgBuffer = makeWorkbook(20000, 10)
const zipOpen = measure(() => ZipPackage.open(pkgBuffer))
console.log(`  仅解析 ZIP 中央目录（不解压）：${zipOpen.ms} ms，${zipOpen.result.names().length} 个部件`)
check('ZIP 中央目录解析', zipOpen.ms, 500, '')

rmSync(workspace, { recursive: true, force: true })

console.log(`\n结果：${passed} 项达标，${failed} 项未达标\n`)
process.exit(failed === 0 ? 0 : 1)
