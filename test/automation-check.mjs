/**
 * 本地 Office/WPS 联动的**真机**验证（阶段 6，需要本机装了 Office/WPS）。
 *
 * 跑什么：
 *   1. 检测：读到真实引擎与版本（只读注册表）；
 *   2. 重算：故意写一个缓存值错误的公式（=1+2 缓存成 999），用真实 Excel 重算后应变成 3；
 *   3. 同一个文件换 WPS 引擎再来一遍（装了 WPS 才有这一步）；
 *   4. 重渲染：Word / PowerPoint 各重写一份，产物必须能被自己的读取器解析且规模一致；
 *   5. 每次调用后都断言：源文件逐字节未变、没有留下本次启动的引擎进程。
 *
 * 运行：node test/automation-check.mjs
 */
import assert from 'node:assert/strict'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync, copyFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficeAutomation } from '../lib/automation.js'
import { DocxDocument, buildBasicDocx } from '../lib/docx.js'
import { PptxPresentation } from '../lib/pptx.js'
import { Workbook } from '../lib/xlsx.js'
import { buildEmptyWorkbook } from '../lib/tools.js'

const here = dirname(fileURLToPath(import.meta.url))
const fixtureDir = join(here, 'fixtures')
const workspace = mkdtempSync(join(tmpdir(), 'dsh-office-engine-'))

let passed = 0
let failed = 0
let skipped = 0

/**
 * 运行一个用例并记录结果。
 * @param {string} name - 用例名。
 * @param {() => Promise<void>} fn - 用例体。
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

/**
 * 跳过一个用例（环境不具备时如实说明，不算失败也不算通过）。
 * @param {string} name - 用例名。
 * @param {string} reason - 原因。
 * @returns {void}
 */
function skip(name, reason) {
  skipped += 1
  console.log(`  ⏭️  ${name}（跳过：${reason}）`)
}

/**
 * 造一个「公式缓存值故意写错」的工作簿。
 * @returns {Buffer} 字节。
 */
function makeStaleFormulaWorkbook() {
  const wb = Workbook.open(buildEmptyWorkbook('数据'))
  wb.writeCells('数据', [
    { ref: 'A1', value: '公式重算验证' },
    { ref: 'B1', value: 999, formula: '1+2' },
    { ref: 'C1', value: 111, formula: 'SUM(1,2,3)' }
  ])
  return wb.save()
}

const automation = new OfficeAutomation({
  enabled: true,
  scratchDir: join(workspace, 'scratch'),
  auditPath: join(workspace, 'audit.jsonl'),
  timeoutMs: 180000
})

console.log(`\n工作目录：${workspace}`)
console.log(`自动化脚本：${join(here, '..', 'lib', 'office-automation.ps1')}\n`)

console.log('=== 1. 检测本机引擎 ===')

let detect = null
await test('读到引擎清单与版本（只读注册表，不启动程序）', async () => {
  detect = await automation.detectEngines({})
  assert.ok(detect.engines.length >= 3, '至少应报告 Excel/Word/PowerPoint 三类')
  for (const engine of detect.engines) {
    assert.equal(typeof engine.prog_id, 'string')
    assert.equal(typeof engine.installed, 'boolean')
  }
  const installed = detect.engines.filter((e) => e.installed).map((e) => `${e.label} ${e.file_version ?? ''}`)
  console.log(`     本机引擎：${installed.length > 0 ? installed.join('、') : '无'}`)
})

const hasOffice = detect?.microsoft_office === true
const hasWps = detect?.wps === true

console.log('\n=== 2. Excel / WPS 重算（公式缓存）===')

const stalePath = join(workspace, 'stale.xlsx')
writeFileSync(stalePath, makeStaleFormulaWorkbook())
const staleBytes = readFileSync(stalePath)

if (!hasOffice) {
  skip('Excel 重算', '本机未安装 Microsoft Office')
} else {
  await test('真实 Excel 重算：缓存值 999 → 3、111 → 6，源文件逐字节不变', async () => {
    const out = join(workspace, 'excel-recalc.xlsx')
    const result = await automation.recalculate({ inputPath: stalePath, outputPath: out, kind: 'xlsx', engine: 'excel' })
    assert.equal(result.engine, 'excel')
    assert.deepEqual(result.leftover_pids, [], `不应残留进程：${result.leftover_pids.join(',')}`)
    const wb = Workbook.open(readFileSync(out))
    assert.equal(wb.readRange('数据', 'B1:B1').rows[0][0].value, 3, 'B1 的 1+2 应被重算为 3')
    assert.equal(wb.readRange('数据', 'C1:C1').rows[0][0].value, 6, 'C1 的 SUM(1,2,3) 应被重算为 6')
    assert.ok(readFileSync(stalePath).equals(staleBytes), '源文件必须逐字节不变')
  })
}

if (!hasWps) {
  skip('WPS 重算', '本机未安装 WPS')
} else {
  await test('WPS 表格重算：结果与 Excel 一致，源文件逐字节不变', async () => {
    const out = join(workspace, 'wps-recalc.xlsx')
    const result = await automation.recalculate({ inputPath: stalePath, outputPath: out, kind: 'xlsx', engine: 'wps' })
    assert.equal(result.engine, 'wps')
    assert.deepEqual(result.leftover_pids, [])
    const wb = Workbook.open(readFileSync(out))
    assert.equal(wb.readRange('数据', 'B1:B1').rows[0][0].value, 3)
    assert.equal(wb.readRange('数据', 'C1:C1').rows[0][0].value, 6)
    assert.ok(readFileSync(stalePath).equals(staleBytes), '源文件必须逐字节不变')
  })
}

console.log('\n=== 3. Word / PowerPoint 重渲染 ===')

const docxPath = join(workspace, 'source.docx')
writeFileSync(docxPath, buildBasicDocx({ title: '重渲染验证' }))
const docxBytes = readFileSync(docxPath)

if (!hasOffice) {
  skip('Word 重渲染', '本机未安装 Microsoft Office')
} else {
  await test('真实 Word 重写一份：产物可被自身读取器解析、段落数一致、源文件不变', async () => {
    const out = join(workspace, 'word-rerendered.docx')
    const result = await automation.rerender({ inputPath: docxPath, outputPath: out, kind: 'docx', engine: 'word' })
    assert.equal(result.engine, 'word')
    assert.deepEqual(result.leftover_pids, [])
    const before = DocxDocument.open(docxBytes).paragraphs().total
    const after = DocxDocument.open(readFileSync(out)).paragraphs().total
    assert.equal(after, before, `段落数应一致：${after} ↔ ${before}`)
    assert.ok(readFileSync(docxPath).equals(docxBytes), '源文件必须逐字节不变')
  })
}

const deckSource = join(fixtureDir, 'deck-edited.pptx')
if (!existsSync(deckSource)) {
  skip('PowerPoint 重渲染', `缺少样本 ${deckSource}`)
} else if (!hasOffice) {
  skip('PowerPoint 重渲染', '本机未安装 Microsoft Office')
} else {
  await test('真实 PowerPoint 重写一份：页数一致、源文件不变', async () => {
    const source = join(workspace, 'deck-source.pptx')
    copyFileSync(deckSource, source)
    const sourceBytes = readFileSync(source)
    const out = join(workspace, 'deck-rerendered.pptx')
    const result = await automation.rerender({ inputPath: source, outputPath: out, kind: 'pptx', engine: 'powerpoint' })
    assert.equal(result.engine, 'powerpoint')
    assert.deepEqual(result.leftover_pids, [])
    const before = PptxPresentation.open(sourceBytes).slides({ includeShapes: false }).count
    const after = PptxPresentation.open(readFileSync(out)).slides({ includeShapes: false }).count
    assert.equal(after, before)
    assert.ok(readFileSync(source).equals(sourceBytes), '源文件必须逐字节不变')
  })
}

console.log('\n=== 4. 进程与审计 ===')

await test('调用结束后没有本次启动的引擎进程残留', async () => {
  const again = await automation.detectEngines({})
  assert.deepEqual(again.leftover_pids, [], `检测后仍有残留：${again.leftover_pids.join(',')}`)
})

await test('审计日志按行记录且不含绝对路径', () => {
  const raw = readFileSync(join(workspace, 'audit.jsonl'), 'utf8')
  const lines = raw.trim().split('\n').filter((l) => l !== '')
  assert.ok(lines.length >= 1, '应至少有一条审计记录')
  for (const line of lines) {
    const entry = JSON.parse(line)
    assert.equal(typeof entry.action, 'string')
    assert.equal(typeof entry.ok, 'boolean')
  }
  assert.equal(raw.includes(workspace), false, '审计日志不应出现本机绝对路径')
})

rmSync(workspace, { recursive: true, force: true })

console.log(`\n结果：${passed} 通过，${skipped} 跳过，${failed} 失败`)
console.log(`说明：跳过的用例表示本机缺少对应引擎，不是通过。\n`)
process.exit(failed === 0 ? 0 : 1)
void mkdirSync
