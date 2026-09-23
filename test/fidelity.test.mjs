/**
 * 保真度回归测试：以「真实 Microsoft Excel 创作的多特性工作簿」为样本，
 * 断言修改普通单元格不会破坏图表、绘图、条件格式、合并单元格、主题与样式。
 *
 * 依赖 test/make-rich-fixture.ps1 先生成 test/fixtures/excel-rich.xlsx。
 * 若样本不存在则跳过（不失败），以便在无 Office 的环境下仍可运行其它测试。
 *
 * 运行：node test/fidelity.test.mjs
 */
import assert from 'node:assert/strict'
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Workbook, parseRef } from '../lib/xlsx.js'
import { ZipPackage, XmlDoc, findAll, find, attr } from '../lib/ooxml.js'

const here = dirname(fileURLToPath(import.meta.url))
const fixture = join(here, 'fixtures', 'excel-rich.xlsx')

if (!existsSync(fixture)) {
  console.log('⏭️  跳过：未找到 test/fixtures/excel-rich.xlsx（需先在装有 Excel 的机器上生成）。')
  process.exit(0)
}

let passed = 0
let failed = 0

/**
 * 运行一个用例并记录结果。
 * @param {string} name - 用例名。
 * @param {() => void} fn - 用例体。
 */
function test(name, fn) {
  try {
    fn()
    passed += 1
    console.log(`  ✅ ${name}`)
  } catch (err) {
    failed += 1
    console.log(`  ❌ ${name}\n     ${err.message}`)
  }
}

/**
 * 把 ZIP 包解析为「部件名 → 解压字节」映射。
 * @param {Buffer} buffer - ZIP 字节。
 * @returns {Map<string, Buffer>} 部件映射。
 */
function entriesOf(buffer) {
  const pkg = ZipPackage.open(buffer)
  const map = new Map()
  for (const name of pkg.names()) map.set(name, pkg.read(name))
  return map
}

const original = readFileSync(fixture)
const packageNames = ZipPackage.open(original).names()

console.log('\n=== 样本特征（由真实 Excel 创作）===')

test('样本含图表部件', () => {
  assert.ok(packageNames.some((n) => /^xl\/charts\/chart\d+\.xml$/.test(n)), `实际部件：${packageNames.join(', ')}`)
})

test('样本含绘图与图表关系部件', () => {
  assert.ok(packageNames.some((n) => n.startsWith('xl/drawings/')), '缺少绘图部件')
})

test('样本含主题与样式部件', () => {
  assert.ok(packageNames.includes('xl/theme/theme1.xml'))
  assert.ok(packageNames.includes('xl/styles.xml'))
})

const sheet1 = ZipPackage.open(original).readText('xl/worksheets/sheet1.xml')

// Excel 的 Worksheets.Add() 会把新表插到当前表之前，因此按名字定位部件而不是假定 sheet1。
const probe = Workbook.open(original)
const salesPart = probe.sheet('销售数据').path
const salesXml = ZipPackage.open(original).readText(salesPart)

test('样本含条件格式规则', () => {
  assert.ok(salesXml.includes('<conditionalFormatting'), '未找到 conditionalFormatting')
})

test('样本含合并单元格', () => {
  assert.ok(salesXml.includes('<mergeCell'), '未找到 mergeCell')
})

test('样本含公式', () => {
  assert.ok(salesXml.includes('<f>'), '未找到公式节点')
})

test('样本工作表含图表绘图引用', () => {
  assert.ok(/<drawing r:id=/.test(salesXml), '未找到 drawing 引用')
})

console.log('\n=== 修改单元格后的保真度 ===')

const wb = Workbook.open(original)
const before = entriesOf(original)

// 只改一个普通数值单元格
wb.writeCells('销售数据', [{ ref: 'B2', value: 99999 }])
const after = entriesOf(wb.save())

test('目标单元格确实被改写', () => {
  const reopened = Workbook.open(wb.save())
  assert.equal(reopened.readRange('销售数据', 'B2:B2').rows[0][0].value, 99999)
})

test('★ 图表部件逐字节不变', () => {
  for (const name of before.keys()) {
    if (!/^xl\/charts\//.test(name)) continue
    assert.ok(after.get(name).equals(before.get(name)), `${name} 被改动`)
  }
})

test('★ 绘图与图表关系部件逐字节不变', () => {
  for (const name of before.keys()) {
    if (!name.startsWith('xl/drawings/')) continue
    assert.ok(after.get(name).equals(before.get(name)), `${name} 被改动`)
  }
})

test('★ 主题、样式、共享字符串、文档属性全部逐字节不变', () => {
  for (const name of ['xl/theme/theme1.xml', 'xl/styles.xml', 'docProps/core.xml', 'docProps/app.xml', '[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels']) {
    if (!before.has(name)) continue
    assert.ok(after.get(name).equals(before.get(name)), `${name} 被改动`)
  }
})

test('★ 只有目标工作表部件发生变化', () => {
  const changed = []
  for (const [name, bytes] of before) {
    if (!after.get(name).equals(bytes)) changed.push(name)
  }
  assert.deepEqual(changed, [salesPart], `意外改动的部件：${changed.join(', ')}（目标应为 ${salesPart}）`)
})

test('★ 条件格式规则原样保留', () => {
  const afterSheet = ZipPackage.open(wb.save()).readText(salesPart)
  const beforeRules = findAll(XmlDoc.parse(salesXml).root, 'conditionalFormatting').map((n) => attr(n, 'sqref'))
  const afterRules = findAll(XmlDoc.parse(afterSheet).root, 'conditionalFormatting').map((n) => attr(n, 'sqref'))
  assert.deepEqual(afterRules, beforeRules)
  assert.ok(afterSheet.includes('<conditionalFormatting'))
})

test('★ 合并单元格原样保留', () => {
  const afterSheet = ZipPackage.open(wb.save()).readText(salesPart)
  const beforeMerges = findAll(XmlDoc.parse(salesXml).root, 'mergeCell').map((n) => attr(n, 'ref'))
  const afterMerges = findAll(XmlDoc.parse(afterSheet).root, 'mergeCell').map((n) => attr(n, 'ref'))
  assert.deepEqual(afterMerges, beforeMerges)
  assert.ok(beforeMerges.length > 0)
})

test('★ 图表锚点与绘图引用未被改写', () => {
  const afterSheet = ZipPackage.open(wb.save()).readText(salesPart)
  assert.ok(afterSheet.includes('<drawing r:id='), '工作表丢失了 drawing 引用')
})

test('★ 其它工作表的跨表公式原样保留', () => {
  const reopened = Workbook.open(wb.save())
  const cell = reopened.readRange('汇总', 'B2:B2').rows[0][0]
  assert.match(cell.formula, /销售数据'?!B8$/)
})

test('★ 未被触碰的行 XML 逐字节一致', () => {
  const afterSheet = ZipPackage.open(wb.save()).readText(salesPart)
  const cut = (t) => t.slice(t.indexOf('<row r="3"'))
  assert.equal(cut(afterSheet), cut(salesXml), '第 3 行及其后的 XML 被改动')
})

// 输出供真实 Excel/WPS 再次打开验证
writeFileSync(join(here, 'fixtures', 'excel-rich-modified.xlsx'), wb.save())

console.log(`\n结果：${passed} 通过，${failed} 失败`)
console.log(`输出文件：test/fixtures/excel-rich-modified.xlsx\n`)
process.exit(failed === 0 ? 0 : 1)
