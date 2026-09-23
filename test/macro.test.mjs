/**
 * 宏启用工作簿（.xlsm）测试。
 *
 * **这个套件能证明什么、不能证明什么，必须说清楚**：
 *   能证明：① 真实类型检测识别出 xlsm 与「含宏」；
 *           ② 编辑普通单元格后 `xl/vbaProject.bin` 逐字节不变（宏按原字节保留）；
 *           ③ 编辑后的文件仍能被本插件重新打开、宏标志仍在。
 *   不能证明：Excel 的 VBA 编辑器仍能运行其中的宏。
 *           样本里的 vbaProject.bin 是**合成**的（带 OLE 复合文档头但非真实 VBA 工程），
 *           Excel 会拒绝打开它。要验证这一点需要一个真实的宏文件样本，
 *           或在 Office 中开启「信任对 VBA 工程对象模型的访问」后由 Excel 生成 ——
 *           后者属于改动用户的 Office 安全设置，未获授权不做。
 *
 * 运行：node test/macro.test.mjs
 */
import assert from 'node:assert/strict'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Workbook } from '../lib/xlsx.js'
import { ZipPackage } from '../lib/ooxml.js'
import { detectFileType } from '../lib/workspace.js'
import { syntheticVbaProject, toMacroEnabled } from './make-macro-fixture.mjs'
import { buildEmptyWorkbook } from '../lib/tools.js'

const here = dirname(fileURLToPath(import.meta.url))
const fixture = join(here, 'fixtures', 'macro.xlsm')

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

if (!existsSync(fixture)) {
  console.log('⏭️  跳过：未找到 test/fixtures/macro.xlsm（先运行 node test/make-macro-fixture.mjs）。')
  process.exit(0)
}

const original = readFileSync(fixture)
const vbaBytes = ZipPackage.open(original).read('xl/vbaProject.bin')

console.log('\n=== 1. 宏文件识别 ===')

test('真实类型检测识别为 xlsm 且标记含宏', () => {
  const type = detectFileType(original, 'macro.xlsm')
  assert.equal(type.ext, 'xlsm')
  assert.equal(type.macro, true)
  assert.equal(type.mismatch, false)
})

test('把 .xlsm 改名成 .xlsx 也识破为宏文件', () => {
  const type = detectFileType(original, '伪装.xlsx')
  assert.equal(type.ext, 'xlsm')
  assert.equal(type.macro, true)
  assert.equal(type.mismatch, true, '应报告扩展名与实际类型不符')
})

test('工作簿对象报告 hasMacro', () => {
  assert.equal(Workbook.open(original).info.hasMacro, true)
})

test('宏部件存在于包内', () => {
  assert.ok(ZipPackage.open(original).has('xl/vbaProject.bin'))
})

console.log('\n=== 2. 编辑后的宏保留 ===')

test('★ 编辑普通单元格后 vbaProject.bin 逐字节不变', () => {
  const before = ZipPackage.open(original)
  const wb = Workbook.open(original)
  wb.writeCells('数据', [
    { ref: 'A1', value: '改过了' },
    { ref: 'B1', value: 42 },
    { ref: 'A2', value: '新行' }
  ])
  const after = ZipPackage.open(wb.save())
  assert.ok(after.read('xl/vbaProject.bin').equals(before.read('xl/vbaProject.bin')), 'vbaProject.bin 被改动')
  assert.ok(after.read('xl/vbaProject.bin').equals(vbaBytes), '宏字节与原始样本不一致')
})

test('★ 编辑只影响工作表与共享字符串，宏与主题等部件不动', () => {
  const before = ZipPackage.open(original)
  const wb = Workbook.open(original)
  wb.writeCells('数据', [{ ref: 'D5', value: 'x' }])
  const after = ZipPackage.open(wb.save())
  const changed = before.names().filter((n) => !after.read(n).equals(before.read(n)))
  assert.deepEqual(changed.sort(), ['xl/sharedStrings.xml', 'xl/worksheets/sheet1.xml'], `意外改动：${changed.join(', ')}`)
})

test('★ 编辑后重新打开，宏标志仍在且数据正确', () => {
  const wb = Workbook.open(original)
  wb.writeCells('数据', [{ ref: 'A1', value: '改过了' }])
  const reopened = Workbook.open(wb.save())
  assert.equal(reopened.info.hasMacro, true)
  assert.equal(reopened.readRange('数据', 'A1:A1').rows[0][0].value, '改过了')
})

test('★ 多次编辑后宏部件依然完好', () => {
  let buffer = original
  for (let i = 0; i < 5; i += 1) {
    const wb = Workbook.open(buffer)
    wb.writeCells('数据', [{ ref: 'A1', value: `第 ${i} 轮` }])
    buffer = wb.save()
  }
  const pkg = ZipPackage.open(buffer)
  assert.ok(pkg.read('xl/vbaProject.bin').equals(vbaBytes), '多轮编辑后宏字节发生变化')
  assert.equal(Workbook.open(buffer).info.hasMacro, true)
})

test('★ 增删工作表后宏部件依然完好', () => {
  const wb = Workbook.open(original)
  wb.addWorksheet('新表')
  wb.deleteWorksheet('新表')
  const pkg = ZipPackage.open(wb.save())
  assert.ok(pkg.read('xl/vbaProject.bin').equals(vbaBytes))
})

console.log('\n=== 3. 生成器的确定性 ===')

test('合成宏部件具有 OLE 复合文档签名', () => {
  assert.deepEqual([...vbaBytes.subarray(0, 8)], [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
})

test('可从零构造同构样本（生成器可复现）', () => {
  const rebuilt = toMacroEnabled(buildEmptyWorkbook('数据'), syntheticVbaProject())
  const pkg = ZipPackage.open(rebuilt)
  assert.ok(pkg.has('xl/vbaProject.bin'))
  assert.ok(pkg.readText('[Content_Types].xml').includes('vbaProject'))
  assert.ok(pkg.readText('xl/_rels/workbook.xml.rels').includes('vbaProject'))
  assert.equal(Workbook.open(rebuilt).info.hasMacro, true)
})

console.log('\n⚠️  未验证：Excel 能否运行其中的宏。样本的 vbaProject.bin 是合成的（非真实 VBA 工程），')
console.log('    Excel 会拒绝打开；要验证这一点需要真实宏样本或开启 Office 的 VBA 工程访问信任。')

console.log(`\n结果：${passed} 通过，${failed} 失败\n`)
process.exit(failed === 0 ? 0 : 1)
