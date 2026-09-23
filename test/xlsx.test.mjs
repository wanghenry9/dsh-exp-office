/**
 * XLSX 适配器测试：以第三方库（SheetJS）产出的真实工作簿为样本，
 * 验证读取、写入、保真度（除目标部件外逐字节不变）与安全防护。
 *
 * 运行：node test/xlsx.test.mjs
 */
import assert from 'node:assert/strict'
import os from 'node:os'
import { createRequire } from 'node:module'
import { mkdirSync, writeFileSync, readFileSync, existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Workbook, parseRef, formatRef, colToIndex, indexToCol, parseRange, serialToIso, isoToSerial } from '../lib/xlsx.js'
import { ZipPackage, XmlDoc, findAll } from '../lib/ooxml.js'

const here = dirname(fileURLToPath(import.meta.url))
const fixtureDir = join(here, 'fixtures')
mkdirSync(fixtureDir, { recursive: true })

// SheetJS 只用来「生成一份第三方库产出的样本」，插件本身不依赖它。
// 解析顺序：工作区自己的依赖 → 本机 DSH profile 的依赖。
// 两者都没有时**不报错**：改用仓库里已经由 SheetJS 产出并提交的 fixtures/sales.xlsx，
// 样本同样是真实第三方实现写出来的，测试意图不变。
let XLSX = null
const requireAnchors = [
  join(here, '..', 'node_modules', 'noop.js'),
  `${os.homedir()}/.dsh/profiles/web/node_modules/`,
  `${os.homedir()}/.dsh/profiles/node_modules/`
]
for (const anchor of requireAnchors) {
  try {
    XLSX = createRequire(anchor)('xlsx')
    break
  } catch {
    XLSX = null
  }
}
if (XLSX) {
  console.log(`（样本由 SheetJS ${XLSX.version ?? ''} 现场生成）`)
} else {
  console.log('（未找到 SheetJS，改用仓库中已提交的 SheetJS 样本 test/fixtures/sales.xlsx）')
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
 * 生成样本：优先用 SheetJS 现场生成；没有 SheetJS 时读仓库里已提交的同源样本。
 * @returns {Buffer} .xlsx 字节。
 */
function buildFixture() {
  if (!XLSX) {
    const committed = join(fixtureDir, 'sales.xlsx')
    if (!existsSync(committed)) {
      throw new Error('既没有 SheetJS，也没有已提交的 test/fixtures/sales.xlsx 样本。')
    }
    return readFileSync(committed)
  }
  const wb = XLSX.utils.book_new()
  const sales = XLSX.utils.aoa_to_sheet([
    ['产品', '单价', '数量', '金额'],
    ['笔记本', 5999.5, 3, { t: 'n', f: 'B2*C2' }],
    ['显示器', 1299, 10, { t: 'n', f: 'B3*C3' }],
    ['键盘', 299.9, 25, { t: 'n', f: 'B4*C4' }],
    ['合计', null, null, { t: 'n', f: 'SUM(D2:D4)' }]
  ])
  XLSX.utils.book_append_sheet(wb, sales, '销售数据')
  const summary = XLSX.utils.aoa_to_sheet([
    ['指标', '数值'],
    ['销售总额', { t: 'n', f: "'销售数据'!D5" }],
    ['备注', '本表由测试生成']
  ])
  XLSX.utils.book_append_sheet(wb, summary, '汇总')
  return Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }))
}

/**
 * 把 ZIP 包解析为「条目名 → 解压字节」的映射。
 * @param {Buffer} buffer - ZIP 字节。
 * @returns {Map<string, Buffer>} 条目映射。
 */
function entriesOf(buffer) {
  const pkg = ZipPackage.open(buffer)
  const map = new Map()
  for (const name of pkg.names()) map.set(name, pkg.read(name))
  return map
}

const original = buildFixture()
writeFileSync(join(fixtureDir, 'sales.xlsx'), original)

console.log('\n=== 1. 引用工具 ===')

test('列号与列字母互转', () => {
  assert.equal(colToIndex('A'), 1)
  assert.equal(colToIndex('Z'), 26)
  assert.equal(colToIndex('AA'), 27)
  assert.equal(colToIndex('AB'), 28)
  assert.equal(indexToCol(1), 'A')
  assert.equal(indexToCol(26), 'Z')
  assert.equal(indexToCol(27), 'AA')
  assert.equal(indexToCol(703), 'AAA')
})

test('单元格引用解析与格式化', () => {
  assert.deepEqual(parseRef('B12'), { col: 2, row: 12 })
  assert.deepEqual(parseRef('$AA$3'), { col: 27, row: 3 })
  assert.equal(formatRef(2, 12), 'B12')
  assert.throws(() => parseRef('12B'), (e) => e.code === 'INVALID_REQUEST')
})

test('区域解析会自动纠正反向边界', () => {
  assert.deepEqual(parseRange('C10:A1'), { start: { col: 1, row: 1 }, end: { col: 3, row: 10 } })
  assert.deepEqual(parseRange('B2'), { start: { col: 2, row: 2 }, end: { col: 2, row: 2 } })
})

test('日期序列值双向转换', () => {
  const iso = serialToIso(45000, false)
  assert.equal(isoToSerial(iso, false), 45000)
  assert.equal(serialToIso(0, false), '1899-12-30T00:00:00.000Z')
})

console.log('\n=== 2. 读取 SheetJS 生成的真实工作簿 ===')

test('识别工作表列表与顺序', () => {
  const wb = Workbook.open(original)
  assert.deepEqual(wb.sheetNames(), ['销售数据', '汇总'])
})

test('读取单元格文本与数值', () => {
  const wb = Workbook.open(original)
  const range = wb.readRange('销售数据', 'A1:D5')
  assert.equal(range.rows[0][0].value, '产品')
  assert.equal(range.rows[1][0].value, '笔记本')
  assert.equal(range.rows[1][1].value, 5999.5)
  assert.equal(range.rows[1][2].value, 3)
})

test('区分公式本身与公式缓存结果', () => {
  const wb = Workbook.open(original)
  const range = wb.readRange('销售数据', 'D2:D5')
  assert.equal(range.rows[0][0].formula, 'B2*C2')
  assert.equal(range.rows[3][0].formula, 'SUM(D2:D4)')
})

test('跨工作表公式保留原始引用写法', () => {
  const wb = Workbook.open(original)
  const cell = wb.readRange('汇总', 'B2:B2').rows[0][0]
  assert.equal(cell.formula, "'销售数据'!D5")
})

console.log('\n=== 3. 写入与保真度 ===')

test('写入既有单元格后重新读取可见新值', () => {
  const wb = Workbook.open(original)
  wb.writeCells('销售数据', [{ ref: 'B2', value: 6888 }, { ref: 'A2', value: '笔记本电脑' }])
  const out = wb.save()
  const reopened = Workbook.open(out)
  const range = reopened.readRange('销售数据', 'A2:B2')
  assert.equal(range.rows[0][0].value, '笔记本电脑')
  assert.equal(range.rows[0][1].value, 6888)
})

test('写入新行新列（超出原范围）', () => {
  const wb = Workbook.open(original)
  wb.writeCells('销售数据', [{ ref: 'F2', value: '新增列' }, { ref: 'A8', value: '新增行' }, { ref: 'B8', value: 123 }])
  const reopened = Workbook.open(wb.save())
  const range = reopened.readRange('销售数据', 'A8:F8')
  assert.equal(range.rows[0][0].value, '新增行')
  assert.equal(range.rows[0][1].value, 123)
  assert.equal(reopened.readRange('销售数据', 'F2:F2').rows[0][0].value, '新增列')
})

test('公式单元格保留公式，仅替换缓存值', () => {
  const wb = Workbook.open(original)
  wb.writeCells('销售数据', [{ ref: 'D2', formula: 'B2*C2', value: 20664 }])
  const reopened = Workbook.open(wb.save())
  const cell = reopened.readRange('销售数据', 'D2:D2').rows[0][0]
  assert.equal(cell.formula, 'B2*C2')
  assert.equal(cell.value, 20664)
})

test('写入布尔值与清空单元格', () => {
  const wb = Workbook.open(original)
  wb.writeCells('销售数据', [{ ref: 'E2', value: true }, { ref: 'E3', value: null }])
  const reopened = Workbook.open(wb.save())
  assert.equal(reopened.readRange('销售数据', 'E2:E2').rows[0][0].value, true)
  assert.equal(reopened.readRange('销售数据', 'E3:E3').rows[0][0], null)
})

test('★ 保真度：除目标工作表外所有部件逐字节不变', () => {
  const wb = Workbook.open(original)
  wb.writeCells('销售数据', [{ ref: 'B2', value: 6888 }])
  const after = entriesOf(wb.save())
  const before = entriesOf(original)

  assert.deepEqual([...after.keys()].sort(), [...before.keys()].sort(), '部件清单发生了变化')
  const changed = []
  for (const [name, bytes] of before) {
    if (!after.get(name).equals(bytes)) changed.push(name)
  }
  // 只允许工作表部件与共享字符串表变化
  assert.ok(
    changed.every((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n) || n === 'xl/sharedStrings.xml'),
    `意外的部件被改动：${changed.join(', ')}`
  )
  assert.ok(changed.includes('xl/worksheets/sheet1.xml'), '目标工作表应被修改')
  // 未改动的工作表必须逐字节相同
  assert.ok(after.get('xl/worksheets/sheet2.xml').equals(before.get('xl/worksheets/sheet2.xml')), '无关工作表被改动')
})

test('★ 保真度：未触碰的工作表 XML 内容逐字节保留', () => {
  const wb = Workbook.open(original)
  wb.writeCells('销售数据', [{ ref: 'B2', value: 6888 }])
  const out = wb.save()
  const before = ZipPackage.open(original).readText('xl/worksheets/sheet1.xml')
  const after = ZipPackage.open(out).readText('xl/worksheets/sheet1.xml')
  // 除 B2 单元格所在片段外，其余部分长度差异应极小
  const beforeOther = before.replace(/<c r="B2"[\s\S]*?<\/c>|<c r="B2"[^>]*\/>/, '')
  const afterOther = after.replace(/<c r="B2"[\s\S]*?<\/c>|<c r="B2"[^>]*\/>/, '')
  assert.equal(afterOther, beforeOther, 'B2 之外的 XML 被改动')
})

test('写入长文本与特殊字符', () => {
  const wb = Workbook.open(original)
  const tricky = '中文<标签>&符号"引号" 换行\n制表\t'
  wb.writeCells('销售数据', [{ ref: 'A10', value: tricky }])
  const reopened = Workbook.open(wb.save())
  assert.equal(reopened.readRange('销售数据', 'A10:A10').rows[0][0].value, tricky)
})

test('无共享字符串表时退化为内联字符串且不新增部件', () => {
  const wb = Workbook.open(original)
  wb.writeCells('销售数据', [{ ref: 'A20', value: '复用' }, { ref: 'A21', value: '复用' }])
  const out = wb.save()
  const names = ZipPackage.open(out).names()
  assert.ok(!names.includes('xl/sharedStrings.xml'), '不应凭空新增 sharedStrings 部件')
  const reopened = Workbook.open(out)
  assert.equal(reopened.readRange('销售数据', 'A20:A21').rows[0][0].value, '复用')
  assert.equal(reopened.readRange('销售数据', 'A20:A21').rows[1][0].value, '复用')
})

test('有共享字符串表时复用索引并更新计数', () => {
  // 手工构造一个带 sharedStrings 的最小工作簿（Excel 的常规形态）。
  const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { maxEntries: 100, maxEntryBytes: 1e6, maxTotalBytes: 1e7, maxRatio: 200 })
  pkg.write(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`
  )
  pkg.write(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
  )
  pkg.write(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`
  )
  pkg.write(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`
  )
  pkg.write('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>已有</t></si></sst>`)
  pkg.write(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`
  )
  const withSst = pkg.toBuffer()
  writeFileSync(join(fixtureDir, 'shared-strings.xlsx'), withSst)

  const wb = Workbook.open(withSst)
  assert.equal(wb.readRange('Sheet1', 'A1').rows[0][0].value, '已有')
  wb.writeCells('Sheet1', [{ ref: 'B1', value: '新词' }, { ref: 'C1', value: '已有' }])
  const out = wb.save()
  const sstXml = ZipPackage.open(out).readText('xl/sharedStrings.xml')
  // 「已有」复用索引 0，只新增一条 <si>
  assert.equal((sstXml.match(/<si>/g) ?? []).length, 2)
  assert.ok(/count="3"/.test(sstXml), `count 应更新为 3，实际：${sstXml}`)
  const reopened = Workbook.open(out)
  assert.equal(reopened.readRange('Sheet1', 'B1').rows[0][0].value, '新词')
  assert.equal(reopened.readRange('Sheet1', 'C1').rows[0][0].value, '已有')
})

console.log('\n=== 4. 工作表管理 ===')

test('新增工作表并写入', () => {
  const wb = Workbook.open(original)
  wb.addWorksheet('新表')
  wb.writeCells('新表', [{ ref: 'A1', value: '你好' }])
  const out = wb.save()
  const reopened = Workbook.open(out)
  assert.deepEqual(reopened.sheetNames(), ['销售数据', '汇总', '新表'])
  assert.equal(reopened.readRange('新表', 'A1:A1').rows[0][0].value, '你好')
})

test('重命名工作表', () => {
  const wb = Workbook.open(original)
  wb.renameWorksheet('汇总', '统计')
  const reopened = Workbook.open(wb.save())
  assert.deepEqual(reopened.sheetNames(), ['销售数据', '统计'])
})

test('删除工作表', () => {
  const wb = Workbook.open(original)
  wb.deleteWorksheet('汇总')
  const reopened = Workbook.open(wb.save())
  assert.deepEqual(reopened.sheetNames(), ['销售数据'])
})

test('拒绝删除最后一张可见工作表', () => {
  const wb = Workbook.open(original)
  wb.deleteWorksheet('汇总')
  assert.throws(() => wb.deleteWorksheet('销售数据'), (e) => e.code === 'INVALID_REQUEST')
})

test('拒绝重名与非法工作表名', () => {
  const wb = Workbook.open(original)
  assert.throws(() => wb.addWorksheet('汇总'), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => wb.addWorksheet('非法/名称'), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => wb.addWorksheet('x'.repeat(32)), (e) => e.code === 'INVALID_REQUEST')
})

console.log('\n=== 5. 查找替换与合并 ===')

test('查找替换只影响命中单元格', () => {
  const wb = Workbook.open(original)
  const result = wb.findAndReplace({ find: '笔记本', replace: '便携机' })
  assert.equal(result.replacements, 1)
  const reopened = Workbook.open(wb.save())
  assert.equal(reopened.readRange('销售数据', 'A2:A2').rows[0][0].value, '便携机')
  assert.equal(reopened.readRange('销售数据', 'A3:A3').rows[0][0].value, '显示器')
})

test('默认不修改公式内容', () => {
  const wb = Workbook.open(original)
  const result = wb.findAndReplace({ find: 'SUM', replace: 'MAX' })
  assert.equal(result.replacements, 0)
})

test('整格匹配模式', () => {
  const wb = Workbook.open(original)
  const result = wb.findAndReplace({ find: '笔记', replace: 'X', wholeCell: true })
  assert.equal(result.replacements, 0)
})

test('合并单元格写入 mergeCells 节点', () => {
  const wb = Workbook.open(original)
  wb.mergeCells('销售数据', 'A6:C6')
  const xml = ZipPackage.open(wb.save()).readText('xl/worksheets/sheet1.xml')
  assert.ok(xml.includes('<mergeCell ref="A6:C6"/>'))
  assert.ok(xml.includes('<mergeCells'))
})

test('重复合并同一区域被拒绝', () => {
  const wb = Workbook.open(original)
  wb.mergeCells('销售数据', 'A6:C6')
  assert.throws(() => wb.mergeCells('销售数据', 'A6:C6'), (e) => e.code === 'INVALID_REQUEST')
})

console.log('\n=== 6. 行列移位 ===')

test('插入行会下移后续行号与单元格引用', () => {
  const wb = Workbook.open(original)
  wb.shiftRowsOrColumns('销售数据', { axis: 'row', action: 'insert', start: 2, count: 1 })
  const reopened = Workbook.open(wb.save())
  assert.equal(reopened.readRange('销售数据', 'A3:A3').rows[0][0].value, '笔记本')
})

test('插入列会右移后续列引用', () => {
  const wb = Workbook.open(original)
  wb.shiftRowsOrColumns('销售数据', { axis: 'column', action: 'insert', start: 2, count: 1 })
  const reopened = Workbook.open(wb.save())
  assert.equal(reopened.readRange('销售数据', 'C2:C2').rows[0][0].value, 5999.5)
})

console.log('\n=== 7. 表格对象（ListObject） ===')

/** 检查一个 XML 部件是否是「格式良好」的单根文档（标签配平、只有一个根元素）。 */
function assertSingleRoot(xml, label) {
  const body = xml.replace(/<\?xml[^>]*\?>/g, '').replace(/<!--[\s\S]*?-->/g, '')
  let depth = 0
  let rootClosed = false
  let extra = 0
  for (const match of body.matchAll(/<[^>]+>/g)) {
    const tag = match[0]
    if (tag.startsWith('<?') || tag.startsWith('<!')) continue
    if (tag.startsWith('</')) {
      depth -= 1
      if (depth === 0) rootClosed = true
      continue
    }
    if (depth === 0 && rootClosed) extra += 1
    if (!tag.endsWith('/>')) depth += 1
  }
  assert.equal(extra, 0, `${label} 出现了第二个根元素（Excel 会判定文件损坏）`)
  assert.equal(depth, 0, `${label} 标签未配平`)
}

test('★ 创建表格会写齐四处接线（少一处 Excel 就要求修复）', () => {
  const wb = Workbook.open(original)
  const info = wb.createTable('销售数据', { ref: 'A1:D5', name: '销售表', style: 'TableStyleMedium9' })
  assert.equal(info.display_name, '销售表')
  assert.deepEqual(info.columns, ['产品', '单价', '数量', '金额'])
  assert.equal(info.header_fixes.length, 0)
  const out = wb.save()

  const reopened = Workbook.open(out)
  const tables = reopened.tables()
  assert.equal(tables.length, 1)
  assert.equal(tables[0].displayName, '销售表')
  assert.equal(tables[0].ref, 'A1:D5')
  assert.equal(tables[0].style, 'TableStyleMedium9')
  assert.deepEqual(tables[0].columns, ['产品', '单价', '数量', '金额'])

  const pkg = ZipPackage.open(out)
  assert.ok(pkg.has('xl/tables/table1.xml'), '表格部件存在')
  assert.match(pkg.readText('xl/worksheets/_rels/sheet1.xml.rels'), /relationships\/table/, '工作表关系指向表格')
  assert.match(pkg.readText('[Content_Types].xml'), /spreadsheetml\.table\+xml/, '内容类型 Override 存在')
  assert.match(pkg.readText('xl/worksheets/sheet1.xml'), /<tableParts count="1"><tablePart r:id="rId1"\/><\/tableParts>/, '工作表里有 tableParts')
})

test('★ 回归：生成的 XML 必须是单一根元素（曾把 tableParts 挂到文档节点，Excel 打不开）', () => {
  const wb = Workbook.open(original)
  wb.createTable('销售数据', { ref: 'A1:D5' })
  const pkg = ZipPackage.open(wb.save())
  for (const part of ['xl/worksheets/sheet1.xml', 'xl/tables/table1.xml', '[Content_Types].xml', 'xl/worksheets/_rels/sheet1.xml.rels']) {
    assertSingleRoot(pkg.readText(part), part)
  }
})

test('★ 表头为空或重复时写回生成的列名', () => {
  const wb = Workbook.open(original)
  wb.writeCells('销售数据', [
    { ref: 'A1', value: '' },
    { ref: 'B1', value: '' }
  ])
  const info = wb.createTable('销售数据', { ref: 'A1:D2', name: '修正表' })
  assert.deepEqual(info.columns, ['列1', '列2', '数量', '金额'])
  assert.equal(info.header_fixes.length, 2)
  const reopened = Workbook.open(wb.save())
  const header = reopened.readRange('销售数据', 'A1:D1').rows[0].map((c) => c?.value ?? null)
  assert.deepEqual(header, ['列1', '列2', '数量', '金额'], '生成的列名必须写回表头，Excel 才不会报修复')
})

test('★ 表头重复时只修正重复的那一列', () => {
  const wb = Workbook.open(original)
  wb.writeCells('销售数据', [{ ref: 'B1', value: '数量' }])
  const info = wb.createTable('销售数据', { ref: 'A1:D2', name: '重复表' })
  assert.deepEqual(info.columns, ['产品', '数量', '列3', '金额'])
  assert.equal(info.header_fixes.length, 1)
})

test('表格名与区域冲突会被拒绝', () => {
  const wb = Workbook.open(original)
  wb.createTable('销售数据', { ref: 'A1:D5', name: '表一' })
  assert.throws(() => wb.createTable('销售数据', { ref: 'A1:D3', name: '表二' }), (e) => e.code === 'INVALID_REQUEST' && /重叠/.test(e.message))
  assert.throws(() => wb.createTable('销售数据', { ref: 'F1:G3', name: '表一' }), (e) => e.code === 'INVALID_REQUEST' && /已存在/.test(e.message))
  assert.throws(() => wb.createTable('销售数据', { ref: 'F1:G3', name: 'A1' }), (e) => e.code === 'INVALID_REQUEST' && /单元格引用/.test(e.message))
  assert.throws(() => wb.createTable('销售数据', { ref: 'A1' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => wb.createTable('销售数据', { ref: 'F1:G3', style: 'NotAStyle' }), (e) => e.code === 'INVALID_REQUEST')
})

test('★ 第二张表能共存，且各自的 tableParts 计数正确', () => {
  const wb = Workbook.open(original)
  wb.createTable('销售数据', { ref: 'A1:D5', name: '表甲' })
  wb.createTable('销售数据', { ref: 'F1:H4', name: '表乙' })
  const out = wb.save()
  const reopened = Workbook.open(out)
  assert.equal(reopened.tables().length, 2)
  const sheetXml = ZipPackage.open(out).readText('xl/worksheets/sheet1.xml')
  assert.match(sheetXml, /<tableParts count="2">/)
  assertSingleRoot(sheetXml, 'sheet1.xml')
})

test('含汇总行时区域必须包含汇总行', () => {
  const wb = Workbook.open(original)
  assert.throws(() => wb.createTable('销售数据', { ref: 'A1:D2', showTotals: true }), (e) => e.code === 'INVALID_REQUEST')
  const info = wb.createTable('销售数据', { ref: 'A1:D5', showTotals: true, name: '带汇总' })
  assert.equal(info.show_totals, true)
  assert.equal(Workbook.open(wb.save()).tables()[0].totalsRowCount, 1)
})

console.log('\n=== 8. 图表 ===')

const chartOptions = {
  type: 'column',
  title: '各产品金额',
  categories: 'A2:A4',
  series: [{ values: 'D2:D4', nameRef: 'D1' }]
}

test('★ 创建图表会写齐五处接线（少一处 Excel 就要求修复）', () => {
  const wb = Workbook.open(original)
  const info = wb.createChart('销售数据', chartOptions)
  assert.equal(info.chart_part, 'xl/charts/chart1.xml')
  assert.equal(info.drawing_part, 'xl/drawings/drawing1.xml')
  assert.equal(info.series_count, 1)
  const out = wb.save()

  const pkg = ZipPackage.open(out)
  assert.ok(pkg.has('xl/charts/chart1.xml'), '图表部件存在')
  assert.ok(pkg.has('xl/drawings/drawing1.xml'), '绘图部件存在')
  assert.match(pkg.readText('xl/worksheets/sheet1.xml'), /<drawing r:id="rId1"\/>/, '工作表引用绘图')
  assert.match(pkg.readText('xl/worksheets/_rels/sheet1.xml.rels'), /relationships\/drawing/, '工作表关系指向绘图')
  assert.match(pkg.readText('xl/drawings/_rels/drawing1.xml.rels'), /relationships\/chart/, '绘图关系指向图表')
  const contentTypes = pkg.readText('[Content_Types].xml')
  assert.match(contentTypes, /drawingml\.chart\+xml/, '图表内容类型')
  assert.match(contentTypes, /drawing\+xml/, '绘图内容类型')
})

test('★ 图表 XML 的结构与缓存值都写对了', () => {
  const wb = Workbook.open(original)
  wb.createChart('销售数据', chartOptions)
  const chart = ZipPackage.open(wb.save()).readText('xl/charts/chart1.xml')
  assert.match(chart, /<c:barChart><c:barDir val="col"\/><c:grouping val="clustered"\/>/, '柱状簇图表组')
  assert.match(chart, /<a:t>各产品金额<\/a:t>/, '标题写进图表')
  assert.match(chart, /<c:f>销售数据!\$D\$2:\$D\$4<\/c:f>/, '数值引用公式')
  assert.match(chart, /<c:f>销售数据!\$A\$2:\$A\$4<\/c:f>/, '分类引用公式')
  assert.match(chart, /<c:numCache><c:formatCode>General<\/c:formatCode><c:ptCount val="3"\/>/, '数值缓存')
  assert.match(chart, /<c:v>笔记本<\/c:v><\/c:pt>/, '分类文本缓存')
  assert.match(chart, /<c:f>销售数据!\$D\$1:\$D\$1<\/c:f>/, '系列名取自单元格时写成引用')
  assert.ok(chart.indexOf('<c:axId') > chart.indexOf('<c:ser>'), '轴 id 必须排在系列之后（schema 顺序）')
})

test('★ 回归：图表相关 XML 必须是单一根元素', () => {
  const wb = Workbook.open(original)
  wb.createChart('销售数据', chartOptions)
  const pkg = ZipPackage.open(wb.save())
  for (const part of ['xl/charts/chart1.xml', 'xl/drawings/drawing1.xml', 'xl/worksheets/sheet1.xml', 'xl/drawings/_rels/drawing1.xml.rels']) {
    assertSingleRoot(pkg.readText(part), part)
  }
})

test('四种图表类型都能生成，且轴只出现在需要轴的类型里', () => {
  for (const type of ['column', 'bar', 'line', 'pie']) {
    const wb = Workbook.open(original)
    wb.createChart('销售数据', { ...chartOptions, type, title: `${type} 图` })
    const chart = ZipPackage.open(wb.save()).readText('xl/charts/chart1.xml')
    const hasAxes = /<c:catAx>/.test(chart)
    assert.equal(hasAxes, type !== 'pie', `${type} 的坐标轴判定不对`)
    if (type === 'pie') assert.match(chart, /<c:pieChart><c:varyColors val="1"\/>/)
    if (type === 'bar') assert.match(chart, /<c:barDir val="bar"\/>/)
    if (type === 'line') assert.match(chart, /<c:lineChart><c:grouping val="standard"\/>/)
  }
})

test('★ 同一张表加第二个图表会复用绘图部件并追加锚点', () => {
  const wb = Workbook.open(original)
  const first = wb.createChart('销售数据', chartOptions)
  const second = wb.createChart('销售数据', { ...chartOptions, type: 'line', anchor: 'F20', title: '趋势' })
  assert.equal(first.drawing_part, second.drawing_part, '应复用同一个绘图部件')
  assert.equal(second.chart_part, 'xl/charts/chart2.xml')
  const pkg = ZipPackage.open(wb.save())
  const drawing = pkg.readText('xl/drawings/drawing1.xml')
  assert.equal((drawing.match(/<xdr:twoCellAnchor>/g) ?? []).length, 2, '绘图里应有两个锚点')
  assert.match(pkg.readText('xl/drawings/_rels/drawing1.xml.rels'), /chart1\.xml[\s\S]*chart2\.xml|chart2\.xml[\s\S]*chart1\.xml/, '两个图表关系都在')
  assert.match(pkg.readText('xl/worksheets/sheet1.xml'), /<drawing r:id="rId1"\/>/, '工作表只引用一个绘图部件')
})

test('图表参数不合法时给出明确错误', () => {
  const wb = Workbook.open(original)
  assert.throws(() => wb.createChart('销售数据', { ...chartOptions, type: 'radar' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => wb.createChart('销售数据', { ...chartOptions, series: [] }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(
    () => wb.createChart('销售数据', { ...chartOptions, type: 'pie', series: [{ values: 'D2:D4' }, { values: 'C2:C4' }] }),
    (e) => e.code === 'INVALID_REQUEST' && /饼图/.test(e.message)
  )
  assert.throws(() => wb.createChart('销售数据', { ...chartOptions, series: [{ values: 'D2' }] }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(
    () => wb.createChart('销售数据', { ...chartOptions, series: [{ values: '不存在表!D2:D4' }] }),
    (e) => e.code === 'FILE_NOT_FOUND'
  )
})

test('★ 跨表引用数据源', () => {
  const wb = Workbook.open(original)
  const info = wb.createChart('汇总', { ...chartOptions, categories: '销售数据!A2:A4', series: [{ values: '销售数据!D2:D4' }], dataSheet: '销售数据' })
  assert.equal(info.data_sheet, '销售数据')
  const chart = ZipPackage.open(wb.save()).readText('xl/charts/chart1.xml')
  assert.match(chart, /<c:f>销售数据!\$D\$2:\$D\$4<\/c:f>/)
})

console.log('\n=== 9. 安全与错误处理 ===')

test('加密的 OOXML（OLE 容器）给出 PASSWORD_REQUIRED', () => {
  const ole = Buffer.alloc(64)
  ole.writeUInt32LE(0xe011cfd0, 0)
  assert.throws(() => Workbook.open(ole), (e) => e.code === 'PASSWORD_REQUIRED')
})

test('非 xlsx 的 ZIP 给出 CORRUPTED_DOCUMENT', () => {
  const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { maxEntries: 10, maxEntryBytes: 1000, maxTotalBytes: 10000, maxRatio: 200 })
  pkg.write('hello.txt', 'hi')
  assert.throws(() => Workbook.open(pkg.toBuffer()), (e) => e.code === 'CORRUPTED_DOCUMENT')
})

test('访问不存在的工作表给出可用列表', () => {
  const wb = Workbook.open(original)
  try {
    wb.sheet('不存在')
    assert.fail('应当抛出')
  } catch (err) {
    assert.equal(err.code, 'FILE_NOT_FOUND')
    assert.deepEqual(err.details.available, ['销售数据', '汇总'])
  }
})

console.log('\n=== 10. 输出可再次打开（保存后重新验证）===')

test('多次修改-保存-重开循环保持稳定', () => {
  let buffer = original
  for (let i = 0; i < 5; i += 1) {
    const wb = Workbook.open(buffer)
    wb.writeCells('销售数据', [{ ref: 'B2', value: 1000 + i }])
    buffer = wb.save()
  }
  const final = Workbook.open(buffer)
  assert.equal(final.readRange('销售数据', 'B2:B2').rows[0][0].value, 1004)
  assert.deepEqual(final.sheetNames(), ['销售数据', '汇总'])
})

writeFileSync(join(fixtureDir, 'sales-roundtrip.xlsx'), (() => {
  const wb = Workbook.open(original)
  wb.writeCells('销售数据', [{ ref: 'B2', value: 6888 }, { ref: 'A2', value: '笔记本电脑' }])
  return wb.save()
})())

console.log('\n=== 11. 样式 ===')

test('设置加粗与字色时既有字体定义保持原样', () => {
  const wb = Workbook.open(original)
  const beforeStyles = ZipPackage.open(original).readText('xl/styles.xml')
  const start = beforeStyles.indexOf('<font>')
  const originalFont = beforeStyles.slice(start, beforeStyles.indexOf('</font>', start) + '</font>'.length)
  wb.applyCellStyle('销售数据', [{ ref: 'A1', bold: true, fontColor: '#FF0000' }])
  const afterStyles = ZipPackage.open(wb.save()).readText('xl/styles.xml')
  assert.ok(afterStyles.includes(originalFont), `既有 <font> 定义被改动：${originalFont}`)
  assert.ok(afterStyles.includes('<b/>'))
  assert.ok(afterStyles.includes('<color rgb="FFFF0000"/>'))
})

test('设置填充底色、对齐与内置数字格式', () => {
  const wb = Workbook.open(original)
  wb.applyCellStyle('销售数据', [{ ref: 'B2', fillColor: '#FFFF00', numberFormat: '#,##0.00', horizontal: 'center' }])
  const styles = ZipPackage.open(wb.save()).readText('xl/styles.xml')
  assert.ok(styles.includes('patternType="solid"'), '缺少 solid 填充')
  assert.ok(styles.includes('<fgColor rgb="FFFFFF00"/>'), '填充色不正确')
  assert.ok(styles.includes('<alignment horizontal="center"/>'), '缺少对齐')
  assert.ok(styles.includes('numFmtId="4"'), '内置格式 #,##0.00 应映射到 numFmtId=4')
})

test('自定义数字格式登记到 numFmts（从 164 起）', () => {
  const wb = Workbook.open(original)
  wb.applyCellStyle('销售数据', [{ ref: 'B2', numberFormat: '0.0000"元"' }])
  const styles = ZipPackage.open(wb.save()).readText('xl/styles.xml')
  assert.ok(styles.includes('<numFmts'), '未创建 numFmts 容器')
  assert.ok(styles.includes('numFmtId="164"'), '自定义格式应从 164 起编号')
  assert.ok(styles.includes('formatCode="0.0000&quot;元&quot;"'), '格式代码被错误转义或丢失')
})

test('样式只作用于目标单元格，相邻单元格不受影响', () => {
  const wb = Workbook.open(original)
  wb.applyCellStyle('销售数据', [{ ref: 'A2', bold: true }])
  const reopened = Workbook.open(wb.save())
  const cells = reopened.readSheet('销售数据').cells
  const a2 = cells.find((c) => c.ref === 'A2')
  const a3 = cells.find((c) => c.ref === 'A3')
  assert.ok(a2.style >= 1, `A2 应指向新样式，实际 ${a2.style}`)
  assert.equal(a3.style, 0, `A3 样式不应被改动，实际 ${a3.style}`)
  // 数据本身没有被样式操作破坏
  assert.equal(a2.value, '笔记本')
  assert.equal(a3.value, '显示器')
})

test('同一次调用中相同样式只追加一份定义', () => {
  const wb = Workbook.open(original)
  wb.applyCellStyle('销售数据', [
    { ref: 'A2', bold: true },
    { ref: 'A3', bold: true },
    { ref: 'A4', bold: true }
  ])
  const styles = ZipPackage.open(wb.save()).readText('xl/styles.xml')
  assert.equal((styles.match(/<b\/>/g) ?? []).length, 1, '相同字体定义被重复追加')
})

test('对不存在的单元格设置样式会报错而不是凭空创建', () => {
  const wb = Workbook.open(original)
  assert.throws(() => wb.applyCellStyle('销售数据', [{ ref: 'Z99', bold: true }]), (e) => e.code === 'FILE_NOT_FOUND')
})

test('非法颜色被拒绝', () => {
  const wb = Workbook.open(original)
  assert.throws(() => wb.applyCellStyle('销售数据', [{ ref: 'A1', fontColor: 'red' }]), (e) => e.code === 'INVALID_REQUEST')
})

test('样式修改后文件仍可正常重新打开', () => {
  const wb = Workbook.open(original)
  wb.applyCellStyle('销售数据', [{ ref: 'A1', bold: true, fillColor: '#00FF00', fontSize: 14, fontName: '宋体' }])
  const reopened = Workbook.open(wb.save())
  assert.deepEqual(reopened.sheetNames(), ['销售数据', '汇总'])
  assert.equal(reopened.readRange('销售数据', 'A1:A1').rows[0][0].value, '产品')
})

test('★ 回归：不同样式不会指向同一份字体定义', () => {
  const wb = Workbook.open(original)
  wb.applyCellStyle('销售数据', [
    { ref: 'A1', bold: true, fontColor: '#FF0000', fontSize: 14 },
    { ref: 'A2', italic: true, underline: true, fontName: 'Arial' },
    { ref: 'A3', bold: true, fontColor: '#FF0000', fontSize: 14 }
  ])
  const out = wb.save()
  const cells = Workbook.open(out).readSheet('销售数据').cells
  const a1 = cells.find((c) => c.ref === 'A1')
  const a2 = cells.find((c) => c.ref === 'A2')
  const a3 = cells.find((c) => c.ref === 'A3')
  assert.equal(a1.style, a3.style, '相同样式应复用同一份定义')
  assert.notEqual(a1.style, a2.style, '不同样式不应共用同一份定义')
  const fontCount = findAll(XmlDoc.parse(ZipPackage.open(out).readText('xl/styles.xml')).root, 'font').length
  assert.equal(fontCount, 3, `应为「1 个原有 + 2 个新增」共 3 份字体定义，实际 ${fontCount}`)
})

test('★ 回归：字体子元素顺序符合 ECMA-376 sequence', () => {
  const wb = Workbook.open(original)
  wb.applyCellStyle('销售数据', [{ ref: 'A1', bold: true, italic: true, underline: true, fontSize: 12, fontColor: '#00FF00', fontName: 'Arial' }])
  const styles = ZipPackage.open(wb.save()).readText('xl/styles.xml')
  assert.match(styles, /<font><b\/><i\/><u\/><sz val="12"\/><color rgb="FF00FF00"\/><name val="Arial"\/><\/font>/)
})

test('★ 回归：同一次会话内先写入再设样式，能找到刚写入的单元格', () => {
  const wb = Workbook.open(original)
  wb.writeCells('销售数据', [{ ref: 'H5', value: '新写的' }])
  wb.applyCellStyle('销售数据', [{ ref: 'H5', bold: true }])
  const reopened = Workbook.open(wb.save())
  const cell = reopened.readSheet('销售数据').cells.find((c) => c.ref === 'H5')
  assert.ok(cell, 'H5 应存在')
  assert.equal(cell.value, '新写的')
  assert.ok(cell.style >= 1, `H5 应带样式，实际 style=${cell.style}`)
})

test('★ 回归：同一次会话内连续增删改工作表不会写坏 workbook.xml', () => {
  const wb = Workbook.open(original)
  wb.addWorksheet('T1')
  wb.addWorksheet('T2')
  wb.renameWorksheet('T2', 'T2改')
  wb.deleteWorksheet('T1')
  const out = wb.save()
  const reopened = Workbook.open(out)
  assert.deepEqual(reopened.sheetNames(), ['销售数据', '汇总', 'T2改'])
  assert.ok(ZipPackage.open(out).readText('xl/workbook.xml').includes('T2改'))
})

test('★ 回归：同一次会话内写入后立即查找替换', () => {
  const wb = Workbook.open(original)
  wb.writeCells('销售数据', [{ ref: 'H7', value: '待替换词' }])
  const result = wb.findAndReplace({ find: '待替换词', replace: '已替换' })
  assert.equal(result.replacements, 1)
  const reopened = Workbook.open(wb.save())
  assert.equal(reopened.readRange('销售数据', 'H7:H7').rows[0][0].value, '已替换')
})

console.log(`\n结果：${passed} 通过，${failed} 失败\n`)
process.exit(failed === 0 ? 0 : 1)
