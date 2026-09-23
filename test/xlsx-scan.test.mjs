/**
 * 区域读取的字节扫描路径测试（开发要求 §14「只读取用户要求的区域」）。
 *
 * 读取不再先建整张表的 DOM，而是直接扫工作表部件的原始字节、扫到区域末尾就停。
 * 这条路径必须与旧的「整表建 DOM 再挑区域」**结果完全一致**，因此本套件里
 * 保留了一份独立的 DOM 参照实现（只用 XmlDoc 的公开 API），逐单元格比对。
 *
 * 另外用「远处有一个非法引用」的样本**证明早停是真的**：
 * 读前两行时不会碰到第 1000 行，只有整表读取才会撞上它。
 *
 * 运行：node test/xlsx-scan.test.mjs
 */
import assert from 'node:assert/strict'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { MAX_EDITABLE_CELLS, MAX_RANGE_CELLS, MAX_SHEET_CELLS, Workbook, parseRef, serialToIso } from '../lib/xlsx.js'
import { XmlDoc, ZipPackage, attr, findAll, find, decodeEntities } from '../lib/ooxml.js'

const workspace = mkdtempSync(join(tmpdir(), 'dsh-office-scan-'))

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

/** 表头区：六种单元格类型各一个。 */
const HEADER_ROW = `<row r="1" spans="1:6">` +
  `<c r="A1" t="s"><v>0</v></c>` +
  `<c r="B1"><v>42</v></c>` +
  `<c r="C1" s="1"><v>45000</v></c>` +
  `<c r="D1" t="inlineStr"><is><t>行内 &amp; 转义</t></is></c>` +
  `<c r="E1" t="str"><f>CONCATENATE("a","b")</f><v>ab</v></c>` +
  `<c r="F1" t="b"><v>1</v></c>` +
  `</row>`

/** 边界行：错误值、自定义日期格式、公式缓存、自闭合与空值。 */
const EDGE_ROW = `<row r="2">` +
  `<c r="A2" t="e"><v>#DIV/0!</v></c>` +
  `<c r="B2" s="2"><v>45000.5</v></c>` +
  `<c r="C2"><f>B1*2</f><v>84</v></c>` +
  `<c r="D2" t="b"><v>0</v></c>` +
  `<c r="E2"/>` +
  `<c r="F2"><v/></c>` +
  `</row>`

/** 少见的写法：缺 r 的 c、分段行内串、保留空格、CDATA、值里有空白。 */
const ODD_ROW = `<row r="3">` +
  `<c><v>99</v></c>` +
  `<c r="B3" t="inlineStr"><is><r><t>分段</t></r><r><t>文本</t></r></is></c>` +
  `<c r="C3" t="inlineStr"><is><t xml:space="preserve"> 前导空格 </t></is></c>` +
  `<c r="D3" t="s"><v>1</v></c>` +
  `<c r="E3"><v> 7 </v></c>` +
  `<c r="F3" t="inlineStr"><is><t><![CDATA[CDATA & 内容]]></t></is></c>` +
  `</row>`

/** 没有 r 的行：行号必须按隐含序号推算，且单元格位置以自身 r 为准。 */
const IMPLIED_ROW = `<row>` +
  `<c r="A4"><v>1</v></c>` +
  `<c r="B4"><f t="shared" si="0" ref="B4:B5"/><v>10</v></c>` +
  `<c r="C4" s="3" t="s"><v>2</v></c>` +
  `<c r="D4" t="str"><v>结果文本</v></c>` +
  `</row>`

/** 带行高属性的行 + 共享公式的从属格。 */
const STYLED_ROW = `<row r="5" ht="22" customHeight="1">` +
  `<c r="A5"><v>2</v></c>` +
  `<c r="B5"><f t="shared" si="0"/><v>20</v></c>` +
  `<c r="C5" t="s"><v>3</v></c>` +
  `<c r="F5" t="inlineStr"><is><t>最后一个</t></is></c>` +
  `</row>`

/** 末尾两行：越界的共享字符串下标、非数字的值。 */
const TAIL_ROWS =
  `<row r="6" />` +
  `<row r="7">` +
  `<c r="A7" t="s"><v>999</v></c>` +
  `<c r="B7"><v>not-a-number</v></c>` +
  `</row>`

const STYLES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
  `<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>` +
  `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="1"><border/></borders>` +
  `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
  `<cellXfs count="4">` +
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `</cellXfs></styleSheet>`

const SHARED_STRINGS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">` +
  `<si><t>共享字符串一</t></si><si><t>共享字符串二</t></si><si><t>共享文本三</t></si><si><t>共享文本四</t></si>` +
  `</sst>`

/**
 * 拼一个最小的 .xlsx，工作表内容与部件内容都可控。
 * @param {object} options - 选项。
 * @param {string} options.sheetData - sheetData 的内容。
 * @param {string} [options.afterSheetData] - `</sheetData>` 之后的内容。
 * @param {boolean} [options.closeSheetData] - 是否写 `</sheetData>`（默认写）。
 * @param {boolean} [options.sharedStrings] - 是否带共享字符串部件。
 * @returns {Buffer} .xlsx 字节。
 */
function buildFixture({ sheetData, afterSheetData = '<sheetProtection sheet="1" objects="0"/>', closeSheetData = true, sharedStrings = true }) {
  const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { maxEntries: 64, maxEntryBytes: 1 << 24, maxTotalBytes: 1 << 26, maxRatio: 200 })
  const overrides =
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
    `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
    (sharedStrings ? `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` : '')
  pkg.write(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`
  )
  pkg.write(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
  )
  pkg.write(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="数据" sheetId="1" r:id="rId1"/></sheets></workbook>`
  )
  pkg.write(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${sharedStrings ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' : ''}</Relationships>`
  )
  pkg.write(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:F7"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="12"/></cols><sheetData>${sheetData}${closeSheetData ? '</sheetData>' : ''}${afterSheetData}</worksheet>`
  )
  pkg.write('xl/styles.xml', STYLES_XML)
  if (sharedStrings) pkg.write('xl/sharedStrings.xml', SHARED_STRINGS_XML)
  return pkg.toBuffer()
}

const SHEET_DATA = HEADER_ROW + EDGE_ROW + ODD_ROW + IMPLIED_ROW + STYLED_ROW + TAIL_ROWS
const fixturePath = join(workspace, 'scan.xlsx')
writeFileSync(fixturePath, buildFixture({ sheetData: SHEET_DATA }))

// ───────────────────────── 参照实现（DOM 路径）─────────────────────────

/**
 * 取节点本地名。
 * @param {object} node - 元素节点。
 * @returns {string} 本地名。
 */
function localName(node) {
  return node.name.includes(':') ? node.name.slice(node.name.indexOf(':') + 1) : node.name
}

/**
 * 收集元素的全部文本（与 DOM 路径的 collectText 相同语义）。
 * @param {object} node - 元素节点。
 * @returns {string} 文本。
 */
function collectText(node) {
  let text = ''
  for (const child of node.children ?? []) {
    if (child.type === 'text') text += decodeEntities(child.raw)
    else if (child.type === 'cdata') text += child.raw
    else if (child.type === 'element') text += collectText(child)
  }
  return text
}

/**
 * 独立的参照实现：整表建 DOM，再挑出区域——也就是改造前的那条路径。
 *
 * 它只依赖 XmlDoc 的公开 API，因此可以当作「扫描路径必须复现」的规格。
 *
 * @param {Buffer} buffer - .xlsx 字节。
 * @param {string} sheetPart - 工作表部件名。
 * @param {{start: {col: number, row: number}, end: {col: number, row: number}}|null} bounds - 区域；null 表示整表。
 * @returns {object[]} 单元格列表（文档顺序）。
 */
function domReference(buffer, sheetPart, bounds) {
  const pkg = ZipPackage.open(buffer)
  const doc = XmlDoc.parse(pkg.readText(sheetPart))
  const sstDoc = pkg.has('xl/sharedStrings.xml') ? XmlDoc.parse(pkg.readText('xl/sharedStrings.xml')) : null
  const sstRoot = sstDoc ? sstDoc.root.children.find((c) => c.type === 'element') : null
  const items = sstRoot ? findAll(sstRoot, 'si').map((si) => collectText(si)) : []
  const stylesDoc = XmlDoc.parse(pkg.readText('xl/styles.xml'))
  const cellXfsNode = find(stylesDoc.root, 'cellXfs')
  const cellXfs = (cellXfsNode.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'xf').map((xf) => Number(attr(xf, 'numFmtId') ?? 0))
  const builtinDates = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57])
  const numFmts = new Map(findAll(stylesDoc.root, 'numFmt').map((nf) => [Number(attr(nf, 'numFmtId')), attr(nf, 'formatCode') ?? '']))
  const isDate = (index) => {
    if (cellXfs[index] === undefined) return false
    const numFmtId = cellXfs[index]
    if (builtinDates.has(numFmtId)) return true
    const code = numFmts.get(numFmtId)
    return code ? /[ymdhs]/i.test(code.replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, '')) : false
  }

  const cells = []
  const sheetData = find(doc.root, 'sheetData')
  if (!sheetData) return cells
  let impliedRow = 0
  for (const rowNode of sheetData.children ?? []) {
    if (rowNode.type !== 'element' || localName(rowNode) !== 'row') continue
    const rawRow = attr(rowNode, 'r')
    const rowNumber = rawRow === undefined ? impliedRow + 1 : Number(rawRow)
    impliedRow = rowNumber
    if (bounds && (rowNumber < bounds.start.row || rowNumber > bounds.end.row)) continue
    for (const c of rowNode.children ?? []) {
      if (c.type !== 'element' || localName(c) !== 'c') continue
      const ref = attr(c, 'r')
      if (!ref) continue
      const { col, row } = parseRef(ref)
      if (bounds && (row < bounds.start.row || row > bounds.end.row || col < bounds.start.col || col > bounds.end.col)) continue
      const styleIndex = Number(attr(c, 's') ?? 0)
      const type = attr(c, 't') ?? 'n'
      const fNode = find(c, 'f')
      const vNode = find(c, 'v')
      const isNode = find(c, 'is')
      const formula = fNode ? collectText(fNode) : null
      const valueTextFor = (node) => (node.selfClosing ? '' : doc.text(node))
      let value = null
      let valueType = 'number'
      if (type === 's') {
        const idx = Number(vNode ? valueTextFor(vNode) : -1)
        value = items[idx] ?? null
        valueType = 'string'
      } else if (type === 'inlineStr') {
        value = isNode ? collectText(isNode) : null
        valueType = 'string'
      } else if (type === 'str') {
        value = vNode ? valueTextFor(vNode) : null
        valueType = 'string'
      } else if (type === 'b') {
        value = vNode ? valueTextFor(vNode) === '1' : false
        valueType = 'boolean'
      } else if (type === 'e') {
        value = vNode ? valueTextFor(vNode) : null
        valueType = 'error'
      } else if (vNode) {
        const num = Number(valueTextFor(vNode))
        if (isDate(styleIndex)) {
          value = serialToIso(num, false)
          valueType = 'date'
        } else {
          value = num
          valueType = 'number'
        }
      }
      cells.push({ ref, col, row, value, valueType, formula, style: styleIndex, raw: type })
    }
  }
  return cells
}

console.log(`\n样本：${fixturePath}`)
console.log('参照实现：整表建 DOM 后再挑区域（改造前的路径）\n')

console.log('=== 1. 与 DOM 参照实现逐单元格一致 ===')

const original = readFileSync(fixturePath)

test('readSheet 的结果与 DOM 参照实现完全一致', () => {
  const wb = Workbook.open(original)
  const actual = wb.readSheet('数据').cells
  const expected = domReference(original, 'xl/worksheets/sheet1.xml', null)
  assert.deepEqual(
    actual.map((c) => c.ref),
    expected.map((c) => c.ref),
    `单元格清单不同：扫描路径 ${actual.length} 个，参照实现 ${expected.length} 个`
  )
  assert.deepEqual(actual, expected)
  assert.equal(actual.length, 27)
})

test('readSheet 的行列规模由单元格自身引用算出', () => {
  const wb = Workbook.open(original)
  const content = wb.readSheet('数据')
  assert.equal(content.row_count, 7)
  assert.equal(content.column_count, 6)
})

test('六种单元格类型的取值与类型都对', () => {
  const wb = Workbook.open(original)
  const byRef = new Map(wb.readSheet('数据').cells.map((c) => [c.ref, c]))
  assert.deepEqual(
    [byRef.get('A1').value, byRef.get('A1').valueType, byRef.get('A1').raw],
    ['共享字符串一', 'string', 's']
  )
  assert.deepEqual([byRef.get('B1').value, byRef.get('B1').valueType], [42, 'number'])
  assert.deepEqual([byRef.get('C1').value, byRef.get('C1').valueType], [serialToIso(45000, false), 'date'])
  assert.deepEqual([byRef.get('D1').value, byRef.get('D1').valueType], ['行内 & 转义', 'string'])
  assert.deepEqual([byRef.get('E1').value, byRef.get('E1').formula], ['ab', 'CONCATENATE("a","b")'])
  assert.deepEqual([byRef.get('F1').value, byRef.get('F1').valueType], [true, 'boolean'])
})

test('边界写法：错误值 / 自定义日期 / 公式缓存 / 自闭合 / 空值', () => {
  const wb = Workbook.open(original)
  const byRef = new Map(wb.readSheet('数据').cells.map((c) => [c.ref, c]))
  assert.deepEqual([byRef.get('A2').value, byRef.get('A2').valueType], ['#DIV/0!', 'error'])
  assert.deepEqual([byRef.get('B2').value, byRef.get('B2').valueType], [serialToIso(45000.5, false), 'date'])
  assert.deepEqual([byRef.get('C2').value, byRef.get('C2').formula], [84, 'B1*2'])
  assert.deepEqual([byRef.get('D2').value, byRef.get('D2').valueType], [false, 'boolean'])
  assert.deepEqual([byRef.get('E2').value, byRef.get('E2').formula], [null, null])
  // `<c><v/></c>`：空的 v 与改造前的 DOM 路径一样按 `Number('')` 处理得到 0。
  // 这是**刻意保留**的既有行为（本次改造的目标就是结果逐字节等价），
  // 已记在任务清单的已知限制里；要改成 null 属于语义变更，需单独决策。
  assert.deepEqual([byRef.get('F2').value, byRef.get('F2').formula], [0, null])
})

test('少见写法：缺 r 的 c 被跳过、分段行内串、CDATA、值里的空白', () => {
  const wb = Workbook.open(original)
  const cells = wb.readSheet('数据').cells
  assert.equal(cells.some((c) => c.value === 99), false, '缺 r 的单元格必须被跳过（与改造前一致）')
  const byRef = new Map(cells.map((c) => [c.ref, c]))
  assert.equal(byRef.get('B3').value, '分段文本')
  assert.equal(byRef.get('C3').value, ' 前导空格 ')
  assert.equal(byRef.get('E3').value, 7)
  assert.equal(byRef.get('F3').value, 'CDATA & 内容')
})

test('缺 r 的行按隐含序号推进，共享公式的从属格公式为空串', () => {
  const wb = Workbook.open(original)
  const byRef = new Map(wb.readSheet('数据').cells.map((c) => [c.ref, c]))
  assert.equal(byRef.get('A4').row, 4)
  assert.equal(byRef.get('A4').value, 1)
  assert.equal(byRef.get('B4').formula, '', '共享公式主格没有公式文本')
  assert.equal(byRef.get('B5').formula, '', '共享公式从属格同样没有公式文本')
  assert.equal(byRef.get('B5').value, 20)
  assert.equal(byRef.get('F5').value, '最后一个')
})

test('越界的共享字符串下标与非法数字如实返回', () => {
  const wb = Workbook.open(original)
  const byRef = new Map(wb.readSheet('数据').cells.map((c) => [c.ref, c]))
  assert.deepEqual([byRef.get('A7').value, byRef.get('A7').valueType], [null, 'string'])
  assert.equal(Number.isNaN(byRef.get('B7').value), true)
})

test('自闭合行不产生单元格，也不影响后续行号', () => {
  const wb = Workbook.open(original)
  const cells = wb.readSheet('数据').cells
  assert.equal(cells.some((c) => c.row === 6), false)
  assert.equal(cells.some((c) => c.row === 7), true)
})

console.log('\n=== 2. 区域读取与参照实现一致 ===')

const ranges = ['A1:F7', 'A1:A1', 'F1:F1', 'B2:B5', 'C1:C3', 'D4:E5', 'A7:F7', 'E2:F2', 'A3:F3', 'A8:F9', 'C4:D5', 'F7:A1', 'B1:B1']

for (const range of ranges) {
  test(`区域 ${range} 与参照实现一致`, () => {
    const wb = Workbook.open(original)
    const actual = wb.readRange('数据', range)
    const bounds = (() => {
      const [a, b] = range.split(':')
      const start = parseRef(a)
      const end = parseRef(b)
      return {
        start: { col: Math.min(start.col, end.col), row: Math.min(start.row, end.row) },
        end: { col: Math.max(start.col, end.col), row: Math.max(start.row, end.row) }
      }
    })()
    const width = bounds.end.col - bounds.start.col + 1
    const height = bounds.end.row - bounds.start.row + 1
    const expected = Array.from({ length: height }, () => new Array(width).fill(null))
    for (const cell of domReference(original, 'xl/worksheets/sheet1.xml', bounds)) {
      expected[cell.row - bounds.start.row][cell.col - bounds.start.col] = {
        ref: cell.ref,
        value: cell.value,
        type: cell.valueType,
        formula: cell.formula
      }
    }
    assert.deepEqual(actual.rows, expected)
  })
}

test('区域外的单元格一律不返回（B2:B2 不会带出整行）', () => {
  const wb = Workbook.open(original)
  const rows = wb.readRange('数据', 'B2:B2').rows
  assert.equal(rows.length, 1)
  assert.equal(rows[0].length, 1)
  assert.equal(rows[0][0].ref, 'B2')
  const wide = wb.readRange('数据', 'A2:C2').rows[0]
  assert.deepEqual(wide.map((c) => c?.ref ?? null), ['A2', 'B2', 'C2'])
})

test('整表读取会报告工作表保护状态', () => {
  const wb = Workbook.open(original)
  wb.readSheet('数据')
  assert.equal(wb.info.sheetProtection.get('数据'), true)
})

console.log('\n=== 3. 早停是真的（功能证明，不靠计时）===')

// 第 1000 行放一个非法引用：只有真的扫到那里才会报错。
const farRows = []
for (let r = 1; r <= 2000; r += 1) {
  farRows.push(`<row r="${r}"><c r="A${r}"><v>${r}</v></c><c r="B${r}"><v>${r * 2}</v></c></row>`)
}
farRows[999] = `<row r="1000"><c r="???"/><c r="B1000"><v>2000</v></c></row>`
const farPath = join(workspace, 'far.xlsx')
writeFileSync(farPath, buildFixture({ sheetData: farRows.join('') }))
const farBytes = readFileSync(farPath)

test('读前 2 行不会碰到第 1000 行的非法引用（证明扫到区域末尾就停）', () => {
  const wb = Workbook.open(farBytes)
  const rows = wb.readRange('数据', 'A1:B2').rows
  assert.deepEqual(rows.map((row) => row.map((c) => c?.value ?? null)), [[1, 2], [2, 4]])
})

test('反面：整表读取一定会撞上那个非法引用（证明前一用例不是空跑）', () => {
  const wb = Workbook.open(farBytes)
  assert.throws(
    () => wb.readSheet('数据'),
    (err) => err.code === 'INVALID_REQUEST' && /非法单元格引用/.test(err.message)
  )
})

test('区域跨过非法引用所在行时同样报错', () => {
  const wb = Workbook.open(farBytes)
  assert.throws(
    () => wb.readRange('数据', 'A999:B1000'),
    (err) => err.code === 'INVALID_REQUEST'
  )
})

test('读末尾区域时只要行号对得上就没问题（B 列仍然可读）', () => {
  const wb = Workbook.open(farBytes)
  const rows = wb.readRange('数据', 'B2000:B2000').rows
  assert.equal(rows[0][0].value, 4000)
})

console.log('\n=== 4. 上限与可操作错误 ===')

test('区域过大在分配内存之前就被拒绝，且给出可操作建议', () => {
  const wb = Workbook.open(original)
  assert.throws(
    () => wb.readRange('数据', 'A1:XFD1048576'),
    (err) =>
      err.code === 'MEMORY_LIMIT' &&
      /17179869184 个单元格/.test(err.message) &&
      /按行分段读取/.test(err.message) &&
      err.details.max_cells === MAX_RANGE_CELLS
  )
})

test('区域上限可用参数放宽（按需读取大区域）', () => {
  const wb = Workbook.open(original)
  const rows = wb.readRange('数据', 'A1:F7', { maxCells: 100000 }).rows
  assert.equal(rows.length, 7)
})

test('整表读取超过上限时在扫描途中报错并指向区域读取', () => {
  const wb = Workbook.open(original)
  assert.throws(
    () => wb.readSheet('数据', { maxCells: 5 }),
    (err) => err.code === 'MEMORY_LIMIT' && /office_read_range/.test(err.message) && err.details.max_cells === 5
  )
})

test('上限常量自洽：区域上限不超过整表上限', () => {
  assert.equal(Number.isInteger(MAX_SHEET_CELLS) && MAX_SHEET_CELLS > 0, true)
  assert.equal(Number.isInteger(MAX_RANGE_CELLS) && MAX_RANGE_CELLS > 0, true)
  assert.equal(MAX_RANGE_CELLS <= MAX_SHEET_CELLS, true)
})

console.log('\n=== 5. 写入侧的规模闸门（读取不受影响）===')

test('写入超过可编辑上限时明确拒绝，并说清代价与出路', () => {
  const wb = Workbook.open(original)
  wb.maxEditableCells = 10
  assert.throws(
    () => wb.writeCells('数据', [{ ref: 'B2', value: 1 }]),
    (err) =>
      err.code === 'MEMORY_LIMIT' &&
      /可安全编辑的规模/.test(err.message) &&
      /maxEditableCells/.test(err.message) &&
      err.details.max_editable_cells === 10
  )
})

test('闸门只挡写入：同一张表照样能读', () => {
  const wb = Workbook.open(original)
  wb.maxEditableCells = 10
  assert.equal(wb.readRange('数据', 'A1:B2').rows[0][0].value, '共享字符串一')
  assert.equal(wb.readSheet('数据').cells.length, 27)
})

test('把上限调回默认值后写入恢复正常（闸门不是硬编码的）', () => {
  const wb = Workbook.open(original)
  wb.maxEditableCells = 10
  wb.maxEditableCells = 100000
  // C4 的样式不是日期格式，写进去的数字读回来还是数字。
  const changes = wb.writeCells('数据', [{ ref: 'C4', value: 777 }])
  assert.equal(changes.length, 1)
  const reopened = Workbook.open(wb.save())
  assert.equal(reopened.readRange('数据', 'C4:C4').rows[0][0].value, 777)
})

test('默认上限是个正数，且不比读取上限小', () => {
  assert.equal(Number.isInteger(MAX_EDITABLE_CELLS) && MAX_EDITABLE_CELLS > 0, true)
  assert.equal(MAX_EDITABLE_CELLS >= MAX_RANGE_CELLS, true)
})

test('★ 回归：写「行缺 r 属性」的单元格是就地更新，不再多出一个重复行', () => {
  const wb = Workbook.open(original)
  const changes = wb.writeCells('数据', [{ ref: 'A4', value: 555 }])
  assert.equal(changes.length, 1)
  const out = wb.save()
  const xml = ZipPackage.open(out).readText('xl/worksheets/sheet1.xml')
  assert.equal((xml.match(/<row/g) ?? []).length, 7, '行元素个数不变')
  assert.equal((xml.match(/r="A4"/g) ?? []).length, 1, 'A4 在文件里只能出现一次')
  const reopened = Workbook.open(out)
  assert.equal(reopened.readSheet('数据').cells.length, 27, '单元格总数不变')
  assert.equal(reopened.readRange('数据', 'A4:A4').rows[0][0].value, 555)
})

console.log('\n=== 6. 缺少 </sheetData> 的畸形文件不会把后续元素当行 ===')

const noClosePath = join(workspace, 'no-close.xlsx')
writeFileSync(
  noClosePath,
  buildFixture({
    sheetData: `<row r="1"><c r="A1"><v>1</v></c></row>`,
    afterSheetData: `<rowBreaks count="0" manualBreakCount="0"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`,
    closeSheetData: false
  })
)
const noCloseBytes = readFileSync(noClosePath)

test('<rowBreaks> 不会被当成 <row>', () => {
  const wb = Workbook.open(noCloseBytes)
  const cells = wb.readSheet('数据').cells
  assert.deepEqual(cells.map((c) => c.ref), ['A1'])
})

test('<cols> 不会被当成 <c>', () => {
  const wb = Workbook.open(noCloseBytes)
  const rows = wb.readRange('数据', 'A1:B1').rows
  assert.deepEqual(rows[0].map((c) => c?.ref ?? null), ['A1', null])
})

rmSync(workspace, { recursive: true, force: true })

console.log(`\n结果：${passed} 通过，${failed} 失败\n`)
process.exit(failed === 0 ? 0 : 1)
