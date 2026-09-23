/**
 * 引擎级自检：ZIP 容器往返 + XML 最小修改精度 + 安全防护。
 * 运行：node test/engine.test.mjs
 */
import assert from 'node:assert/strict'
import { ZipPackage, XmlDoc, findAll, find, attr, decodeEntities, escapeXmlAttr } from '../lib/ooxml.js'
import { OfficeError } from '../lib/errors.js'

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

console.log('\n=== 1. ZIP 容器 ===')

test('写入后可重新打开并逐条读回', () => {
  const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { maxEntries: 100, maxEntryBytes: 1e6, maxTotalBytes: 1e7, maxRatio: 200 })
  pkg.write('a.txt', 'hello 世界')
  pkg.write('dir/b.xml', '<x>1</x>')
  const buf = pkg.toBuffer()

  const reopened = ZipPackage.open(buf)
  assert.deepEqual(reopened.names().sort(), ['a.txt', 'dir/b.xml'])
  assert.equal(reopened.readText('a.txt'), 'hello 世界')
  assert.equal(reopened.readText('dir/b.xml'), '<x>1</x>')
})

test('未修改条目在重写后逐字节保持不变', () => {
  const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { maxEntries: 100, maxEntryBytes: 1e6, maxTotalBytes: 1e7, maxRatio: 200 })
  pkg.write('keep.bin', Buffer.from([1, 2, 3, 4, 5]))
  pkg.write('edit.txt', 'v1')
  const first = pkg.toBuffer()

  const second = ZipPackage.open(first)
  const rawBefore = second.read('keep.bin')
  second.write('edit.txt', 'v2')
  const third = ZipPackage.open(second.toBuffer())
  assert.deepEqual(third.read('keep.bin'), rawBefore)
  assert.equal(third.readText('edit.txt'), 'v2')
})

test('删除条目生效', () => {
  const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { maxEntries: 100, maxEntryBytes: 1e6, maxTotalBytes: 1e7, maxRatio: 200 })
  pkg.write('a.txt', 'a')
  pkg.write('b.txt', 'b')
  pkg.delete('a.txt')
  assert.deepEqual(ZipPackage.open(pkg.toBuffer()).names(), ['b.txt'])
})

test('非 ZIP 输入被拒绝', () => {
  assert.throws(() => ZipPackage.open(Buffer.from('not a zip at all, definitely not')), (e) => e.code === 'CORRUPTED_DOCUMENT')
})

test('条目数超限被拒绝（防压缩炸弹）', () => {
  const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { maxEntries: 2, maxEntryBytes: 1e6, maxTotalBytes: 1e7, maxRatio: 200 })
  pkg.write('a', 'a')
  pkg.write('b', 'b')
  pkg.write('c', 'c')
  assert.throws(() => ZipPackage.open(pkg.toBuffer(), { maxEntries: 2 }), (e) => e.code === 'MEMORY_LIMIT')
})

console.log('\n=== 2. XML 最小修改 ===')

const SRC = '<?xml version="1.0"?>\n<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>10</v></c></row><row r="2"><c r="A2"><v>20</v></c></row></sheetData><other attr="untouched">KEEP</other></worksheet>'

test('解析出正确的元素与属性', () => {
  const doc = XmlDoc.parse(SRC)
  const rows = findAll(doc.root, 'row')
  assert.equal(rows.length, 2)
  assert.equal(attr(rows[0], 'r'), '1')
  const cells = findAll(doc.root, 'c')
  assert.equal(cells.length, 3)
  assert.equal(attr(cells[0], 'r'), 'A1')
  assert.equal(attr(cells[0], 't'), 's')
  assert.equal(doc.text(find(cells[0], 'v')), '0')
})

test('改单元格值只改目标区间，其余字节完全不变', () => {
  const doc = XmlDoc.parse(SRC)
  const cells = findAll(doc.root, 'c')
  const vNode = find(cells[1], 'v')
  doc.setText(vNode, '99')
  const out = doc.toString()
  // 仅 <v>10</v> 变成 <v>99</v>
  assert.ok(out.includes('<c r="B1"><v>99</v></c>'))
  assert.ok(out.includes('<other attr="untouched">KEEP</other>'))
  assert.ok(out.startsWith('<?xml version="1.0"?>'))
  // 未触碰区域逐字节一致
  const before = SRC.slice(SRC.indexOf('<row r="2"'))
  const after = out.slice(out.indexOf('<row r="2"'))
  assert.equal(after, before)
})

test('设置属性值不改变标签其它部分', () => {
  const doc = XmlDoc.parse(SRC)
  const cell = findAll(doc.root, 'c')[1]
  doc.setAttr(cell, 't', 'n')
  assert.ok(doc.toString().includes('<c r="B1" t="n"><v>10</v></c>'))
})

test('新增属性插入到起始标签内', () => {
  const doc = XmlDoc.parse('<c r="A1"><v>1</v></c>')
  const c = find(doc.root, 'c')
  doc.setAttr(c, 's', '3')
  assert.equal(doc.toString(), '<c r="A1" s="3"><v>1</v></c>')
})

test('对同一新增属性重复赋值只保留最后一次', () => {
  const doc = XmlDoc.parse('<c r="A1"/>')
  const c = find(doc.root, 'c')
  doc.setAttr(c, 't', 's')
  doc.setAttr(c, 't', 'n')
  assert.equal(doc.toString(), '<c r="A1" t="n"/>')
})

test('自闭合标签新增属性位置正确', () => {
  const doc = XmlDoc.parse('<row r="5"/>')
  doc.setAttr(find(doc.root, 'row'), 'spans', '1:3')
  assert.equal(doc.toString(), '<row r="5" spans="1:3"/>')
})

test('删除元素与插入元素', () => {
  const doc = XmlDoc.parse('<row r="1"><c r="A1"/><c r="B1"/></row>')
  const cells = findAll(doc.root, 'c')
  doc.remove(cells[0])
  doc.insertAfter(cells[1], '<c r="C1"/>')
  assert.equal(doc.toString(), '<row r="1"><c r="B1"/><c r="C1"/></row>')
})

test('对既有属性重复赋值不产生补丁冲突（回归）', () => {
  const doc = XmlDoc.parse('<sst count="0" uniqueCount="0"><si><t>a</t></si></sst>')
  const root = find(doc.root, 'sst')
  doc.setAttr(root, 'uniqueCount', '1')
  doc.setAttr(root, 'uniqueCount', '2')
  doc.setAttr(root, 'uniqueCount', '3')
  doc.setAttr(root, 'count', '5')
  assert.equal(doc.toString(), '<sst count="5" uniqueCount="3"><si><t>a</t></si></sst>')
})

test('删除本会话新增的属性可撤销插入补丁', () => {
  const doc = XmlDoc.parse('<c r="A1"/>')
  const c = find(doc.root, 'c')
  doc.setAttr(c, 't', 's')
  assert.equal(doc.toString(), '<c r="A1" t="s"/>')
  doc.removeAttr(c, 't')
  assert.equal(doc.toString(), '<c r="A1"/>')
})

test('实体解码正确且未知实体不展开', () => {
  assert.equal(decodeEntities('a&amp;b&lt;c&gt;d'), 'a&b<c>d')
  assert.equal(decodeEntities('&#65;&#x42;'), 'AB')
  assert.equal(decodeEntities('&evil;'), '&evil;')
})

test('属性值中的引号被正确转义', () => {
  assert.equal(escapeXmlAttr('a"b&c'), 'a&quot;b&amp;c')
})

test('补丁区间重叠时报错而不是静默损坏', () => {
  const doc = XmlDoc.parse('<a><b>x</b></a>')
  const a = find(doc.root, 'a')
  doc.setText(a, 'replaced')
  assert.throws(() => doc.setText(find(doc.root, 'b'), 'y'), (e) => e.code === 'INTERNAL_ERROR')
})

console.log('\n=== 3. 安全防护 ===')

test('含 DOCTYPE 的 XML 被拒绝（防实体攻击）', () => {
  assert.throws(
    () => XmlDoc.parse('<?xml version="1.0"?><!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///c:/windows/win.ini">]><a>&xxe;</a>'),
    (e) => e.code === 'EXTERNAL_RESOURCE_BLOCKED'
  )
})

test('标签不匹配的 XML 被拒绝', () => {
  assert.throws(() => XmlDoc.parse('<a><b></c></a>'), (e) => e.code === 'CORRUPTED_DOCUMENT')
})

test('未闭合标签被拒绝', () => {
  assert.throws(() => XmlDoc.parse('<a><b></b>'), (e) => e.code === 'CORRUPTED_DOCUMENT')
})

test('超深嵌套被拒绝（防 XML 炸弹）', () => {
  const deep = '<a>'.repeat(300) + '</a>'.repeat(300)
  assert.throws(() => XmlDoc.parse(deep, { maxDepth: 64 }), (e) => e.code === 'MEMORY_LIMIT')
})

test('超大 XML 被拒绝', () => {
  assert.throws(() => XmlDoc.parse(`<a>${'x'.repeat(5000)}</a>`, { maxBytes: 100 }), (e) => e.code === 'MEMORY_LIMIT')
})

test('OfficeError 携带完整结构化字段', () => {
  const err = new OfficeError('MACRO_DETECTED', '检测到宏')
  const json = err.toJSON()
  assert.equal(json.code, 'MACRO_DETECTED')
  assert.equal(json.retryable, false)
  assert.equal(json.needs_confirmation, true)
  assert.equal(typeof json.solution, 'string')
  assert.ok('needs_rollback' in json && 'partial_output' in json)
})

console.log(`\n结果：${passed} 通过，${failed} 失败\n`)
process.exit(failed === 0 ? 0 : 1)
