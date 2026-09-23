/**
 * DOCX 适配器测试（阶段 3 起步：读取与结构识别）。
 * 运行：node test/docx.test.mjs
 */
import assert from 'node:assert/strict'
import { writeFileSync, mkdirSync, readFileSync, existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import zlib from 'node:zlib'
import { DocxDocument, buildBasicDocx, detectImageType, readImageSize } from '../lib/docx.js'
import { ZipPackage } from '../lib/ooxml.js'

const here = dirname(fileURLToPath(import.meta.url))
const fixtureDir = join(here, 'fixtures')
mkdirSync(fixtureDir, { recursive: true })

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

const original = buildBasicDocx({ title: '季度销售报告' })
writeFileSync(join(fixtureDir, 'report.docx'), original)

console.log('\n=== 1. 打开与部件结构 ===')

test('打开自建 DOCX 并识别部件', () => {
  const doc = DocxDocument.open(original)
  assert.equal(doc.info.headers.length, 1)
  assert.equal(doc.info.footers.length, 1)
  assert.equal(doc.info.hasMacro, false)
})

test('拒绝非 DOCX 的 ZIP', () => {
  const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { maxEntries: 8, maxEntryBytes: 1000, maxTotalBytes: 10000, maxRatio: 200 })
  pkg.write('hello.txt', 'hi')
  assert.throws(() => DocxDocument.open(pkg.toBuffer()), (e) => e.code === 'CORRUPTED_DOCUMENT')
})

test('加密的 OOXML（OLE 容器）给出 PASSWORD_REQUIRED', () => {
  const ole = Buffer.alloc(64)
  ole.writeUInt32LE(0xe011cfd0, 0)
  assert.throws(() => DocxDocument.open(ole), (e) => e.code === 'PASSWORD_REQUIRED')
})

console.log('\n=== 2. 元数据 ===')

test('读取核心属性', () => {
  const meta = DocxDocument.open(original).metadata()
  assert.equal(meta.title, '季度销售报告')
  assert.equal(meta.creator, 'dsh-exp-office')
  assert.equal(meta.created, '2026-09-21T00:00:00Z')
})

test('读取应用属性', () => {
  const meta = DocxDocument.open(original).metadata()
  assert.equal(meta.application, 'dsh-exp-office')
  assert.equal(meta.pages, 1)
})

console.log('\n=== 3. 段落 ===')

test('读取段落文本', () => {
  const result = DocxDocument.open(original).paragraphs()
  assert.equal(result.total, 3)
  assert.equal(result.paragraphs[0].text, '季度销售报告')
  assert.equal(result.paragraphs[1].text, '本文件由 dsh-exp-office 生成。加粗片段')
})

test('识别标题样式与大纲级别', () => {
  const first = DocxDocument.open(original).paragraphs().paragraphs[0]
  assert.equal(first.style, 'Heading1')
  assert.equal(first.outline_level, 0)
})

test('识别加粗片段', () => {
  const p = DocxDocument.open(original).paragraphs().paragraphs[1]
  assert.equal(p.has_bold_run, true)
  assert.equal(p.run_count, 2)
})

test('段落分页读取', () => {
  const result = DocxDocument.open(original).paragraphs({ offset: 1, limit: 1 })
  assert.equal(result.returned, 1)
  assert.equal(result.paragraphs[0].index, 1)
})

console.log('\n=== 4. 表格 ===')

test('读取表格结构与单元格文本', () => {
  const result = DocxDocument.open(original).tables()
  assert.equal(result.count, 1)
  const table = result.tables[0]
  assert.equal(table.rows, 2)
  assert.equal(table.columns, 2)
  assert.equal(table.cells[0][0].text, '月份')
  assert.equal(table.cells[1][1].text, '12000')
})

console.log('\n=== 5. 页眉页脚与页码域 ===')

test('读取页眉文本', () => {
  const headers = DocxDocument.open(original).headersOrFooters('header')
  assert.equal(headers.length, 1)
  assert.deepEqual(headers[0].paragraphs, ['季度销售报告'])
})

test('识别页脚中的页码域', () => {
  const footers = DocxDocument.open(original).headersOrFooters('footer')
  assert.equal(footers[0].has_page_number_field, true)
  assert.ok(footers[0].paragraphs[0].includes('第'))
})

console.log('\n=== 6. 结构识别 ===')

test('统计正文结构并识别域代码', () => {
  const s = DocxDocument.open(original).structure()
  assert.equal(s.tables, 1)
  assert.equal(s.headers, 1)
  assert.equal(s.footers, 1)
  assert.equal(s.has_fields, true, '页脚有 PAGE 域，应被识别')
  assert.equal(s.has_comments, false)
  assert.equal(s.has_revisions, false)
  assert.equal(s.has_macro, false)
})

test('读取页面尺寸（twips）', () => {
  const s = DocxDocument.open(original).structure()
  assert.equal(s.section_properties[0].page_width_twips, 11906)
  assert.equal(s.section_properties[0].page_height_twips, 16838)
})

console.log('\n=== 7. 全文提取 ===')

test('提取全文并保留表格行', () => {
  const result = DocxDocument.open(original).text()
  assert.ok(result.text.includes('季度销售报告'))
  assert.ok(result.text.includes('月份\t金额'), `表格行未按制表符输出：${result.text}`)
  assert.equal(result.truncated, false)
})

test('全文提取支持截断', () => {
  const result = DocxDocument.open(original).text(10)
  assert.equal(result.text.length, 10)
  assert.equal(result.truncated, true)
  assert.ok(result.length > 10)
})

console.log('\n=== 8. 安全防护 ===')

test('DOCX 同样受 ZIP 炸弹上限约束', () => {
  const fake = Buffer.concat([original, Buffer.alloc(0)])
  assert.doesNotThrow(() => DocxDocument.open(fake, { maxEntries: 4096 }))
  assert.throws(() => DocxDocument.open(fake, { maxEntries: 2 }), (e) => e.code === 'MEMORY_LIMIT')
})

test('空文件被拒绝', () => {
  assert.throws(() => DocxDocument.open(Buffer.alloc(0)), (e) => e.code === 'CORRUPTED_DOCUMENT')
})

console.log('\n=== 9. 写入：段落 ===')

test('末尾插入段落', () => {
  const doc = DocxDocument.open(original)
  doc.insertParagraph({ after: doc.bodyParagraphs().length - 1, text: '新增的结尾段落' })
  const paragraphs = DocxDocument.open(doc.save()).paragraphs()
  assert.equal(paragraphs.total, 4)
  assert.equal(paragraphs.paragraphs[3].text, '新增的结尾段落')
})

test('最前插入带样式的段落', () => {
  const doc = DocxDocument.open(original)
  doc.insertParagraph({ after: -1, text: '文档副标题', style: 'Heading1' })
  const paragraphs = DocxDocument.open(doc.save()).paragraphs()
  assert.equal(paragraphs.total, 4)
  assert.equal(paragraphs.paragraphs[0].text, '文档副标题')
  assert.equal(paragraphs.paragraphs[0].style, 'Heading1')
  assert.equal(paragraphs.paragraphs[1].text, '季度销售报告')
})

test('插入位置越界被拒绝', () => {
  const doc = DocxDocument.open(original)
  assert.throws(() => doc.insertParagraph({ after: 99, text: 'x' }), (e) => e.code === 'INVALID_REQUEST')
})

test('改写段落保留原样式', () => {
  const doc = DocxDocument.open(original)
  doc.updateParagraph({ index: 0, text: '改过的标题' })
  const paragraphs = DocxDocument.open(doc.save()).paragraphs()
  assert.equal(paragraphs.paragraphs[0].text, '改过的标题')
  assert.equal(paragraphs.paragraphs[0].style, 'Heading1', '段落属性必须保留')
  assert.equal(paragraphs.paragraphs[0].outline_level, 0)
})

test('改写段落替换原内容而非追加', () => {
  const doc = DocxDocument.open(original)
  doc.updateParagraph({ index: 1, text: '只有这一句' })
  const p = DocxDocument.open(doc.save()).paragraphs().paragraphs[1]
  assert.equal(p.text, '只有这一句')
  assert.equal(p.run_count, 1, '旧 run 应被清除')
})

test('多行文本转成软换行', () => {
  const doc = DocxDocument.open(original)
  doc.updateParagraph({ index: 1, text: '第一行\n第二行' })
  const text = DocxDocument.open(doc.save()).text().text
  assert.ok(text.includes('第一行\n第二行'), `软换行未生效：${text}`)
})

test('删除段落', () => {
  const doc = DocxDocument.open(original)
  doc.deleteParagraph({ index: 1 })
  const paragraphs = DocxDocument.open(doc.save()).paragraphs()
  assert.equal(paragraphs.total, 2)
  assert.equal(paragraphs.paragraphs[1].text, '正文结束。')
})

test('删除不存在的段落报错', () => {
  const doc = DocxDocument.open(original)
  assert.throws(() => doc.deleteParagraph({ index: 99 }), (e) => e.code === 'FILE_NOT_FOUND')
})

console.log('\n=== 10. 写入：查找替换 ===')

test('替换单个 run 内的文本', () => {
  const doc = DocxDocument.open(original)
  const result = doc.findAndReplace({ find: '季度', replace: '年度' })
  assert.equal(result.replacements, 1)
  const paragraphs = DocxDocument.open(doc.save()).paragraphs()
  assert.equal(paragraphs.paragraphs[0].text, '年度销售报告')
})

test('替换正文中的英文串', () => {
  const doc = DocxDocument.open(original)
  const result = doc.findAndReplace({ find: 'dsh-exp-office', replace: 'DSH' })
  assert.equal(result.replacements, 1)
  assert.ok(DocxDocument.open(doc.save()).text().text.includes('本文件由 DSH 生成。'))
})

test('默认不触碰页眉页脚', () => {
  const doc = DocxDocument.open(original)
  doc.findAndReplace({ find: '季度', replace: '年度' })
  const headers = DocxDocument.open(doc.save()).headersOrFooters('header')
  assert.deepEqual(headers[0].paragraphs, ['季度销售报告'], '页眉不应被改动')
})

test('可选同时替换页眉页脚', () => {
  const doc = DocxDocument.open(original)
  const result = doc.findAndReplace({ find: '季度', replace: '年度', includeHeaders: true })
  assert.equal(result.replacements, 2)
  const headers = DocxDocument.open(doc.save()).headersOrFooters('header')
  assert.deepEqual(headers[0].paragraphs, ['年度销售报告'])
})

test('区分大小写开关生效', () => {
  assert.equal(DocxDocument.open(original).findAndReplace({ find: 'DSH-EXP-OFFICE', replace: 'X' }).replacements, 1)
  assert.equal(DocxDocument.open(original).findAndReplace({ find: 'DSH-EXP-OFFICE', replace: 'X', matchCase: true }).replacements, 0)
})

test('空查找内容被拒绝', () => {
  assert.throws(() => DocxDocument.open(original).findAndReplace({ find: '', replace: 'x' }), (e) => e.code === 'INVALID_REQUEST')
})

test('被拆散在多个格式片段中的文本给出明确警告而非静默跳过', () => {
  const doc = DocxDocument.open(original)
  doc.updateParagraph({ index: 1, text: '目标词在这里' })
  // 手工把该段文本拆到多个 run，模拟真实模板的格式切分
  const pkg = ZipPackage.open(doc.save())
  const xml = pkg.readText('word/document.xml').replace(
    '<w:t xml:space="preserve">目标词在这里</w:t>',
    '<w:t xml:space="preserve">目</w:t></w:r><w:r><w:t>标词</w:t></w:r><w:r><w:t>在这里</w:t>'
  )
  pkg.write('word/document.xml', xml)
  const result = DocxDocument.open(pkg.toBuffer()).findAndReplace({ find: '目标词', replace: '替换后' })
  assert.equal(result.replacements, 0)
  assert.equal(result.split_run_warnings.length, 1)
  assert.equal(result.split_run_warnings[0].code, 'TEXT_SPLIT_ACROSS_RUNS')
})

console.log('\n=== 11. 写入保真度 ===')

test('★ 只改段落，其余部件逐字节不变', () => {
  const before = new Map()
  const pkgBefore = ZipPackage.open(original)
  for (const name of pkgBefore.names()) before.set(name, pkgBefore.read(name))

  const doc = DocxDocument.open(original)
  doc.updateParagraph({ index: 1, text: '只改这一段' })
  const after = new Map()
  const pkgAfter = ZipPackage.open(doc.save())
  for (const name of pkgAfter.names()) after.set(name, pkgAfter.read(name))

  const changed = [...before.keys()].filter((n) => !after.get(n).equals(before.get(n)))
  assert.deepEqual(changed, ['word/document.xml'], `意外改动的部件：${changed.join(', ')}`)
  assert.ok(after.get('word/header1.xml').equals(before.get('word/header1.xml')), '页眉被改动')
  assert.ok(after.get('word/styles.xml').equals(before.get('word/styles.xml')), '样式定义被改动')
  assert.ok(after.get('docProps/core.xml').equals(before.get('docProps/core.xml')), '文档属性被改动')
})

test('★ 未触碰的 XML 区域逐字节一致', () => {
  const before = ZipPackage.open(original).readText('word/document.xml')
  const doc = DocxDocument.open(original)
  doc.updateParagraph({ index: 1, text: '只改这一段' })
  const after = ZipPackage.open(doc.save()).readText('word/document.xml')
  const cut = (t) => t.slice(t.indexOf('<w:tbl>'))
  assert.equal(cut(after), cut(before), '表格及其后的 XML 被改动')
})

test('写入后文档仍可被自身重新解析', () => {
  const doc = DocxDocument.open(original)
  doc.insertParagraph({ after: 0, text: 'A' })
  doc.updateParagraph({ index: 2, text: 'B' })
  doc.deleteParagraph({ index: 1 })
  const reopened = DocxDocument.open(doc.save())
  assert.equal(reopened.structure().tables, 1)
  assert.equal(reopened.structure().has_page_number_field, true)
  assert.equal(reopened.metadata().title, '季度销售报告')
})

test('连续多次写入保持稳定', () => {
  let buffer = original
  for (let i = 0; i < 5; i += 1) {
    const doc = DocxDocument.open(buffer)
    doc.updateParagraph({ index: 0, text: `第 ${i} 轮` })
    buffer = doc.save()
  }
  const paragraphs = DocxDocument.open(buffer).paragraphs()
  assert.equal(paragraphs.paragraphs[0].text, '第 4 轮')
  assert.equal(paragraphs.total, 3)
})

writeFileSync(join(fixtureDir, 'report-edited.docx'), (() => {
  const doc = DocxDocument.open(original)
  doc.insertParagraph({ after: 0, text: '本段由插件插入', style: 'Heading1' })
  doc.findAndReplace({ find: 'dsh-exp-office', replace: 'DeepSeek Harness' })
  return doc.save()
})())

console.log('\n=== 12. validate_docx ===')

test('自建文档校验通过', () => {
  const report = DocxDocument.open(original).validate()
  assert.equal(report.valid, true, JSON.stringify(report.checks.filter((c) => !c.ok)))
  assert.ok(report.checks.length >= 10)
})

test('校验报告包含必需部件与关系检查', () => {
  const names = DocxDocument.open(original).validate().checks.map((c) => c.name)
  for (const expected of ['ZIP 包结构与中央目录', '必需部件齐全', '正文存在', '关系目标部件存在', '样式引用有定义']) {
    assert.ok(names.includes(expected), `缺少检查项：${expected}`)
  }
})

test('★ 明确列出需要渲染才能判定的项，不假装通过', () => {
  const report = DocxDocument.open(original).validate()
  assert.ok(report.not_checked.length >= 5)
  assert.ok(report.not_checked.some((c) => c.name === '空白页面'))
  assert.ok(report.not_checked.some((c) => c.name === '字体替换'))
  assert.ok(report.not_checked.every((c) => typeof c.reason === 'string' && c.reason.length > 0))
})

test('输出文档使用的字体清单', () => {
  const doc = DocxDocument.open(original)
  doc.updateParagraph({ index: 1, text: '带字体的段落' })
  const report = DocxDocument.open(doc.save()).validate()
  assert.ok(Array.isArray(report.fonts))
})

test('★ 检出悬空的图片/超链接关系', () => {
  const pkg = ZipPackage.open(original)
  const xml = pkg.readText('word/document.xml').replace(
    '<w:p><w:r><w:t>正文结束。</w:t></w:r></w:p>',
    '<w:p><w:hyperlink r:id="rId99"><w:r><w:t>悬空链接</w:t></w:r></w:hyperlink></w:p>'
  )
  pkg.write('word/document.xml', xml)
  const report = DocxDocument.open(pkg.toBuffer()).validate()
  assert.equal(report.valid, false)
  const check = report.checks.find((c) => c.name === '图片/超链接/页眉页脚引用可解析')
  assert.equal(check.ok, false)
  assert.ok(check.detail.includes('rId99'), check.detail)
})

test('★ 检出引用了未定义的样式', () => {
  const pkg = ZipPackage.open(original)
  const xml = pkg.readText('word/document.xml').replace('<w:pStyle w:val="Heading1"/>', '<w:pStyle w:val="不存在的样式"/>')
  pkg.write('word/document.xml', xml)
  const report = DocxDocument.open(pkg.toBuffer()).validate()
  assert.equal(report.valid, false)
  const check = report.checks.find((c) => c.name === '样式引用有定义')
  assert.equal(check.ok, false)
  assert.ok(check.detail.includes('不存在的样式'))
})

test('文档含域代码时给出需人工复核的提示', () => {
  const report = DocxDocument.open(original).validate()
  const check = report.checks.find((c) => c.name === '域代码状态')
  assert.equal(check.warning, true)
  assert.ok(check.detail.includes('F9'))
})

test('写入后的文档仍然校验通过', () => {
  const doc = DocxDocument.open(original)
  doc.insertParagraph({ after: 0, text: '新段', style: 'Heading1' })
  doc.findAndReplace({ find: '季度', replace: '年度' })
  const report = DocxDocument.open(doc.save()).validate()
  assert.equal(report.valid, true, JSON.stringify(report.checks.filter((c) => !c.ok)))
})

console.log('\n=== 13. 页眉页脚与页码 ===')

test('设置页眉文本', () => {
  const doc = DocxDocument.open(original)
  const result = doc.setHeaderFooter({ kind: 'header', text: '新页眉' })
  assert.equal(result.part, 'word/header1.xml')
  const headers = DocxDocument.open(doc.save()).headersOrFooters('header')
  assert.deepEqual(headers[0].paragraphs, ['新页眉'])
})

test('设置页脚并附加页码域', () => {
  const doc = DocxDocument.open(original)
  const result = doc.setHeaderFooter({ kind: 'footer', text: '第 ', pageNumber: true })
  assert.equal(result.page_number_field, true)
  const footers = DocxDocument.open(doc.save()).headersOrFooters('footer')
  assert.equal(footers[0].has_page_number_field, true)
  assert.ok(footers[0].paragraphs[0].startsWith('第'), footers[0].paragraphs[0])
})

test('页码域用 fldChar 三段式而非 fldSimple', () => {
  const doc = DocxDocument.open(original)
  doc.setHeaderFooter({ kind: 'footer', text: '', pageNumber: true })
  const xml = ZipPackage.open(doc.save()).readText('word/footer1.xml')
  assert.ok(xml.includes('w:fldCharType="begin"'))
  assert.ok(xml.includes('w:fldCharType="end"'))
  assert.ok(xml.includes('<w:instrText xml:space="preserve"> PAGE </w:instrText>'))
})

test('★ 页眉页脚写入只改目标部件', () => {
  const before = new Map()
  const pkgBefore = ZipPackage.open(original)
  for (const name of pkgBefore.names()) before.set(name, pkgBefore.read(name))

  const doc = DocxDocument.open(original)
  doc.setHeaderFooter({ kind: 'header', text: '只改页眉' })
  const after = new Map()
  const pkgAfter = ZipPackage.open(doc.save())
  for (const name of pkgAfter.names()) after.set(name, pkgAfter.read(name))

  const changed = [...before.keys()].filter((n) => !after.get(n).equals(before.get(n)))
  assert.deepEqual(changed, ['word/header1.xml'], `意外改动的部件：${changed.join(', ')}`)
})

test('写入页眉后正文段落不受影响', () => {
  const doc = DocxDocument.open(original)
  doc.setHeaderFooter({ kind: 'header', text: '新页眉' })
  const reopened = DocxDocument.open(doc.save())
  assert.equal(reopened.paragraphs().total, 3)
  assert.equal(reopened.paragraphs().paragraphs[0].text, '季度销售报告')
  assert.equal(reopened.structure().tables, 1)
})

test('文档没有页眉时会自动创建（不再报错）', () => {
  const pkg = ZipPackage.open(original)
  pkg.delete('word/header1.xml')
  const doc = DocxDocument.open(pkg.toBuffer())
  const result = doc.setHeaderFooter({ kind: 'header', text: '自动创建' })
  assert.equal(result.created, true)
  assert.equal(DocxDocument.open(doc.save()).headersOrFooters('header')[0].paragraphs[0], '自动创建')
})

console.log('\n=== 14. 表格单元格 ===')

test('改写表格单元格文本', () => {
  const doc = DocxDocument.open(original)
  doc.updateTableCell({ table: 0, row: 1, column: 1, text: '15000' })
  const tables = DocxDocument.open(doc.save()).tables()
  assert.equal(tables.tables[0].cells[1][1].text, '15000')
  assert.equal(tables.tables[0].cells[0][0].text, '月份', '其它单元格不应被改动')
})

test('保留表格结构与其它单元格', () => {
  const doc = DocxDocument.open(original)
  doc.updateTableCell({ table: 0, row: 0, column: 0, text: '季度' })
  const t = DocxDocument.open(doc.save()).tables().tables[0]
  assert.equal(t.rows, 2)
  assert.equal(t.columns, 2)
  assert.equal(t.cells[0][0].text, '季度')
  assert.equal(t.cells[0][1].text, '金额')
})

test('行列越界时给出可用规模', () => {
  const doc = DocxDocument.open(original)
  try {
    doc.updateTableCell({ table: 0, row: 9, column: 0, text: 'x' })
    assert.fail('应当抛出')
  } catch (err) {
    assert.equal(err.code, 'FILE_NOT_FOUND')
    assert.equal(err.details.row_count, 2)
  }
})

test('不存在的表格给出表格总数', () => {
  const doc = DocxDocument.open(original)
  try {
    doc.updateTableCell({ table: 5, row: 0, column: 0, text: 'x' })
    assert.fail('应当抛出')
  } catch (err) {
    assert.equal(err.details.table_count, 1)
  }
})

test('★ 表格写入只改 document.xml', () => {
  const before = new Map()
  const pkgBefore = ZipPackage.open(original)
  for (const n of pkgBefore.names()) before.set(n, pkgBefore.read(n))
  const doc = DocxDocument.open(original)
  doc.updateTableCell({ table: 0, row: 1, column: 0, text: 'X' })
  const pkgAfter = ZipPackage.open(doc.save())
  const changed = [...before.keys()].filter((n) => !pkgAfter.read(n).equals(before.get(n)))
  assert.deepEqual(changed, ['word/document.xml'], `意外改动：${changed.join(', ')}`)
})

console.log('\n=== 15. 表格行与合并 ===')

test('★ 插入行：继承参照行的单元格属性，内容可控', () => {
  const doc = DocxDocument.open(original)
  const result = doc.insertTableRow({ table: 0, after: 0, cells: ['2月', '8000'] })
  assert.equal(result.row, 1)
  assert.equal(result.columns, 2)
  assert.equal(result.row_count, 3)
  const out = doc.save()
  const table = DocxDocument.open(out).tables().tables[0]
  assert.equal(table.rows, 3)
  assert.deepEqual(table.cells.map((row) => row.map((cell) => cell.text)), [
    ['月份', '金额'],
    ['2月', '8000'],
    ['1月', '12000']
  ])
  // 新行必须带单元格属性（否则 Word 会按默认列宽重排整张表）
  const xml = ZipPackage.open(out).readText('word/document.xml')
  const rows = xml.match(/<w:tr>[\s\S]*?<\/w:tr>/g) ?? []
  assert.equal(rows.length, 3)
  assert.match(rows[1], /<w:tcPr>/, '新行应继承 w:tcPr')
  assert.equal((rows[1].match(/<w:tcPr>/g) ?? []).length, 2, '每个单元格都要有 w:tcPr')
})

test('after=-1 插到表格最前；单元格少于列数时其余留空', () => {
  const doc = DocxDocument.open(original)
  doc.insertTableRow({ table: 0, after: -1, cells: ['表头前'] })
  const table = DocxDocument.open(doc.save()).tables().tables[0]
  assert.equal(table.rows, 3)
  assert.deepEqual(table.cells[0].map((cell) => cell.text), ['表头前', ''])
})

test('★ 删除行；只剩一行时拒绝', () => {
  const doc = DocxDocument.open(original)
  const result = doc.deleteTableRow({ table: 0, row: 0 })
  assert.equal(result.row_count, 1)
  assert.equal(result.removed_text, '月份金额')
  const reopened = DocxDocument.open(doc.save())
  assert.equal(reopened.tables().tables[0].rows, 1)
  assert.throws(() => reopened.deleteTableRow({ table: 0, row: 0 }), (err) => err.code === 'INVALID_REQUEST')
})

test('★ 纵向合并：文字并入左上角、原格清空、vMerge 写对', () => {
  const doc = DocxDocument.open(original)
  const result = doc.mergeTableCells({ table: 0, row: 0, column: 0, rowSpan: 2, colSpan: 1 })
  assert.deepEqual(result.merged_text, ['1月'])
  const out = doc.save()
  const table = DocxDocument.open(out).tables().tables[0]
  assert.equal(table.cells[0][0].text, '月份\n1月', '被并入的文字应保留在左上角')
  assert.equal(table.cells[1][0].text, '', '被并入的单元格应清空，避免取消合并时出现重复文字')
  const xml = ZipPackage.open(out).readText('word/document.xml')
  assert.match(xml, /<w:vMerge w:val="restart"\/>/, '首格写 restart')
  assert.match(xml, /<w:vMerge\/>/, '其余格写继续合并')
})

test('★ 横向合并：写 gridSpan 并删掉被并入的单元格', () => {
  const doc = DocxDocument.open(original)
  const result = doc.mergeTableCells({ table: 0, row: 1, column: 0, colSpan: 2 })
  assert.deepEqual(result.merged_text, ['12000'])
  const out = doc.save()
  const xml = ZipPackage.open(out).readText('word/document.xml')
  assert.match(xml, /<w:gridSpan w:val="2"\/>/)
  const rows = xml.match(/<w:tr>[\s\S]*?<\/w:tr>/g) ?? []
  assert.equal((rows[1].match(/<w:tc>/g) ?? []).length, 1, '第二行应只剩一个单元格')
  const table = DocxDocument.open(out).tables().tables[0]
  assert.equal(table.cells[1][0].text, '1月\n12000')
})

test('★ 合并后的 gridSpan 必须排在 tcW 之后、vMerge 之前（schema 顺序）', () => {
  const doc = DocxDocument.open(original)
  doc.mergeTableCells({ table: 0, row: 0, column: 0, rowSpan: 2, colSpan: 2 })
  const xml = ZipPackage.open(doc.save()).readText('word/document.xml')
  const tcPr = /<w:tcPr>([\s\S]*?)<\/w:tcPr>/.exec(xml)
  assert.ok(tcPr, '应有 tcPr')
  const body = tcPr[1]
  const grid = body.indexOf('<w:gridSpan')
  const vmerge = body.indexOf('<w:vMerge')
  assert.ok(grid >= 0 && vmerge > grid, `tcPr 里 gridSpan 应排在 vMerge 之前：${body}`)
})

test('合并参数不合法时拒绝', () => {
  const doc = DocxDocument.open(original)
  assert.throws(() => doc.mergeTableCells({ table: 0, row: 0, column: 0 }), (err) => err.code === 'INVALID_REQUEST')
  assert.throws(() => doc.mergeTableCells({ table: 0, row: 0, column: 0, colSpan: 5 }), (err) => err.code === 'INVALID_REQUEST')
  assert.throws(() => doc.mergeTableCells({ table: 0, row: 1, column: 0, rowSpan: 5 }), (err) => err.code === 'INVALID_REQUEST')
})

test('★ 表格结构操作只改 document.xml', () => {
  const pkgBefore = ZipPackage.open(original)
  const before = new Map(pkgBefore.names().map((n) => [n, pkgBefore.read(n)]))
  const doc = DocxDocument.open(original)
  doc.insertTableRow({ table: 0, after: 0, cells: ['x', 'y'] })
  doc.mergeTableCells({ table: 0, row: 1, column: 0, colSpan: 2 })
  const pkgAfter = ZipPackage.open(doc.save())
  const changed = [...before.keys()].filter((n) => !pkgAfter.read(n).equals(before.get(n)))
  assert.deepEqual(changed, ['word/document.xml'], `意外改动：${changed.join(', ')}`)
})

test('★ 套用内置表格样式：tblStyle 与 tblLook 都写对', () => {
  const doc = DocxDocument.open(original)
  const result = doc.setTableStyle({ table: 0, styleId: 'TableGrid', firstRow: true, bandedRows: true })
  assert.equal(result.style_id, 'TableGrid')
  assert.equal(result.style_defined_in_document, false, '内置样式定义已存在，不应重复创建')
  const xml = ZipPackage.open(doc.save()).readText('word/document.xml')
  const tblPr = /<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(xml)
  assert.ok(tblPr, '应有 tblPr')
  assert.match(tblPr[0], /<w:tblStyle w:val="TableGrid"\/>/)
  assert.match(tblPr[0], /<w:tblLook [^>]*w:firstRow="1"/)
  assert.match(tblPr[0], /<w:tblLook [^>]*w:noHBand="0"/, '镶边行开启时 noHBand 应为 0')
})

test('★ 回归：重复套用样式不能留下多个 tblStyle（删除后 children 是旧快照）', () => {
  const doc = DocxDocument.open(original)
  doc.setTableStyle({ table: 0, styleId: 'TableGrid' })
  doc.setTableStyle({ table: 0, styleId: 'TableGrid' })
  doc.setTableStyle({ table: 0, styleId: 'TableGrid' })
  const xml = ZipPackage.open(doc.save()).readText('word/document.xml')
  const tblPr = /<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(xml)[0]
  assert.equal((tblPr.match(/<w:tblStyle /g) ?? []).length, 1, `tblStyle 应只有一个：${tblPr}`)
  assert.equal((tblPr.match(/<w:tblLook /g) ?? []).length, 1, `tblLook 应只有一个：${tblPr}`)
})

test('★ 文档里没有该样式时补一个最小定义（否则 Word 静默退回普通表格）', () => {
  const doc = DocxDocument.open(original)
  const result = doc.setTableStyle({ table: 0, styleId: 'DshCustomTable' })
  assert.equal(result.style_defined_in_document, true)
  const out = doc.save()
  assert.match(ZipPackage.open(out).readText('word/styles.xml'), /w:styleId="DshCustomTable"/)
  // 第二次套用同一 ID 时不应重复新建
  const again = DocxDocument.open(out)
  assert.equal(again.setTableStyle({ table: 0, styleId: 'DshCustomTable' }).style_defined_in_document, false)
})

test('★ 样式写入只改 document.xml 与 styles.xml', () => {
  const pkgBefore = ZipPackage.open(original)
  const before = new Map(pkgBefore.names().map((n) => [n, pkgBefore.read(n)]))
  const doc = DocxDocument.open(original)
  doc.setTableStyle({ table: 0, styleId: 'DshCustomTable' })
  const pkgAfter = ZipPackage.open(doc.save())
  const changed = [...before.keys()].filter((n) => !pkgAfter.read(n).equals(before.get(n))).sort()
  assert.deepEqual(changed, ['word/document.xml', 'word/styles.xml'], `意外改动：${changed.join(', ')}`)
})

test('样式 ID 不合法与表格越界时拒绝', () => {
  const doc = DocxDocument.open(original)
  assert.throws(() => doc.setTableStyle({ table: 0, styleId: '9bad id' }), (err) => err.code === 'INVALID_REQUEST')
  assert.throws(() => doc.setTableStyle({ table: 9, styleId: 'TableGrid' }), (err) => err.code === 'FILE_NOT_FOUND')
})

console.log('\n=== 16. 新建表格 ===')

test('★ 新建表格：网格、表头加粗、单元格宽度都写对', () => {
  const doc = DocxDocument.open(original)
  const result = doc.createTable({ after: 0, rows: [['产品', '数量'], ['笔记本', '3']], header: true })
  assert.equal(result.rows, 2)
  assert.equal(result.columns, 2)
  assert.equal(result.column_widths.length, 2)
  const out = doc.save()
  const xml = ZipPackage.open(out).readText('word/document.xml')
  assert.match(xml, /<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"\/>/)
  assert.match(xml, /<w:tblGrid><w:gridCol w:w="\d+"\/><w:gridCol w:w="\d+"\/><\/w:tblGrid>/, 'tblGrid 必须给出每列宽度')
  assert.match(xml, /<w:trPr><w:tblHeader\/><\/w:trPr>/, '表头行标记 tblHeader')
  assert.match(xml, /<w:rPr><w:b\/><\/w:rPr>/, '表头文字加粗')
  const tables = DocxDocument.open(out).tables()
  assert.equal(tables.count, 2, '原有一个表格 + 新建一个')
  // 新表插在正文最前，所以按内容定位而不赌下标（原表在后面）
  const created = tables.tables.find((t) => t.cells[0][0].text === '产品')
  assert.ok(created, `找不到新建的表格：${JSON.stringify(tables.tables.map((t) => t.cells[0][0].text))}`)
  assert.deepEqual(created.cells.map((row) => row.map((cell) => cell.text)), [['产品', '数量'], ['笔记本', '3']])
})

test('★ 新建表格：可指定列宽与自定义样式（样式定义自动补齐）', () => {
  const doc = DocxDocument.open(original)
  const result = doc.createTable({
    after: -1,
    rows: [['A', 'B']],
    header: false,
    styleId: 'DshNewTable',
    columnWidths: [2000, 3000]
  })
  assert.deepEqual(result.column_widths, [2000, 3000])
  assert.equal(result.style_defined_in_document, true)
  const out = doc.save()
  assert.match(ZipPackage.open(out).readText('word/styles.xml'), /w:styleId="DshNewTable"/)
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(ZipPackage.open(out).readText('word/document.xml'))[1]
  assert.ok(body.trimStart().startsWith('<w:tbl>'), 'after=-1 应插到正文最前')
})

test('★ 表格紧跟在最后一个段落之后时自动补空段落（Word 的硬性要求）', () => {
  const doc = DocxDocument.open(original)
  const last = doc.bodyParagraphs().length - 1
  doc.createTable({ after: last, rows: [['尾表']] })
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(ZipPackage.open(doc.save()).readText('word/document.xml'))[1]
  // 表格之后必须紧跟一个段落；再往后才是 sectPr（如果文档有的话）
  assert.match(body, /<\/w:tbl><w:p\/>(<w:sectPr|$)/, `表格后应有空段落：${body.slice(-120)}`)
})

test('新建表格的参数校验', () => {
  const doc = DocxDocument.open(original)
  assert.throws(() => doc.createTable({ after: 0, rows: [] }), (err) => err.code === 'INVALID_REQUEST')
  assert.throws(() => doc.createTable({ after: 0, rows: [[]] }), (err) => err.code === 'INVALID_REQUEST')
  assert.throws(() => doc.createTable({ after: 99, rows: [['x']] }), (err) => err.code === 'INVALID_REQUEST')
})

console.log('\n=== 17. 图片插入 ===')

/**
 * 生成一张合法的 PNG（RGB，无滤波），用于测试真实图片路径。
 * @param {number} width - 宽。
 * @param {number} height - 高。
 * @param {number[]} rgb - 颜色。
 * @returns {Buffer} PNG 字节。
 */
function makePng(width, height, rgb = [255, 0, 0]) {
  const raw = Buffer.alloc((width * 3 + 1) * height)
  let o = 0
  for (let y = 0; y < height; y += 1) {
    raw[o] = 0
    o += 1
    for (let x = 0; x < width; x += 1) {
      raw[o] = rgb[0]
      raw[o + 1] = rgb[1]
      raw[o + 2] = rgb[2]
      o += 3
    }
  }
  const chunk = (type, payload) => {
    const len = Buffer.alloc(4)
    len.writeUInt32BE(payload.length)
    const t = Buffer.from(type, 'latin1')
    const crc = Buffer.alloc(4)
    crc.writeUInt32BE(zlib.crc32(Buffer.concat([t, payload])) >>> 0)
    return Buffer.concat([len, t, payload, crc])
  }
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(width, 0)
  ihdr.writeUInt32BE(height, 4)
  ihdr[8] = 8
  ihdr[9] = 2
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlib.deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0))
  ])
}

const png = makePng(120, 60)

test('按真实字节识别图片类型', () => {
  assert.equal(detectImageType(png), 'png')
  assert.throws(() => detectImageType(Buffer.from('not an image')), (e) => e.code === 'UNSUPPORTED_FILE_TYPE')
})

test('读取 PNG 像素尺寸', () => {
  assert.deepEqual(readImageSize(png, 'png'), { width: 120, height: 60 })
})

test('插入图片补齐媒体部件、内容类型与关系', () => {
  const doc = DocxDocument.open(original)
  const result = doc.insertImage({ after: 0, data: png, extension: 'png', altText: '测试图' })
  assert.equal(result.media_part, 'word/media/image1.png')
  const pkg = ZipPackage.open(doc.save())
  assert.ok(pkg.has('word/media/image1.png'), '缺少媒体部件')
  assert.ok(pkg.read('word/media/image1.png').equals(png), '媒体部件内容不一致')
  assert.ok(pkg.readText('[Content_Types].xml').includes('Extension="png"'), '缺少 png 内容类型声明')
  assert.ok(pkg.readText('word/_rels/document.xml.rels').includes('relationships/image'), '缺少 image 关系')
})

test('★ 插入后 validate_docx 通过（证明关系链完整）', () => {
  const doc = DocxDocument.open(original)
  doc.insertImage({ after: 0, data: png, extension: 'png' })
  const report = DocxDocument.open(doc.save()).validate()
  assert.equal(report.valid, true, JSON.stringify(report.checks.filter((c) => !c.ok)))
})

test('★ 图片引用能被校验器解析（不是悬空关系）', () => {
  const doc = DocxDocument.open(original)
  doc.insertImage({ after: 0, data: png, extension: 'png' })
  const report = DocxDocument.open(doc.save()).validate()
  const check = report.checks.find((c) => c.name === '图片/超链接/页眉页脚引用可解析')
  assert.equal(check.ok, true, check.detail)
  // 引用总数 = 1 张图片 + 页眉引用 + 页脚引用
  assert.ok(check.detail.includes('3 个引用全部可解析'), check.detail)
})

test('未指定尺寸时用图片原始像素', () => {
  const doc = DocxDocument.open(original)
  const result = doc.insertImage({ after: 0, data: png, extension: 'png' })
  assert.deepEqual(result.display_size_px, { width: 120, height: 60 })
})

test('只给宽度时按比例缩放', () => {
  const doc = DocxDocument.open(original)
  const result = doc.insertImage({ after: 0, data: png, extension: 'png', widthPx: 240 })
  assert.deepEqual(result.display_size_px, { width: 240, height: 120 })
})

test('drawing id 不与既有 drawing 冲突', () => {
  const doc = DocxDocument.open(original)
  const a = doc.insertImage({ after: 0, data: png, extension: 'png' })
  // 第二张图必须拿到不同的 docPr id，Word 对重复 id 会报错
  const doc2 = DocxDocument.open(doc.save())
  doc2.insertImage({ after: 0, data: png, extension: 'png' })
  const xml = ZipPackage.open(doc2.save()).readText('word/document.xml')
  const ids = [...xml.matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => m[1])
  assert.equal(ids.length, 2)
  assert.notEqual(ids[0], ids[1], `docPr id 重复：${ids.join(', ')}`)
})

test('多张图片各自获得独立媒体部件与关系', () => {
  const doc = DocxDocument.open(original)
  const r1 = doc.insertImage({ after: 0, data: png, extension: 'png' })
  const r2 = doc.insertImage({ after: 1, data: makePng(10, 10, [0, 0, 255]), extension: 'png' })
  assert.notEqual(r1.media_part, r2.media_part)
  assert.notEqual(r1.relationship_id, r2.relationship_id)
  const pkg = ZipPackage.open(doc.save())
  assert.ok(pkg.has('word/media/image1.png'))
  assert.ok(pkg.has('word/media/image2.png'))
})

test('插入位置越界被拒绝', () => {
  const doc = DocxDocument.open(original)
  assert.throws(() => doc.insertImage({ after: 99, data: png, extension: 'png' }), (e) => e.code === 'INVALID_REQUEST')
})

test('★ 插入图片只新增媒体部件，正文与样式等部件不受损', () => {
  const before = new Map()
  const pkgBefore = ZipPackage.open(original)
  for (const n of pkgBefore.names()) before.set(n, pkgBefore.read(n))

  const doc = DocxDocument.open(original)
  doc.insertImage({ after: 0, data: png, extension: 'png' })
  const pkgAfter = ZipPackage.open(doc.save())

  const added = pkgAfter.names().filter((n) => !before.has(n))
  assert.deepEqual(added, ['word/media/image1.png'], `新增部件异常：${added.join(', ')}`)
  const removed = [...before.keys()].filter((n) => !pkgAfter.has(n))
  assert.deepEqual(removed, [], `不应删除任何部件：${removed.join(', ')}`)
  // 样式、页眉、文档属性必须逐字节不变
  for (const n of ['word/styles.xml', 'word/header1.xml', 'word/footer1.xml', 'docProps/core.xml']) {
    assert.ok(pkgAfter.read(n).equals(before.get(n)), `${n} 被改动`)
  }
})

test('图片段落出现在指定位置', () => {
  const doc = DocxDocument.open(original)
  doc.insertImage({ after: 0, data: png, extension: 'png' })
  const paragraphs = DocxDocument.open(doc.save()).paragraphs()
  assert.equal(paragraphs.total, 4)
  assert.equal(paragraphs.paragraphs[1].text, '', '图片段落本身不含文本')
  assert.equal(paragraphs.paragraphs[2].text, '本文件由 dsh-exp-office 生成。加粗片段')
})

console.log('\n=== 17. 图片清单 / 尺寸 / 删除 ===')

/** 造一份带一张图片的文档（复用图片插入的实现）。 */
function withImageDoc() {
  const doc = DocxDocument.open(original)
  doc.insertImage({ after: 0, data: makePng(240, 120), extension: 'png', widthPx: 240, heightPx: 120, altText: '占位图' })
  return DocxDocument.open(doc.save())
}

test('★ 图片清单：下标、尺寸（像素）、关系 id 都对', () => {
  const doc = withImageDoc()
  const images = doc.images()
  assert.equal(images.length, 1)
  assert.equal(images[0].index, 0)
  assert.equal(images[0].kind, 'inline')
  assert.equal(images[0].width_px, 240)
  assert.equal(images[0].height_px, 120)
  assert.ok(images[0].relationship_id, '应能读到图片关系 id')
})

test('★ 调整尺寸：只给宽度按比例，给两向按给定值，两份 EMU 同步', () => {
  const doc = withImageDoc()
  const scaled = doc.resizeImage({ index: 0, widthPx: 480 })
  assert.deepEqual(scaled.to, { width: 480, height: 240 }, '只给宽度应保持 2:1 比例')
  const both = doc.resizeImage({ index: 0, widthPx: 200, heightPx: 90 })
  assert.deepEqual(both.to, { width: 200, height: 90 })
  const out = doc.save()
  const reopened = DocxDocument.open(out)
  assert.deepEqual(
    { width: reopened.images()[0].width_px, height: reopened.images()[0].height_px },
    { width: 200, height: 90 }
  )
  // wp:extent 与 a:ext 必须一致（只改一处 Word 会把图片弹回原大小）
  const xml = ZipPackage.open(out).readText('word/document.xml')
  const extent = /<wp:extent cx="(\d+)" cy="(\d+)"\/>/.exec(xml)
  const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(xml)
  assert.deepEqual([extent[1], extent[2]], [ext[1], ext[2]], 'wp:extent 与 a:ext 必须同步')
  assert.equal(Number(extent[1]), 200 * 9525)
})

test('★ 删除图片：只含图片的段落会被一起删掉，媒体与关系保留', () => {
  const doc = withImageDoc()
  const paragraphsBefore = doc.bodyParagraphs().length
  const result = doc.deleteImage({ index: 0 })
  assert.equal(result.removed_paragraph, true)
  const out = doc.save()
  const reopened = DocxDocument.open(out)
  assert.equal(reopened.images().length, 0)
  assert.equal(reopened.bodyParagraphs().length, paragraphsBefore - 1, '空段落应一并删除')
  assert.match(result.note, /媒体部件与关系保留/)
  // 媒体确实还在包里
  assert.ok(ZipPackage.open(out).names().some((n) => n.startsWith('word/media/')), '媒体部件应保留')
})

test('图片下标越界与参数校验', () => {
  const doc = withImageDoc()
  assert.throws(() => doc.resizeImage({ index: 0 }), (err) => err.code === 'INVALID_REQUEST')
  assert.throws(() => doc.resizeImage({ index: 0, widthPx: -5 }), (err) => err.code === 'INVALID_REQUEST')
  assert.throws(() => doc.resizeImage({ index: 3, widthPx: 10 }), (err) => err.code === 'FILE_NOT_FOUND')
  assert.throws(() => doc.deleteImage({ index: 3 }), (err) => err.code === 'FILE_NOT_FOUND')
})

console.log('\n=== 18. 段落样式 / 字符样式 / 分页符 ===')

test('★ 设置段落样式：只加 pStyle，原样式与其它段落属性保留', () => {
  const doc = DocxDocument.open(original)
  const result = doc.setParagraphStyle({ index: 0, styleId: 'Heading2' })
  assert.equal(result.from, 'Heading1')
  assert.equal(result.style_id, 'Heading2')
  const out = doc.save()
  const reopened = DocxDocument.open(out)
  assert.equal(reopened.paragraphs().paragraphs[0].style, 'Heading2')
  assert.equal(reopened.paragraphs().paragraphs[1].style, null, '其它段落不应被改动')
  // pStyle 只应出现一次（同一段落不能有两个）
  const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(ZipPackage.open(out).readText('word/document.xml'))
  assert.equal((pPr[0].match(/<w:pStyle /g) ?? []).length, 1)
})

test('★ 回归：重复设置段落样式不能留下多个 pStyle（删完再取锚点会踩旧快照）', () => {
  const doc = DocxDocument.open(original)
  doc.setParagraphStyle({ index: 0, styleId: 'Heading2' })
  doc.setParagraphStyle({ index: 0, styleId: 'Heading3' })
  doc.setParagraphStyle({ index: 0, styleId: 'Heading4' })
  const xml = ZipPackage.open(doc.save()).readText('word/document.xml')
  const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)[0]
  assert.equal((pPr.match(/<w:pStyle /g) ?? []).length, 1, `pStyle 应只有一个：${pPr}`)
  assert.match(pPr, /w:val="Heading4"/, '应以最后一次为准')
})

test('★ 样式定义缺失时自动补齐（否则 Word 静默退回默认样式）', () => {
  const doc = DocxDocument.open(original)
  const para = doc.setParagraphStyle({ index: 1, styleId: 'DshBodyText' })
  assert.equal(para.style_defined_in_document, true)
  const run = doc.setCharacterStyle({ index: 1, styleId: 'DshEmphasis' })
  assert.equal(run.style_defined_in_document, true)
  const styles = ZipPackage.open(doc.save()).readText('word/styles.xml')
  assert.match(styles, /w:type="paragraph" w:styleId="DshBodyText"/)
  assert.match(styles, /w:type="character" w:styleId="DshEmphasis"/)
})

test('★ 字符样式写到该段所有 run 上', () => {
  const doc = DocxDocument.open(original)
  const result = doc.setCharacterStyle({ index: 1, styleId: 'Strong' })
  assert.ok(result.runs >= 1)
  const xml = ZipPackage.open(doc.save()).readText('word/document.xml')
  assert.equal((xml.match(/<w:rStyle /g) ?? []).length, result.runs)
})

test('★ 分页符：插入独立段落且带 w:br type=page', () => {
  const doc = DocxDocument.open(original)
  const before = doc.bodyParagraphs().length
  const result = doc.insertPageBreak({ after: 0 })
  assert.equal(result.index, 1)
  const reopened = DocxDocument.open(doc.save())
  assert.equal(reopened.bodyParagraphs().length, before + 1)
  assert.match(ZipPackage.open(doc.save()).readText('word/document.xml'), /<w:br w:type="page"\/>/)
})

test('段落/字符样式与分页符的参数校验', () => {
  const doc = DocxDocument.open(original)
  assert.throws(() => doc.setParagraphStyle({ index: 9, styleId: 'Heading2' }), (err) => err.code === 'FILE_NOT_FOUND')
  assert.throws(() => doc.setParagraphStyle({ index: 0, styleId: '9bad' }), (err) => err.code === 'INVALID_REQUEST')
  assert.throws(() => doc.setCharacterStyle({ index: 0, styleId: 'bad id' }), (err) => err.code === 'INVALID_REQUEST')
  assert.throws(() => doc.insertPageBreak({ after: 99 }), (err) => err.code === 'INVALID_REQUEST')
})

console.log('\n=== 19. 新建页眉页脚部件 ===')

test('★ 文档没有页眉部件时自动新建并接线', () => {
  const pkg = ZipPackage.open(original)
  pkg.delete('word/header1.xml')
  const doc = DocxDocument.open(pkg.toBuffer())
  assert.equal(doc.info.headers.length, 0)

  const result = doc.setHeaderFooter({ kind: 'header', text: '新建的页眉' })
  assert.equal(result.created, true)
  assert.equal(result.part, 'word/header1.xml')

  const after = ZipPackage.open(doc.save())
  assert.ok(after.has('word/header1.xml'), '未创建页眉部件')
  assert.ok(after.readText('[Content_Types].xml').includes('header+xml'), '缺少页眉内容类型 Override')
  assert.ok(after.readText('word/_rels/document.xml.rels').includes('relationships/header'), '缺少 header 关系')
  assert.ok(after.readText('word/document.xml').includes('<w:headerReference'), 'sectPr 里缺少页眉引用')
})

test('★ 新建页眉后 validate_docx 通过', () => {
  const pkg = ZipPackage.open(original)
  pkg.delete('word/header1.xml')
  const doc = DocxDocument.open(pkg.toBuffer())
  doc.setHeaderFooter({ kind: 'header', text: '新页眉' })
  const report = DocxDocument.open(doc.save()).validate()
  assert.equal(report.valid, true, JSON.stringify(report.checks.filter((c) => !c.ok)))
})

test('★ 新建页脚并加盖页码域', () => {
  const pkg = ZipPackage.open(original)
  pkg.delete('word/footer1.xml')
  const doc = DocxDocument.open(pkg.toBuffer())
  const result = doc.setHeaderFooter({ kind: 'footer', text: '第 ', pageNumber: true })
  assert.equal(result.created, true)
  const reopened = DocxDocument.open(doc.save())
  assert.equal(reopened.info.footers.length, 1)
  assert.equal(reopened.headersOrFooters('footer')[0].has_page_number_field, true)
  assert.equal(reopened.validate().valid, true)
})

test('★ 原文没有 sectPr 时自动补一个', () => {
  const pkg = ZipPackage.open(original)
  pkg.delete('word/header1.xml')
  const xml = pkg.readText('word/document.xml').replace(/<w:sectPr>[\s\S]*?<\/w:sectPr>/, '')
  pkg.write('word/document.xml', xml)
  const doc = DocxDocument.open(pkg.toBuffer())
  doc.setHeaderFooter({ kind: 'header', text: '无节文档的页眉' })
  const out = ZipPackage.open(doc.save()).readText('word/document.xml')
  assert.ok(out.includes('<w:sectPr>'), '未补出 sectPr')
  assert.ok(out.includes('<w:headerReference'), 'sectPr 内缺少页眉引用')
  assert.equal(DocxDocument.open(ZipPackage.open(doc.save()).toBuffer()).validate().valid, true)
})

test('已存在的页眉不会被重复新建', () => {
  const doc = DocxDocument.open(original)
  const result = doc.setHeaderFooter({ kind: 'header', text: '只改文本' })
  assert.equal(result.created, false)
  assert.equal(ZipPackage.open(doc.save()).names().filter((n) => /^word\/header\d*\.xml$/.test(n)).length, 1)
})

console.log('\n=== 20. 页面设置与页码域 ===')

test('设置纸张与方向：横向交换宽高并写 orient', () => {
  const doc = DocxDocument.open(original)
  const result = doc.setPageLayout({ orientation: 'landscape' })
  assert.equal(result.from.width_twips, 11906)
  assert.equal(result.to.width_twips, 16838)
  assert.equal(result.to.height_twips, 11906)
  assert.equal(result.to.orientation, 'landscape')

  const xml = ZipPackage.open(doc.save()).readText('word/document.xml')
  assert.ok(xml.includes('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'), 'pgSz 应写成横向')
  assert.equal(DocxDocument.open(doc.save()).validate().valid, true)
})

test('★ 切回纵向时不写 orient（默认可省，写错反被 Word 挑剔）', () => {
  const doc = DocxDocument.open(original)
  doc.setPageLayout({ orientation: 'landscape' })
  doc.setPageLayout({ orientation: 'portrait' })
  const xml = ZipPackage.open(doc.save()).readText('word/document.xml')
  assert.ok(xml.includes('<w:pgSz w:w="11906" w:h="16838"/>'), `纵向 pgSz 不正确：${xml.match(/<w:pgSz[^>]*>/)}`)
  assert.ok(!/w:orient="portrait"/.test(xml), '纵向不应写 orient 属性')
})

test('纸张预设与自定义尺寸', () => {
  const a3 = DocxDocument.open(original)
  assert.deepEqual(a3.setPageLayout({ paper: 'A3' }).to.width_twips, 16838)
  const custom = DocxDocument.open(original)
  const r = custom.setPageLayout({ widthCm: 20, heightCm: 30 })
  assert.equal(r.to.width_twips, 11339)
  assert.equal(r.to.height_twips, 17008)
  assert.equal(r.to.width_cm, 20)
  assert.equal(r.to.height_cm, 30)
})

test('★ 纸张/方向参数校验', () => {
  const doc = DocxDocument.open(original)
  assert.throws(() => doc.setPageLayout({ paper: 'B5' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.setPageLayout({ orientation: 'sideways' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.setPageLayout({ widthCm: -5 }), (e) => e.code === 'INVALID_REQUEST')
})

test('★ 页边距逐项改写：只覆盖传入的项', () => {
  const doc = DocxDocument.open(original)
  const before = Number(/w:left="(\d+)"/.exec(ZipPackage.open(original).readText('word/document.xml'))[1])
  const result = doc.setMargins({ topCm: 3 })
  assert.equal(result.to.top_cm, 3)
  assert.equal(result.to.bottom_cm, 2.54)
  assert.equal(result.updated, 'in_place')
  const xml = ZipPackage.open(doc.save()).readText('word/document.xml')
  assert.ok(xml.includes('w:top="1701"'), `上边距应为 3cm=1701twips：${xml.match(/<w:pgMar[^>]*>/)}`)
  assert.ok(xml.includes(`w:left="${before}"`), '未传入的左页边距必须保持原值')
})

test('★ 没有 pgMar 时新建并写齐七项（缺项会让阅读器各取默认值）', () => {
  const pkg = ZipPackage.open(original)
  pkg.write('word/document.xml', pkg.readText('word/document.xml').replace(/<w:pgMar[^>]*\/>/, ''))
  const doc = DocxDocument.open(pkg.toBuffer())
  const result = doc.setMargins({ leftCm: 2, rightCm: 2 })
  const xml = ZipPackage.open(doc.save()).readText('word/document.xml')
  const pgMar = /<w:pgMar[^>]*>/.exec(xml)[0]
  for (const key of ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter']) {
    assert.ok(pgMar.includes(`w:${key}="`), `新建的 pgMar 缺少 w:${key}`)
  }
  assert.equal(result.updated, 'created')
  assert.equal(result.to.top_cm, 2.54)
})

test('★ pgMar 必须排在 pgSz 之后（顺序错了 Word 报文档损坏）', () => {
  const pkg = ZipPackage.open(original)
  pkg.write('word/document.xml', pkg.readText('word/document.xml').replace(/<w:pgMar[^>]*\/>/, ''))
  const doc = DocxDocument.open(pkg.toBuffer())
  doc.setMargins({ topCm: 2 })
  const sect = /<w:sectPr[\s\S]*?<\/w:sectPr>/.exec(ZipPackage.open(doc.save()).readText('word/document.xml'))[0]
  assert.ok(sect.indexOf('<w:pgSz') < sect.indexOf('<w:pgMar'), `子元素顺序不对：${sect}`)
})

test('页边距参数校验：至少要给一项、不能为负、不能超过页宽', () => {
  const doc = DocxDocument.open(original)
  assert.throws(() => doc.setMargins({}), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.setMargins({ leftCm: -1 }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.setMargins({ leftCm: 12, rightCm: 12 }), (e) => e.code === 'INVALID_REQUEST')
})

test('★ 页码域：pPr 在前、正文只有一份（零长插入与替换补丁同起点的回归）', () => {
  const doc = DocxDocument.open(original)
  const result = doc.insertPageNumber({ position: 'footer', align: 'center', prefix: '第 ', suffix: ' 页' })
  const xml = ZipPackage.open(doc.save()).readText(result.part)
  // pPr 必须是 w:p 的第一个子元素，且旧内容不能被留下（否则页码/文字会重复）
  const p = /<w:p>[\s\S]*?<\/w:p>/.exec(xml)[0]
  assert.ok(p.indexOf('<w:pPr>') < p.indexOf('<w:r>'), `pPr 必须在正文之前：${p}`)
  assert.equal((p.match(/<w:jc /g) ?? []).length, 1, 'jc 只应有一个')
  assert.ok(p.includes('<w:instrText xml:space="preserve"> PAGE </w:instrText>'), '应写入 PAGE 域')
  assert.equal((p.match(/<w:fldSimple/g) ?? []).length, 0, '旧内容必须被替换掉')
  assert.ok(p.includes('第 ') && p.includes(' 页'), '前后缀应保留')
  assert.equal(DocxDocument.open(doc.save()).headersOrFooters('footer')[0].has_page_number_field, true)
})

test('★ 页码域可放页眉、可右对齐', () => {
  const doc = DocxDocument.open(original)
  const result = doc.insertPageNumber({ position: 'header', align: 'right' })
  assert.equal(result.position, 'header')
  const xml = ZipPackage.open(doc.save()).readText(result.part)
  assert.ok(xml.includes('<w:jc w:val="right"/>'), '应写入右对齐')
  assert.equal(DocxDocument.open(doc.save()).validate().valid, true)
})

test('★ 已有 pPr 的段落：只在 pPr 内加 jc，正文仍被整段替换', () => {
  const pkg = ZipPackage.open(original)
  pkg.write('word/footer1.xml', pkg.readText('word/footer1.xml').replace('<w:p>', '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'))
  const doc = DocxDocument.open(pkg.toBuffer())
  doc.insertPageNumber({ position: 'footer', align: 'left', prefix: 'p.', suffix: '' })
  const xml = ZipPackage.open(doc.save()).readText('word/footer1.xml')
  const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)[0]
  assert.ok(pPr.includes('<w:spacing w:after="0"/>'), '原有段落属性必须保留')
  assert.ok(/<w:spacing[^>]*\/><w:jc /.test(pPr), `jc 应排在 spacing 之后：${pPr}`)
  assert.equal((xml.match(/<w:fldSimple/g) ?? []).length, 0, '旧页码内容应被替换')
})

test('页码域参数校验', () => {
  const doc = DocxDocument.open(original)
  assert.throws(() => doc.insertPageNumber({ position: 'body' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.insertPageNumber({ align: 'justify' }), (e) => e.code === 'INVALID_REQUEST')
})

test('★ 页面设置只改 document.xml / 页脚部件，其余部件逐字节不变', () => {
  const before = new Map()
  const pkgBefore = ZipPackage.open(original)
  for (const n of pkgBefore.names()) before.set(n, pkgBefore.read(n))

  const doc = DocxDocument.open(original)
  doc.setPageLayout({ paper: 'A3', orientation: 'landscape' })
  doc.setMargins({ leftCm: 2, rightCm: 2 })
  const pkgAfter = ZipPackage.open(doc.save())
  const changed = [...before.keys()].filter((n) => !pkgAfter.read(n).equals(before.get(n)))
  assert.deepEqual(changed, ['word/document.xml'], `意外改动：${changed.join(', ')}`)
})

console.log('\n=== 21. 图片环绕方式 ===')

/** @returns {import('../lib/docx.js').DocxDocument} 带一张行内图的文档。 */
function docWithImage() {
  return DocxDocument.open(readFileSync(join(here, 'fixtures', 'report-image.docx')))
}

test('★ 行内图 → 四周型：换容器、写齐 anchor 属性、元素顺序符合 schema', () => {
  const doc = docWithImage()
  const result = doc.setImageWrap({ index: 0, wrap: 'square', offsetXEmu: 114300, offsetYEmu: 57150 })
  assert.equal(result.from, 'inline')
  assert.equal(result.to, 'square')
  assert.equal(result.changed, true)

  const xml = ZipPackage.open(doc.save()).readText('word/document.xml')
  const anchor = /<wp:anchor[\s\S]*?<\/wp:anchor>/.exec(xml)[0]
  for (const attribute of ['simplePos="0"', 'relativeHeight="', 'behindDoc="0"', 'locked="0"', 'layoutInCell="1"', 'allowOverlap="1"']) {
    assert.ok(anchor.includes(attribute), `anchor 缺少必需属性 ${attribute}`)
  }
  const order = (anchor.match(/<wp:(simplePos|positionH|positionV|extent|effectExtent|wrap\w+|docPr)/g) ?? []).join(' ')
  assert.equal(order, '<wp:simplePos <wp:positionH <wp:positionV <wp:extent <wp:effectExtent <wp:wrapSquare <wp:docPr')
  assert.ok(anchor.includes(`<wp:positionH relativeFrom="column"><wp:posOffset>114300</wp:posOffset></wp:positionH>`))
  assert.ok(anchor.includes(`<wp:positionV relativeFrom="paragraph"><wp:posOffset>57150</wp:posOffset></wp:positionV>`))
  assert.equal(DocxDocument.open(doc.save()).validate().valid, true)
})

test('★ 读回环绕方式（images() 的 wrap 字段）', () => {
  for (const [wrap, expected] of [
    ['square', 'square'],
    ['topAndBottom', 'topAndBottom'],
    ['behind', 'behind'],
    ['inFront', 'inFront'],
    // Word 里没有独立的「无环绕」：none 与 inFront 是同一份 XML，统一归一成 inFront
    ['none', 'inFront']
  ]) {
    const doc = docWithImage()
    const result = doc.setImageWrap({ index: 0, wrap })
    assert.equal(result.to, expected, `wrap=${wrap} 的返回名应归一`)
    assert.equal(DocxDocument.open(doc.save()).images()[0].wrap, expected, `wrap=${wrap} 读回不对`)
  }
})

test('★ 衬于文字下方与浮于文字上方只差 behindDoc，z 序也要跟着降', () => {
  const below = docWithImage()
  below.setImageWrap({ index: 0, wrap: 'behind' })
  const belowXml = ZipPackage.open(below.save()).readText('word/document.xml')
  assert.ok(/behindDoc="1"/.test(belowXml))
  assert.ok(/relativeHeight="0"/.test(belowXml), '衬于文字下方应把 z 序降到 0')

  const above = docWithImage()
  above.setImageWrap({ index: 0, wrap: 'inFront' })
  const aboveXml = ZipPackage.open(above.save()).readText('word/document.xml')
  assert.ok(/behindDoc="0"/.test(aboveXml))
  assert.ok(/relativeHeight="251658240"/.test(aboveXml))
})

test('★ 浮动 → 行内：换回 wp:inline 并删掉浮动专有的子元素与属性', () => {
  const doc = docWithImage()
  doc.setImageWrap({ index: 0, wrap: 'square' })
  const back = doc.setImageWrap({ index: 0, wrap: 'inline' })
  assert.equal(back.from, 'anchor')
  assert.equal(back.to, 'inline')

  const bytes = doc.save()
  const xml = ZipPackage.open(bytes).readText('word/document.xml')
  assert.ok(!xml.includes('<wp:anchor'), '不应再有 anchor')
  assert.ok(!xml.includes('<wp:simplePos'), 'simplePos 应被删除')
  assert.ok(!xml.includes('<wp:wrapSquare'), '环绕元素应被删除')
  assert.ok(!/simplePos="0"/.test(xml), 'anchor 专有属性应被删除')
  const inline = /<wp:inline[\s\S]*?<\/wp:inline>/.exec(xml)[0]
  assert.equal(
    (inline.match(/<wp:(simplePos|positionH|positionV|extent|effectExtent|wrap\w+|docPr)/g) ?? []).join(' '),
    '<wp:extent <wp:effectExtent <wp:docPr'
  )
  assert.equal(DocxDocument.open(bytes).validate().valid, true)
  // 尺寸不受环绕方式影响
  assert.equal(DocxDocument.open(bytes).images()[0].width_px, 240)
})

test('本来就 行内 时为空操作（changed=false）', () => {
  const result = docWithImage().setImageWrap({ index: 0, wrap: 'inline' })
  assert.equal(result.changed, false)
})

test('来回切换保持稳定：square → inline → square', () => {
  const doc = docWithImage()
  doc.setImageWrap({ index: 0, wrap: 'square' })
  doc.setImageWrap({ index: 0, wrap: 'inline' })
  doc.setImageWrap({ index: 0, wrap: 'square' })
  const bytes = doc.save()
  const xml = ZipPackage.open(bytes).readText('word/document.xml')
  assert.equal((xml.match(/<wp:anchor/g) ?? []).length, 1)
  assert.equal((xml.match(/<wp:wrapSquare/g) ?? []).length, 1)
  assert.equal((xml.match(/<wp:simplePos/g) ?? []).length, 1)
  assert.equal(DocxDocument.open(bytes).validate().valid, true)
})

test('★ 环绕方式与参数的校验：tight/through 明确拒绝', () => {
  const doc = docWithImage()
  assert.throws(() => doc.setImageWrap({ index: 0, wrap: 'tight' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.setImageWrap({ index: 0, wrap: 'through' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.setImageWrap({ index: 0, wrap: 'diagonal' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.setImageWrap({ index: 0, wrap: 'square', relativeFromH: 'moon' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.setImageWrap({ index: 0, wrap: 'square', relativeFromV: 'moon' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.setImageWrap({ index: 0, wrap: 'square', offsetXEmu: Number.NaN }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.setImageWrap({ index: 9, wrap: 'square' }), (e) => e.code === 'FILE_NOT_FOUND')
})

test('★ 只改目标图片所在的段落，其余部件逐字节不变', () => {
  const source = join(here, 'fixtures', 'report-image.docx')
  const before = new Map()
  const pkgBefore = ZipPackage.open(readFileSync(source))
  for (const n of pkgBefore.names()) before.set(n, pkgBefore.read(n))

  const doc = DocxDocument.open(readFileSync(source))
  doc.setImageWrap({ index: 0, wrap: 'square' })
  const pkgAfter = ZipPackage.open(doc.save())
  const changed = [...before.keys()].filter((n) => !pkgAfter.read(n).equals(before.get(n)))
  assert.deepEqual(changed, ['word/document.xml'], `意外改动：${changed.join(', ')}`)
})

console.log('\n=== 22. 书签 ===')

test('★ 插入书签：一对标记包住段落内容，id 唯一', () => {
  const doc = DocxDocument.open(original)
  const result = doc.insertBookmark({ paragraph: 1, name: 'intro' })
  assert.equal(result.name, 'intro')
  assert.equal(result.id, 0)
  assert.equal(result.bookmark_count, 1)

  const bytes = doc.save()
  const xml = ZipPackage.open(bytes).readText('word/document.xml')
  const p = /<w:p>(?:(?!<\/w:p>)[\s\S])*?bookmarkStart[\s\S]*?<\/w:p>/.exec(xml)[0]
  // start 在正文之前、end 在正文之后
  assert.ok(p.indexOf('<w:bookmarkStart w:id="0" w:name="intro"/>') < p.indexOf('<w:r>'), 'bookmarkStart 应在 run 之前')
  assert.ok(p.trimEnd().endsWith('<w:bookmarkEnd w:id="0"/></w:p>'), `bookmarkEnd 应在段落末尾：${p.slice(-60)}`)
  assert.equal(DocxDocument.open(bytes).validate().valid, true)
})

test('★ 书签读回：名称、id、成对性、所属段落与文本', () => {
  const doc = DocxDocument.open(original)
  doc.insertBookmark({ paragraph: 1, name: 'intro' })
  doc.insertBookmark({ paragraph: 2, name: '章节_1' })
  const list = DocxDocument.open(doc.save()).bookmarks()
  assert.equal(list.length, 2)
  assert.deepEqual(list.map((b) => b.name), ['intro', '章节_1'])
  assert.deepEqual(list.map((b) => b.id), [0, 1])
  assert.ok(list.every((b) => b.has_end), '每个书签都应有配对的 end')
  assert.deepEqual(list.map((b) => b.paragraph), [1, 2])
  assert.ok(list[0].text.includes('本文件由'), `书签文本不对：${list[0].text}`)
})

test('★ id 不与既有书签冲突（重号会让 Word 认不出范围）', () => {
  const doc = DocxDocument.open(original)
  doc.insertBookmark({ paragraph: 0, name: 'a' })
  doc.insertBookmark({ paragraph: 1, name: 'b' })
  doc.insertBookmark({ paragraph: 2, name: 'c' })
  const xml = ZipPackage.open(doc.save()).readText('word/document.xml')
  const ids = [...xml.matchAll(/<w:bookmark(?:Start|End) w:id="(\d+)"/g)].map((m) => m[1])
  const startIds = [...xml.matchAll(/<w:bookmarkStart w:id="(\d+)"/g)].map((m) => m[1])
  assert.deepEqual(startIds, ['0', '1', '2'])
  assert.equal(new Set(startIds).size, startIds.length, '书签 id 不能重复')
  assert.equal(ids.length, 6, '三对标记共 6 个')
})

test('★ 书签名的校验：重名（不分大小写）/ 非法名 / 超长', () => {
  const doc = DocxDocument.open(original)
  doc.insertBookmark({ paragraph: 1, name: 'Intro' })
  assert.throws(() => doc.insertBookmark({ paragraph: 2, name: 'intro' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.insertBookmark({ paragraph: 2, name: '9bad' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.insertBookmark({ paragraph: 2, name: 'has space' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.insertBookmark({ paragraph: 2, name: '' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.insertBookmark({ paragraph: 2, name: 'x'.repeat(41) }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.insertBookmark({ paragraph: 99, name: 'ok' }), (e) => e.code === 'FILE_NOT_FOUND')
})

test('★ 空段落上插入书签也是正确的顺序（两处插入落在同一偏移）', () => {
  const doc = DocxDocument.open(original)
  // 第 0 段通常只有标题文本；构造一个真正空的段落：清空其内容
  const pkg = ZipPackage.open(original)
  const xml = pkg.readText('word/document.xml').replace(
    /(<w:p>(?:<w:pPr>[\s\S]*?<\/w:pPr>)?)([\s\S]*?)(<\/w:p>)/,
    '$1$3'
  )
  pkg.write('word/document.xml', xml)
  const empty = DocxDocument.open(pkg.toBuffer())
  empty.insertBookmark({ paragraph: 0, name: 'empty' })
  const out = ZipPackage.open(empty.save()).readText('word/document.xml')
  assert.ok(/<w:bookmarkStart w:id="0" w:name="empty"\/><w:bookmarkEnd w:id="0"\/>/.test(out), `空段落里的顺序不对：${out.slice(0, 400)}`)
  assert.equal(DocxDocument.open(empty.save()).validate().valid, true)
})

test('★ 插书签只改 document.xml，其余部件逐字节不变', () => {
  const before = new Map()
  const pkgBefore = ZipPackage.open(original)
  for (const n of pkgBefore.names()) before.set(n, pkgBefore.read(n))

  const doc = DocxDocument.open(original)
  doc.insertBookmark({ paragraph: 1, name: 'onlydoc' })
  const pkgAfter = ZipPackage.open(doc.save())
  const changed = [...before.keys()].filter((n) => !pkgAfter.read(n).equals(before.get(n)))
  assert.deepEqual(changed, ['word/document.xml'], `意外改动：${changed.join(', ')}`)
})

console.log('\n=== 23. 批注与修订读取 ===')

/** @returns {import('../lib/docx.js').DocxDocument|null} 真实 Word 复杂样本。 */
function complexDoc() {
  const file = join(here, 'fixtures', 'word-complex.docx')
  return existsSync(file) ? DocxDocument.open(readFileSync(file)) : null
}

test('★ 批注：作者、日期、正文、锚点段落', () => {
  const doc = complexDoc()
  if (!doc) return
  const result = doc.comments()
  assert.equal(result.count, 1)
  assert.equal(result.part, 'word/comments.xml')
  const c = result.comments[0]
  assert.ok(c.author.length > 0, '应有作者')
  assert.ok(c.date.startsWith('20'), `日期格式不对：${c.date}`)
  assert.ok(c.text.includes('法务复核'), `批注正文不对：${c.text}`)
  assert.ok(c.anchored_paragraph >= 0, '应能定位锚点段落')
})

test('★ 线程化批注的附加部件如实报告，不假装没有更多字段', () => {
  const doc = complexDoc()
  if (!doc) return
  const result = doc.comments()
  assert.ok(result.extended_parts.includes('word/commentsExtended.xml'), `应报告附加部件：${result.extended_parts}`)
  assert.ok(result.note.includes('未解析'), `note 应说明未解析：${result.note}`)
})

test('无批注文档返回空列表而不是抛错', () => {
  const result = DocxDocument.open(original).comments()
  assert.equal(result.part, null)
  assert.equal(result.count, 0)
  assert.deepEqual(result.comments, [])
})

test('★ 修订：插入与删除分别读出，段落标记增删单独标注', () => {
  const doc = complexDoc()
  if (!doc) return
  const result = doc.revisions()
  assert.equal(result.count, 2)
  assert.equal(result.by_type['插入'], 1)
  const inserted = result.revisions.find((r) => r.type === '插入')
  assert.ok(inserted.text.includes('修订模式插入'), `插入文字应读出：${inserted.text}`)
  const mark = result.revisions.find((r) => r.paragraph_mark === true)
  assert.ok(mark, '样本里有一条段落标记删除，应被标注')
  assert.equal(mark.text, '', '段落标记增删本身没有文字')
})

test('★ 无修订文档的统计为零', () => {
  const result = DocxDocument.open(original).revisions()
  assert.equal(result.count, 0)
  assert.deepEqual(result.by_type, {})
  assert.deepEqual(result.revisions, [])
})

test('★ 读取批注/修订不产生补丁（部件级逐字节一致）', () => {
  const file = join(here, 'fixtures', 'word-complex.docx')
  if (!existsSync(file)) return
  const bytes = readFileSync(file)
  const before = ZipPackage.open(bytes)
  const doc = DocxDocument.open(bytes)
  doc.comments()
  doc.revisions()
  const after = ZipPackage.open(doc.save())
  // 只读视图不应产生任何 XML 补丁：容器尾部的元数据可以重新生成，但**部件字节**必须一模一样
  for (const name of before.names()) {
    assert.ok(after.read(name).equals(before.read(name)), `部件被改动：${name}`)
  }
  assert.deepEqual(after.names().sort(), before.names().sort(), '不应新增或删除部件')
})

console.log('\n=== 24. 文档比较（段落级差异）===')

/**
 * 造一对「已知差异」的文档：在基准文档上做可数的改动。
 * @returns {{a: Buffer, b: Buffer}} 旧文档与新文档字节。
 */
function makeDiffPair() {
  const a = buildBasicDocx({ title: '季度销售报告' })
  const doc = DocxDocument.open(a)
  // 基准文档的段落：[标题][正文(含加粗)][正文结束。]，其后是表格
  doc.insertParagraph({ after: 2, text: '这是新增的一段。' })
  doc.updateParagraph({ index: 1, text: '本文件由 dsh-exp-office 生成（文字已改）。' })
  const endIndex = doc.paragraphs().paragraphs.findIndex((p) => p.text === '正文结束。')
  doc.setParagraphStyle({ index: endIndex, styleId: 'Heading2' })
  doc.updateTableCell({ table: 0, row: 1, column: 1, text: '一月' })
  const b = doc.save()
  return { a, b }
}

const pair = makeDiffPair()

test('内容完全相同 → identical，且各类差异均为 0', () => {
  const diff = DocxDocument.open(original).compare(DocxDocument.open(original))
  assert.equal(diff.identical, true)
  assert.equal(diff.paragraphs.total_added, 0)
  assert.equal(diff.paragraphs.total_removed, 0)
  assert.equal(diff.paragraphs.total_changed, 0)
  assert.equal(diff.tables.total_changed, 0)
  assert.deepEqual(diff.styles, { added: [], removed: [] })
  assert.deepEqual(diff.metadata.changed, [])
})

test('★ 新增段落只算新增，不会把后面的段落全报成修改（LCS 对齐）', () => {
  const diff = DocxDocument.open(pair.a).compare(DocxDocument.open(pair.b))
  const addedTexts = diff.paragraphs.added.map((p) => p.text)
  assert.ok(addedTexts.includes('这是新增的一段。'), `新增段落应被识别：${JSON.stringify(addedTexts)}`)
  assert.equal(diff.paragraphs.total_removed, 0, `不应产生删除：${JSON.stringify(diff.paragraphs.removed)}`)
  // 未改动的段落不能被报成修改
  const changedTexts = diff.paragraphs.changed.map((c) => c.text_before)
  assert.equal(changedTexts.includes('季度销售报告'), false, '未改动的段落不应出现在修改里')
})

test('★ 改文字被标为 text，改样式被标为 style', () => {
  const diff = DocxDocument.open(pair.a).compare(DocxDocument.open(pair.b))
  const textChange = diff.paragraphs.changed.find((c) => c.kind === 'text')
  assert.ok(textChange, '应有一处纯文字修改')
  assert.equal(textChange.text_after, '本文件由 dsh-exp-office 生成（文字已改）。')
  assert.equal(textChange.style_before, textChange.style_after)
  const styleChange = diff.paragraphs.changed.find((c) => c.kind === 'style')
  assert.ok(styleChange, '应有一处纯样式修改')
  assert.equal(styleChange.text_before, styleChange.text_after)
  assert.equal(styleChange.style_after, 'Heading2')
})

test('表格逐格变化被定位到行列', () => {
  const diff = DocxDocument.open(pair.a).compare(DocxDocument.open(pair.b))
  const cell = diff.tables.changed.find((c) => c.kind === 'cell')
  assert.ok(cell, `应有单元格变化：${JSON.stringify(diff.tables.changed)}`)
  assert.equal(cell.table, 0)
  assert.equal(cell.row, 2)
  assert.equal(cell.column, 2)
  assert.equal(cell.after, '一月')
})

test('样式集合差异与元数据差异各归各类', () => {
  const diff = DocxDocument.open(pair.a).compare(DocxDocument.open(pair.b))
  assert.ok(diff.styles.added.includes('Heading2'), `样式新增应含 Heading2：${JSON.stringify(diff.styles)}`)
  assert.deepEqual(diff.styles.removed, [])
  const b = DocxDocument.open(pair.b)
  b.updateParagraph({ index: 0, text: '标题 1' })
  assert.equal(diff.metadata.changed.length, 0, '样本元数据未改，不应报差异')
})

test('★ 比较是只读的：两份文件的部件一个字节都不变', () => {
  const bytesA = Buffer.from(pair.a)
  const bytesB = Buffer.from(pair.b)
  DocxDocument.open(bytesA).compare(DocxDocument.open(bytesB))
  assert.ok(bytesA.equals(pair.a), 'path_a 的字节不应被改动')
  assert.ok(bytesB.equals(pair.b), 'path_b 的字节不应被改动')
})

test('max_items 截断时 total_* 仍是真实总数，并给出警告', () => {
  const a = buildBasicDocx({ title: '基准' })
  const doc = DocxDocument.open(a)
  for (let i = 0; i < 12; i += 1) doc.insertParagraph({ after: 0, text: `新增段 ${i}` })
  const diff = DocxDocument.open(a).compare(DocxDocument.open(doc.save()), { maxItems: 5 })
  assert.equal(diff.truncated, true)
  assert.equal(diff.paragraphs.added.length, 5)
  assert.equal(diff.paragraphs.total_added, 12)
  assert.ok(diff.warnings.some((w) => w.code === 'COMPARE_TRUNCATED'))
})

test('段落过多时退化为按位置配对，并明确告知（不假装精确）', () => {
  const a = buildBasicDocx({ title: '基准' })
  const doc = DocxDocument.open(a)
  doc.insertParagraph({ after: 0, text: '中间插入的一段' })
  const diff = DocxDocument.open(a).compare(DocxDocument.open(doc.save()), { maxCells: 4 })
  assert.ok(diff.warnings.some((w) => w.code === 'COMPARE_ALIGNMENT_FALLBACK'), `应给出退化警告：${JSON.stringify(diff.warnings)}`)
})

test('可以关掉表格/样式/元数据比较', () => {
  const diff = DocxDocument.open(pair.a).compare(DocxDocument.open(pair.b), {
    includeTables: false,
    includeStyles: false,
    includeMetadata: false
  })
  assert.equal(diff.tables.total_changed, 0)
  assert.deepEqual(diff.styles, { added: [], removed: [] })
  assert.deepEqual(diff.metadata.changed, [])
})

test('参数校验：不是文档对象时明确拒绝', () => {
  assert.throws(
    () => DocxDocument.open(original).compare(null),
    (e) => e.code === 'INVALID_REQUEST'
  )
})

console.log(`\n结果：${passed} 通过，${failed} 失败\n`)
process.exit(failed === 0 ? 0 : 1)
