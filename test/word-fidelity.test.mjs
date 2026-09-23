/**
 * DOCX 保真度回归：以「真实 Microsoft Word 创作的复杂文档」为样本，
 * 断言编辑段落不会破坏批注、修订、目录、书签、脚注与超链接。
 *
 * 样本由 test/make-word-complex-fixture.ps1 生成；不存在则跳过（不失败），
 * 以便在没有 Office 的环境下仍可运行其它测试。
 *
 * 运行：node test/word-fidelity.test.mjs
 */
import assert from 'node:assert/strict'
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { DocxDocument } from '../lib/docx.js'
import { ZipPackage } from '../lib/ooxml.js'

const here = dirname(fileURLToPath(import.meta.url))
const fixture = join(here, 'fixtures', 'word-complex.docx')

if (!existsSync(fixture)) {
  console.log('⏭️  跳过：未找到 test/fixtures/word-complex.docx（需先在装有 Word 的机器上运行 make-word-complex-fixture.ps1）。')
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
const doc0 = DocxDocument.open(original)
const structure = doc0.structure()

console.log('\n=== 样本特征（真实 Word 创作）===')

test('样本含批注', () => {
  assert.equal(structure.has_comments, true)
  assert.ok(structure.comments >= 1, `批注数 ${structure.comments}`)
})

test('样本含未接受的修订', () => {
  assert.equal(structure.has_revisions, true)
})

test('样本含目录域', () => {
  assert.equal(structure.has_toc, true)
})

test('样本含书签', () => {
  assert.equal(structure.has_bookmarks, true)
})

test('样本含超链接', () => {
  assert.equal(structure.has_hyperlink_field, true)
})

test('样本含脚注', () => {
  assert.ok(structure.footnotes >= 1, `脚注数 ${structure.footnotes}`)
})

test('样本校验通过', () => {
  const report = doc0.validate()
  assert.equal(report.valid, true, JSON.stringify(report.checks.filter((c) => !c.ok)))
})

console.log('\n=== 编辑后的保真度 ===')

test('★ 插入段落只改 document.xml，其余部件逐字节不变', () => {
  const before = entriesOf(original)
  const doc = DocxDocument.open(original)
  doc.insertParagraph({ after: doc.bodyParagraphs().length - 1, text: '插件新增的段落' })
  const after = entriesOf(doc.save())

  assert.deepEqual([...after.keys()].sort(), [...before.keys()].sort(), '部件清单发生了变化')
  const changed = [...before.keys()].filter((n) => !after.get(n).equals(before.get(n)))
  assert.deepEqual(changed, ['word/document.xml'], `意外改动的部件：${changed.join(', ')}`)
})

test('★ 批注部件逐字节不变', () => {
  const before = entriesOf(original)
  const commentsPart = doc0.info.commentsPart
  assert.ok(commentsPart, '样本应含批注部件')
  const doc = DocxDocument.open(original)
  doc.insertParagraph({ after: 2, text: '插在中间的段落' })
  const after = entriesOf(doc.save())
  assert.ok(after.get(commentsPart).equals(before.get(commentsPart)), '批注部件被改动')
})

test('★ 修订标记未被破坏', () => {
  const before = ZipPackage.open(original).readText('word/document.xml')
  const doc = DocxDocument.open(original)
  // 改最后一段：修订标记在中间段落里，不应受影响
  doc.updateParagraph({ index: 5, text: '结尾段落被改写' })
  const after = ZipPackage.open(doc.save()).readText('word/document.xml')
  const insCount = (t) => (t.match(/<w:ins[ >]/g) ?? []).length
  const delCount = (t) => (t.match(/<w:del[ >]/g) ?? []).length
  assert.equal(insCount(after), insCount(before), '修订插入标记数量变化')
  assert.equal(delCount(after), delCount(before), '修订删除标记数量变化')
  assert.ok(insCount(before) >= 1, '样本应至少含一处修订插入')
})

test('★ 目录域未被破坏', () => {
  const before = ZipPackage.open(original).readText('word/document.xml')
  const doc = DocxDocument.open(original)
  doc.insertParagraph({ after: 0, text: '封面行' })
  const after = ZipPackage.open(doc.save()).readText('word/document.xml')
  const instr = (t) => (t.match(/<w:instrText[^>]*>([^<]*)<\/w:instrText>/g) ?? []).join('|')
  assert.equal(instr(after), instr(before), '域代码内容变化')
  assert.ok(/TOC/.test(instr(after)), '目录域指令丢失')
})

test('★ 书签标记未被破坏', () => {
  const before = ZipPackage.open(original).readText('word/document.xml')
  const doc = DocxDocument.open(original)
  doc.insertParagraph({ after: 1, text: '中间插入' })
  const after = ZipPackage.open(doc.save()).readText('word/document.xml')
  const marks = (t) => (t.match(/<w:bookmark(Start|End)[^>]*>/g) ?? []).join('|')
  assert.equal(marks(after), marks(before), '书签标记变化')
})

test('★ 超链接关系未被破坏', () => {
  const before = entriesOf(original)
  const doc = DocxDocument.open(original)
  doc.insertParagraph({ after: 3, text: '插入行' })
  const after = entriesOf(doc.save())
  assert.ok(after.get('word/_rels/document.xml.rels').equals(before.get('word/_rels/document.xml.rels')), '关系部件被改动')
})

test('★ 脚注与尾注部件未被破坏', () => {
  const before = entriesOf(original)
  const doc = DocxDocument.open(original)
  doc.updateParagraph({ index: 5, text: '只改这一句' })
  const after = entriesOf(doc.save())
  for (const [name, bytes] of before) {
    if (!/^word\/(foot|end)notes\.xml$/.test(name)) continue
    assert.ok(after.get(name).equals(bytes), `${name} 被改动`)
  }
})

test('★ 编辑后批注、修订、目录仍然可被识别', () => {
  const doc = DocxDocument.open(original)
  // 先改安全的空段落，再插入——插入会使后续段落下标位移
  doc.updateParagraph({ index: 5, text: '改写结尾' })
  doc.insertParagraph({ after: 2, text: '新段落' })
  const reopened = DocxDocument.open(doc.save())
  const s = reopened.structure()
  assert.equal(s.has_comments, true, '批注丢失')
  assert.equal(s.comments, structure.comments, '批注数量变化')
  assert.equal(s.has_revisions, true, '修订标记丢失')
  assert.equal(s.has_toc, true, '目录域丢失')
  assert.equal(s.has_bookmarks, true, '书签丢失')
  assert.equal(s.footnotes, structure.footnotes, '脚注数量变化')
})

test('★ 编辑后 validate_docx 仍然通过', () => {
  const doc = DocxDocument.open(original)
  doc.updateParagraph({ index: 5, text: '改写结尾' })
  doc.insertParagraph({ after: 2, text: '新段落' })
  const report = DocxDocument.open(doc.save()).validate()
  assert.equal(report.valid, true, JSON.stringify(report.checks.filter((c) => !c.ok)))
})

test('多次连续编辑后复杂部件仍然完整', () => {
  let buffer = original
  for (let i = 0; i < 4; i += 1) {
    const doc = DocxDocument.open(buffer)
    doc.insertParagraph({ after: 0, text: `第 ${i} 轮插入` })
    // 每轮重新定位：最后一段始终是那个安全的空段落
    doc.updateParagraph({ index: doc.bodyParagraphs().length - 1, text: `第 ${i} 轮改写` })
    buffer = doc.save()
  }
  const s = DocxDocument.open(buffer).structure()
  assert.equal(s.has_comments, true)
  assert.equal(s.has_revisions, true)
  assert.equal(s.has_toc, true)
  assert.equal(s.has_bookmarks, true)
})

test('paragraphSafety 能指出哪些段落可安全改写', () => {
  const doc = DocxDocument.open(original)
  const unsafe = doc.paragraphSafety(4)
  assert.equal(unsafe.safe, false)
  assert.ok(unsafe.markup.includes('commentReference'), `实际标记：${unsafe.markup.join(', ')}`)
  assert.ok(unsafe.labels.includes('批注引用'), `实际标签：${unsafe.labels.join(', ')}`)
  const safe = doc.paragraphSafety(5)
  assert.equal(safe.safe, true)
  assert.deepEqual(safe.markup, [])
})

test('目录段落含域代码，被识别为不可安全改写', () => {
  const toc = DocxDocument.open(original).paragraphSafety(0)
  assert.equal(toc.safe, false)
  assert.ok(toc.markup.includes('instrText') || toc.markup.includes('fldChar'), `实际：${toc.markup.join(', ')}`)
})

writeFileSync(
  join(here, 'fixtures', 'word-complex-edited.docx'),
  (() => {
    const doc = DocxDocument.open(original)
    // 安全编辑：只新增段落 + 改写不含保护标记的空段落
    doc.insertParagraph({ after: doc.bodyParagraphs().length - 1, text: '本段由 dsh-exp-office 插件插入。' })
    doc.updateParagraph({ index: 5, text: '风险项已确认：预算与排期均无阻塞。' })
    return doc.save()
  })()
)

test('★ 改写含批注锚点的段落会被拒绝（而不是静默孤立批注）', () => {
  const doc = DocxDocument.open(original)
  // 样本的第 4 段同时含书签、批注锚点与修订插入标记
  try {
    doc.updateParagraph({ index: 4, text: '试图整段替换' })
    assert.fail('应当因保护性检查而抛出')
  } catch (err) {
    assert.equal(err.code, 'UNSUPPORTED_FEATURE')
    assert.equal(err.needsConfirmation, true)
    assert.ok(err.details.markup.length > 0, '应列出命中的标记')
    assert.ok(err.solution.includes('allow_markup_loss'), `解决建议不够可操作：${err.solution}`)
  }
})

test('★ 删除含批注锚点的段落同样被拒绝', () => {
  const doc = DocxDocument.open(original)
  assert.throws(() => doc.deleteParagraph({ index: 4 }), (e) => e.code === 'UNSUPPORTED_FEATURE')
})

test('显式 allowMarkupLoss 后才允许改写', () => {
  const doc = DocxDocument.open(original)
  const result = doc.updateParagraph({ index: 4, text: '明知会丢标记', allowMarkupLoss: true })
  assert.equal(result.index, 4)
  assert.equal(DocxDocument.open(doc.save()).paragraphs().paragraphs[4].text, '明知会丢标记')
})

test('★ 不含保护标记的段落仍可正常改写（保护检查不误伤）', () => {
  const doc = DocxDocument.open(original)
  // 第 5 段是空段落，不含任何标记
  const result = doc.updateParagraph({ index: 5, text: '安全的改写' })
  assert.equal(result.index, 5)
  const reopened = DocxDocument.open(doc.save())
  assert.equal(reopened.paragraphs().paragraphs[5].text, '安全的改写')
  assert.equal(reopened.structure().has_comments, true, '批注仍应完好')
})

console.log(`\n结果：${passed} 通过，${failed} 失败`)
console.log('输出文件：test/fixtures/word-complex-edited.docx\n')
process.exit(failed === 0 ? 0 : 1)
