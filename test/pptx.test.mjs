/**
 * PPTX 适配器测试（阶段 4：读取与结构识别）。
 *
 * 样本由真实 Microsoft PowerPoint 创作（test/make-pptx-fixture.ps1）；
 * 不存在则跳过（不失败），以便在没有 Office 的环境下仍可运行其它测试。
 *
 * 运行：node test/pptx.test.mjs
 */
import assert from 'node:assert/strict'
import zlib from 'node:zlib'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { PptxPresentation } from '../lib/pptx.js'
import { ZipPackage, XmlDoc, findAll } from '../lib/ooxml.js'
import { Workbook } from '../lib/xlsx.js'
import { buildBlankPptx } from '../lib/pptx-template.js'

const here = dirname(fileURLToPath(import.meta.url))
const fixture = join(here, 'fixtures', 'deck.pptx')

if (!existsSync(fixture)) {
  console.log('⏭️  跳过：未找到 test/fixtures/deck.pptx（需先在装有 PowerPoint 的机器上运行 make-pptx-fixture.ps1）。')
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

const original = readFileSync(fixture)

console.log('\n=== 1. 打开与结构 ===')

test('打开真实 PowerPoint 演示文稿', () => {
  const p = PptxPresentation.open(original)
  assert.equal(p.info.slides.length, 4)
})

test('识别幻灯片尺寸为 16:9', () => {
  const size = PptxPresentation.open(original).structure().slide_size
  assert.equal(size.aspect, '16:9')
  assert.equal(size.width_inch, 13.33)
  assert.equal(size.height_inch, 7.5)
})

test('识别母版、版式与主题', () => {
  const s = PptxPresentation.open(original).structure()
  assert.ok(s.masters >= 1, `母版数 ${s.masters}`)
  assert.ok(s.layouts >= 1, `版式数 ${s.layouts}`)
  assert.ok(s.themes >= 1, `主题数 ${s.themes}`)
})

test('★ 识别隐藏幻灯片（标记在幻灯片部件根元素的 show 属性上）', () => {
  const s = PptxPresentation.open(original).structure()
  assert.deepEqual(s.hidden_slides, [3], `实际隐藏页：${JSON.stringify(s.hidden_slides)}`)
})

test('读取文档元数据', () => {
  const meta = PptxPresentation.open(original).metadata()
  assert.equal(meta.application, 'Microsoft Office PowerPoint')
  assert.ok(meta.created)
})

console.log('\n=== 2. 幻灯片内容 ===')

test('读取标题页的标题与副标题', () => {
  const slide = PptxPresentation.open(original).readSlide(0)
  assert.equal(slide.title, '季度业务报告')
  const labels = slide.shapes.map((s) => s.placeholder_label)
  assert.ok(labels.includes('居中标题'), `实际占位符：${labels.join(', ')}`)
  assert.ok(labels.includes('副标题'))
})

test('读取项目符号正文并保留多行', () => {
  const slide = PptxPresentation.open(original).readSlide(1)
  assert.equal(slide.title, '本季度进展')
  const body = slide.shapes.find((s) => s.placeholder_label === '正文')
  assert.ok(body, '应找到正文占位符')
  const lines = body.text.split('\n')
  assert.equal(lines.length, 3, `实际行数 ${lines.length}：${body.text}`)
  assert.ok(lines[0].includes('XLSX'))
  assert.ok(lines[2].includes('PPTX'))
})

test('识别表格形状', () => {
  const slide = PptxPresentation.open(original).readSlide(2)
  assert.ok(slide.shapes.some((s) => s.type === 'table'), `实际形状：${slide.shapes.map((s) => s.type).join(', ')}`)
})

test('读取演讲者备注，且不混入幻灯编号占位符', () => {
  const slide = PptxPresentation.open(original).readSlide(1)
  assert.ok(slide.notes, '应有备注')
  assert.ok(slide.notes.includes('讲稿'), slide.notes)
  assert.ok(!/^\d+$/m.test(slide.notes.trim()), `备注混入了编号：${JSON.stringify(slide.notes)}`)
})

test('★ 备注只取 body 占位符（不取 sldImg / sldNum）', () => {
  const s = PptxPresentation.open(original).structure()
  assert.equal(s.notes_slides, 1)
  assert.equal(s.has_speaker_notes, true)
  // 备注页有 3 个占位符（sldImg / body / sldNum），只有 body 是讲稿
  const withNotes = PptxPresentation.open(original).slides().slides.filter((x) => x.notes)
  assert.equal(withNotes.length, 1)
  assert.equal(withNotes[0].notes, '讲稿：重点说明保真度验证方法与真实软件复核。')
})

test('无标题页返回 null 而不是空串', () => {
  const slide = PptxPresentation.open(original).readSlide(2)
  assert.equal(slide.title, null)
})

test('幻灯片越界时给出总数', () => {
  const p = PptxPresentation.open(original)
  try {
    p.readSlide(99)
    assert.fail('应当抛出')
  } catch (err) {
    assert.equal(err.code, 'FILE_NOT_FOUND')
    assert.equal(err.details.slide_count, 4)
  }
})

console.log('\n=== 3. 全文与批量读取 ===')

test('提取全文包含各页标题与正文', () => {
  const { text } = PptxPresentation.open(original).text()
  assert.ok(text.includes('季度业务报告'))
  assert.ok(text.includes('本季度进展'))
  assert.ok(text.includes('第 1 张幻灯片'))
})

test('全文提取支持截断', () => {
  const result = PptxPresentation.open(original).text(20)
  assert.equal(result.text.length, 20)
  assert.equal(result.truncated, true)
})

test('slides() 可只返回摘要而不含形状明细', () => {
  const { slides } = PptxPresentation.open(original).slides({ includeShapes: false })
  assert.equal(slides.length, 4)
  assert.ok(slides.every((s) => s.shape_count >= 0 && s.shapes === undefined))
})

console.log('\n=== 4. 校验 ===')

test('样本校验通过', () => {
  const report = PptxPresentation.open(original).validate()
  assert.equal(report.valid, true, JSON.stringify(report.checks.filter((c) => !c.ok)))
})

test('校验报告包含必需检查项', () => {
  const names = PptxPresentation.open(original).validate().checks.map((c) => c.name)
  for (const expected of ['ZIP 包结构', '必需部件齐全', '幻灯片部件可解析', '幻灯片关系有效', '母版与主题保留']) {
    assert.ok(names.includes(expected), `缺少检查项：${expected}`)
  }
})

test('★ 明确列出需渲染才能判定的项，不假装通过', () => {
  const report = PptxPresentation.open(original).validate()
  assert.ok(report.not_checked.length >= 4)
  assert.ok(report.not_checked.some((c) => c.name === '文本溢出与裁剪'))
  assert.ok(report.not_checked.every((c) => typeof c.reason === 'string' && c.reason.length > 0))
})

test('★ 检出悬空的图片引用', () => {
  const pkg = ZipPackage.open(original)
  const xml = pkg.readText('ppt/slides/slide1.xml').replace('<p:cSld>', '<p:cSld><p:pic><p:blipFill><a:blip r:embed="rId999"/></p:blipFill></p:pic>')
  pkg.write('ppt/slides/slide1.xml', xml)
  const report = PptxPresentation.open(pkg.toBuffer()).validate()
  assert.equal(report.valid, false)
  const check = report.checks.find((c) => c.name === '图片/图表引用可解析')
  assert.equal(check.ok, false)
  assert.ok(check.detail.includes('rId999'), check.detail)
})

console.log('\n=== 5. 安全与错误处理 ===')

test('加密的 OOXML（OLE 容器）给出 PASSWORD_REQUIRED', () => {
  const ole = Buffer.alloc(64)
  ole.writeUInt32LE(0xe011cfd0, 0)
  assert.throws(() => PptxPresentation.open(ole), (e) => e.code === 'PASSWORD_REQUIRED')
})

test('非 PPTX 的 ZIP 给出 CORRUPTED_DOCUMENT', () => {
  const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { maxEntries: 8, maxEntryBytes: 1000, maxTotalBytes: 10000, maxRatio: 200 })
  pkg.write('hello.txt', 'hi')
  assert.throws(() => PptxPresentation.open(pkg.toBuffer()), (e) => e.code === 'CORRUPTED_DOCUMENT')
})

test('空文件被拒绝', () => {
  assert.throws(() => PptxPresentation.open(Buffer.alloc(0)), (e) => e.code === 'CORRUPTED_DOCUMENT')
})

test('PPTX 同样受 ZIP 炸弹上限约束', () => {
  assert.throws(() => PptxPresentation.open(original, { maxEntries: 2 }), (e) => e.code === 'MEMORY_LIMIT')
})

console.log('\n=== 6. 写入：幻灯片文本 ===')

test('改写标题形状的文本', () => {
  const p = PptxPresentation.open(original)
  const result = p.updateSlideText({ index: 0, shape: 0, text: '新的报告标题' })
  assert.equal(result.slide, 0)
  const reopened = PptxPresentation.open(p.save())
  assert.equal(reopened.readSlide(0).title, '新的报告标题')
})

test('★ 形状下标与读取顺序一致（读写共用同一枚举器）', () => {
  const p = PptxPresentation.open(original)
  const before = p.readSlide(1)
  const bodyIndex = before.shapes.findIndex((s) => s.placeholder_label === '正文')
  assert.ok(bodyIndex >= 0, '应先能读到正文形状')
  p.updateSlideText({ index: 1, shape: bodyIndex, text: '写入到正文形状' })
  const after = PptxPresentation.open(p.save()).readSlide(1)
  assert.equal(after.shapes[bodyIndex].text, '写入到正文形状', '写入的形状与读到的不是同一个')
  assert.equal(after.title, before.title, '标题不应被改动')
})

test('★ 多行文本分段，且不凭空造出段落属性', () => {
  const p = PptxPresentation.open(original)
  const bodyIndex = p.readSlide(1).shapes.findIndex((s) => s.placeholder_label === '正文')
  const result = p.updateSlideText({ index: 1, shape: bodyIndex, text: '第一条\n第二条\n第三条' })
  assert.equal(result.paragraphs, 3)
  // 本样本的段落属性来自版式的 <a:lstStyle>，源段落本身没有 <a:pPr>，
  // 因此这里正确地没有可复用的段落属性，也不应凭空造一个出来。
  assert.equal(result.reused_paragraph_properties, false)

  // 段落数用读回的 paragraph_count 断言，而不是数整份 XML 里的 <a:p>
  // （标题形状还有 1 段，整份计数会把它算进来）
  const after = PptxPresentation.open(p.save())
  assert.equal(after.readSlide(1).shapes[bodyIndex].paragraph_count, 3)

  const xml = ZipPackage.open(p.save()).readText('ppt/slides/slide2.xml')
  assert.equal((xml.match(/<a:pPr/g) ?? []).length, 0, '不应凭空造出 <a:pPr>')

  const lines = after.readSlide(1).shapes[bodyIndex].text.split('\n')
  assert.deepEqual(lines, ['第一条', '第二条', '第三条'])
})

test('★ 源段落带 <a:pPr> 时会被复用（保住项目符号层级）', () => {
  const pkg = ZipPackage.open(original)
  // 定位第二个 txBody（正文），在其第一段前注入显式层级属性，模拟真实模板
  const raw = pkg.readText('ppt/slides/slide2.xml')
  const bodyStart = raw.indexOf('<p:txBody>', raw.indexOf('<p:txBody>') + 1)
  const injected = `${raw.slice(0, bodyStart)}${raw.slice(bodyStart).replace('<a:p>', '<a:p><a:pPr lvl="1" marL="742950" indent="-342900"/>')}`
  pkg.write('ppt/slides/slide2.xml', injected)

  const p = PptxPresentation.open(pkg.toBuffer())
  const bodyIndex = p.readSlide(1).shapes.findIndex((s) => s.placeholder_label === '正文')
  const result = p.updateSlideText({ index: 1, shape: bodyIndex, text: '甲\n乙' })
  assert.equal(result.reused_paragraph_properties, true, '应复用原 <a:pPr>')

  const out = ZipPackage.open(p.save()).readText('ppt/slides/slide2.xml')
  assert.equal((out.match(/lvl="1"/g) ?? []).length, 2, '两段都应保留 lvl="1"')
})

test('★ 只改目标幻灯片部件，其余部件逐字节不变', () => {
  const before = new Map()
  const pkgBefore = ZipPackage.open(original)
  for (const n of pkgBefore.names()) before.set(n, pkgBefore.read(n))

  const p = PptxPresentation.open(original)
  p.updateSlideText({ index: 1, shape: 0, text: '只改这一页' })
  const pkgAfter = ZipPackage.open(p.save())

  const changed = [...before.keys()].filter((n) => !pkgAfter.read(n).equals(before.get(n)))
  assert.deepEqual(changed, ['ppt/slides/slide2.xml'], `意外改动：${changed.join(', ')}`)
})

test('写入后校验仍然通过', () => {
  const p = PptxPresentation.open(original)
  p.updateSlideText({ index: 1, shape: 1, text: '新正文' })
  const report = PptxPresentation.open(p.save()).validate()
  assert.equal(report.valid, true, JSON.stringify(report.checks.filter((c) => !c.ok)))
})

test('母版、版式与备注不受影响', () => {
  const p = PptxPresentation.open(original)
  p.updateSlideText({ index: 1, shape: 1, text: '新正文' })
  const reopened = PptxPresentation.open(p.save())
  const s = reopened.structure()
  assert.ok(s.masters >= 1 && s.layouts >= 1 && s.themes >= 1)
  assert.equal(reopened.readSlide(1).notes, '讲稿：重点说明保真度验证方法与真实软件复核。')
})

test('对没有文本体的形状写入会被拒绝', () => {
  const p = PptxPresentation.open(original)
  const tableIndex = p.readSlide(2).shapes.findIndex((s) => s.type === 'table')
  assert.ok(tableIndex >= 0)
  try {
    p.updateSlideText({ index: 2, shape: tableIndex, text: 'x' })
    assert.fail('应当抛出')
  } catch (err) {
    assert.equal(err.code, 'UNSUPPORTED_FEATURE')
    assert.equal(err.details.shape_type, 'table')
  }
})

test('★ 含超链接的形状默认拒绝改写', () => {
  const pkg = ZipPackage.open(original)
  const xml = pkg.readText('ppt/slides/slide2.xml').replace(
    '<a:t>本季度进展</a:t>',
    '<a:t>本季度进展</a:t></a:r><a:r><a:rPr><a:hlinkClick r:id="rId1"/></a:rPr><a:t>链接</a:t>'
  )
  pkg.write('ppt/slides/slide2.xml', xml)
  const p = PptxPresentation.open(pkg.toBuffer())
  try {
    p.updateSlideText({ index: 1, shape: 0, text: '试图改写' })
    assert.fail('应当因保护性检查而抛出')
  } catch (err) {
    assert.equal(err.code, 'UNSUPPORTED_FEATURE')
    assert.equal(err.needsConfirmation, true)
    assert.ok(err.solution.includes('allow_markup_loss'), err.solution)
  }
})

test('显式 allowMarkupLoss 后才允许改写含链接的形状', () => {
  const pkg = ZipPackage.open(original)
  const xml = pkg.readText('ppt/slides/slide2.xml').replace(
    '<a:t>本季度进展</a:t>',
    '<a:t>本季度进展</a:t></a:r><a:r><a:rPr><a:hlinkClick r:id="rId1"/></a:rPr><a:t>链接</a:t>'
  )
  pkg.write('ppt/slides/slide2.xml', xml)
  const p = PptxPresentation.open(pkg.toBuffer())
  const result = p.updateSlideText({ index: 1, shape: 0, text: '明知会丢链接', allowMarkupLoss: true })
  assert.equal(result.slide, 1)
})

test('幻灯片或形状越界时给出可用规模', () => {
  const p = PptxPresentation.open(original)
  assert.throws(() => p.updateSlideText({ index: 99, shape: 0, text: 'x' }), (e) => e.code === 'FILE_NOT_FOUND')
  try {
    p.updateSlideText({ index: 0, shape: 99, text: 'x' })
    assert.fail('应当抛出')
  } catch (err) {
    assert.equal(err.details.shape_count > 0, true)
  }
})

test('连续多轮写入保持稳定', () => {
  let buffer = original
  for (let i = 0; i < 4; i += 1) {
    const p = PptxPresentation.open(buffer)
    p.updateSlideText({ index: 0, shape: 0, text: `第 ${i} 轮` })
    buffer = p.save()
  }
  const p = PptxPresentation.open(buffer)
  assert.equal(p.readSlide(0).title, '第 3 轮')
  assert.equal(p.validate().valid, true)
})

console.log('\n=== 7. 幻灯片增删与排序 ===')

test('新增幻灯片追加到末尾', () => {
  const p = PptxPresentation.open(original)
  const result = p.addSlide()
  assert.equal(result.slide_count, 5)
  assert.equal(result.index, 4)
  const reopened = PptxPresentation.open(p.save())
  assert.equal(reopened.info.slides.length, 5)
  assert.ok(reopened.validate().valid)
})

test('★ 新增幻灯片继承参照页的版式', () => {
  const p = PptxPresentation.open(original)
  const result = p.addSlide({ layoutOf: 1 })
  assert.ok(result.layout, '应返回继承的版式路径')
  const reopened = PptxPresentation.open(p.save())
  assert.equal(reopened.readSlide(4).layout, reopened.readSlide(1).layout, '新页应与参照页同版式')
})

test('在指定位置插入新幻灯片', () => {
  const p = PptxPresentation.open(original)
  const result = p.addSlide({ after: 0 })
  assert.equal(result.index, 1)
  const titles = PptxPresentation.open(p.save()).slides().slides.map((s) => s.title)
  assert.equal(titles[0], '季度业务报告')
  assert.equal(titles[1], null, '新插入的空页应无标题')
  assert.equal(titles[2], '本季度进展')
})

test('after=-1 插到最前面', () => {
  const p = PptxPresentation.open(original)
  const result = p.addSlide({ after: -1 })
  assert.equal(result.index, 0)
  assert.equal(PptxPresentation.open(p.save()).slides().slides[1].title, '季度业务报告')
})

test('★ 新增幻灯片后包结构完整（内容类型与关系都补齐）', () => {
  const p = PptxPresentation.open(original)
  const result = p.addSlide()
  const pkg = ZipPackage.open(p.save())
  assert.ok(pkg.has(result.part), '缺少幻灯片部件')
  assert.ok(pkg.has(`ppt/slides/_rels/slide5.xml.rels`), '缺少幻灯片关系部件')
  assert.ok(pkg.readText('[Content_Types].xml').includes(`/ppt/slides/slide5.xml`), '缺少内容类型 Override')
  assert.ok(pkg.readText('ppt/_rels/presentation.xml.rels').includes('slides/slide5.xml'), '缺少演示文稿关系')
  assert.ok(pkg.readText('ppt/presentation.xml').includes(result.relationship_id), 'sldIdLst 缺少引用')
})

test('删除幻灯片并清理孤儿部件', () => {
  const p = PptxPresentation.open(original)
  const result = p.deleteSlide({ index: 3 })
  assert.equal(result.slide_count, 3)
  const pkg = ZipPackage.open(p.save())
  assert.equal(pkg.has('ppt/slides/slide4.xml'), false, '幻灯片部件应被删除')
  assert.equal(pkg.has('ppt/slides/_rels/slide4.xml.rels'), false, '关系部件应被删除')
  assert.equal(pkg.readText('[Content_Types].xml').includes('/ppt/slides/slide4.xml'), false, '内容类型声明应被移除')
  assert.ok(PptxPresentation.open(p.save()).validate().valid)
})

test('拒绝删除最后一张幻灯片', () => {
  const p = PptxPresentation.open(original)
  p.deleteSlide({ index: 3 })
  p.deleteSlide({ index: 2 })
  p.deleteSlide({ index: 1 })
  assert.throws(() => p.deleteSlide({ index: 0 }), (e) => e.code === 'INVALID_REQUEST')
})

test('调整幻灯片顺序', () => {
  const p = PptxPresentation.open(original)
  const before = p.slides().slides.map((s) => s.title)
  const result = p.moveSlide({ from: 0, to: 2 })
  assert.equal(result.changed, true)
  const after = PptxPresentation.open(p.save()).slides().slides.map((s) => s.title)
  assert.deepEqual(after, [before[1], before[2], before[0], before[3]])
})

test('顺序不变时为空操作', () => {
  const p = PptxPresentation.open(original)
  assert.equal(p.moveSlide({ from: 1, to: 1 }).changed, false)
})

test('排序越界被拒绝', () => {
  const p = PptxPresentation.open(original)
  assert.throws(() => p.moveSlide({ from: 0, to: 99 }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => p.moveSlide({ from: -1, to: 0 }), (e) => e.code === 'INVALID_REQUEST')
})

test('★ 增删排序后母版、主题与备注不受影响', () => {
  const p = PptxPresentation.open(original)
  p.addSlide()
  p.moveSlide({ from: 0, to: 1 })
  const reopened = PptxPresentation.open(p.save())
  const s = reopened.structure()
  assert.ok(s.masters >= 1 && s.layouts >= 1 && s.themes >= 1)
  const withNotes = reopened.slides().slides.find((x) => x.notes)
  assert.equal(withNotes.notes, '讲稿：重点说明保真度验证方法与真实软件复核。')
})

test('★ 组合操作后仍然校验通过且幻灯片数正确', () => {
  const p = PptxPresentation.open(original)
  p.addSlide({ after: 0 })
  p.deleteSlide({ index: 4 })
  p.moveSlide({ from: 0, to: 2 })
  // 操作后顺序为 [新增空页, 本季度进展, 季度业务报告, 表格页]；标题在 2 号位
  const titleIndex = PptxPresentation.open(p.save()).slides().slides.findIndex((s) => s.title === '季度业务报告')
  assert.equal(titleIndex, 2, '标题页应被移到 2 号位')
  p.updateSlideText({ index: titleIndex, shape: 0, text: '组合操作后的标题' })

  const reopened = PptxPresentation.open(p.save())
  assert.equal(reopened.info.slides.length, 4)
  assert.equal(reopened.validate().valid, true)
  assert.equal(reopened.readSlide(titleIndex).title, '组合操作后的标题')
})

console.log('\n=== 8. 幻灯片图片插入 ===')

/**
 * 生成一张合法的 RGB PNG。
 * @param {number} width - 宽。
 * @param {number} height - 高。
 * @returns {Buffer} PNG 字节。
 */
function makePng(width, height) {
  const raw = Buffer.alloc((width * 3 + 1) * height)
  let o = 0
  for (let y = 0; y < height; y += 1) {
    raw[o] = 0
    o += 1
    for (let x = 0; x < width; x += 1) {
      raw[o] = 31
      raw[o + 1] = 119
      raw[o + 2] = 180
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

const png = makePng(200, 100)

test('插入图片并补齐媒体部件、内容类型与关系', () => {
  const p = PptxPresentation.open(original)
  const result = p.insertSlideImage({ index: 1, data: png, extension: 'png', altText: '测试图' })
  assert.equal(result.media_part, 'ppt/media/image1.png')
  const pkg = ZipPackage.open(p.save())
  assert.ok(pkg.has('ppt/media/image1.png'), '缺少媒体部件')
  assert.ok(pkg.read('ppt/media/image1.png').equals(png), '媒体内容不一致')
  assert.ok(pkg.readText('[Content_Types].xml').includes('Extension="png"'), '缺少 png 内容类型声明')
  assert.ok(pkg.readText('ppt/slides/_rels/slide2.xml.rels').includes('relationships/image'), '缺少 image 关系')
})

test('★ 插入后校验通过（证明关系链完整）', () => {
  const p = PptxPresentation.open(original)
  p.insertSlideImage({ index: 1, data: png, extension: 'png' })
  const report = PptxPresentation.open(p.save()).validate()
  assert.equal(report.valid, true, JSON.stringify(report.checks.filter((c) => !c.ok)))
})

test('未指定尺寸时用原始像素并居中', () => {
  const p = PptxPresentation.open(original)
  const result = p.insertSlideImage({ index: 1, data: png, extension: 'png' })
  assert.deepEqual(result.display_size_px, { width: 200, height: 100 })
  // 幻灯片 12192000 × 6858000 EMU，图片应在水平/垂直方向居中
  assert.equal(result.offset_emu.x, Math.round((12192000 - 200 * 9525) / 2))
  assert.equal(result.offset_emu.y, Math.round((6858000 - 100 * 9525) / 2))
})

test('只给宽度时按比例缩放', () => {
  const p = PptxPresentation.open(original)
  const result = p.insertSlideImage({ index: 1, data: png, extension: 'png', widthPx: 400 })
  assert.deepEqual(result.display_size_px, { width: 400, height: 200 })
})

test('可指定位置', () => {
  const p = PptxPresentation.open(original)
  const result = p.insertSlideImage({ index: 1, data: png, extension: 'png', leftPx: 10, topPx: 20 })
  assert.deepEqual(result.offset_emu, { x: 10 * 9525, y: 20 * 9525 })
})

test('★ 形状列表新增一个 picture 项', () => {
  const p = PptxPresentation.open(original)
  const before = p.readSlide(1).shape_count
  p.insertSlideImage({ index: 1, data: png, extension: 'png' })
  const after = PptxPresentation.open(p.save()).readSlide(1)
  assert.equal(after.shape_count, before + 1)
  assert.ok(after.shapes.some((s) => s.type === 'picture'), `实际形状：${after.shapes.map((s) => s.type).join(', ')}`)
})

test('多张图片各自获得独立部件与关系', () => {
  const p = PptxPresentation.open(original)
  const r1 = p.insertSlideImage({ index: 1, data: png, extension: 'png' })
  const r2 = p.insertSlideImage({ index: 1, data: makePng(50, 50), extension: 'png' })
  assert.notEqual(r1.media_part, r2.media_part)
  assert.notEqual(r1.relationship_id, r2.relationship_id)
  const pkg = ZipPackage.open(p.save())
  assert.ok(pkg.has('ppt/media/image1.png'))
  assert.ok(pkg.has('ppt/media/image2.png'))
})

test('★ 只改目标幻灯片与其关系，其余部件逐字节不变', () => {
  const before = new Map()
  const pkgBefore = ZipPackage.open(original)
  for (const n of pkgBefore.names()) before.set(n, pkgBefore.read(n))

  const p = PptxPresentation.open(original)
  p.insertSlideImage({ index: 1, data: png, extension: 'png' })
  const pkgAfter = ZipPackage.open(p.save())

  const added = pkgAfter.names().filter((n) => !before.has(n)).sort()
  assert.deepEqual(added, ['ppt/media/image1.png'], `新增部件异常：${added.join(', ')}`)
  const changed = [...before.keys()].filter((n) => !pkgAfter.read(n).equals(before.get(n))).sort()
  assert.deepEqual(changed, ['[Content_Types].xml', 'ppt/slides/_rels/slide2.xml.rels', 'ppt/slides/slide2.xml'], `意外改动：${changed.join(', ')}`)
  // 母版、主题、其它幻灯片必须逐字节不变
  for (const n of before.keys()) {
    if (n.startsWith('ppt/slideMasters/') || n.startsWith('ppt/theme/') || n === 'ppt/slides/slide1.xml') {
      assert.ok(pkgAfter.read(n).equals(before.get(n)), `${n} 被改动`)
    }
  }
})

test('幻灯片越界时给出总数', () => {
  const p = PptxPresentation.open(original)
  assert.throws(() => p.insertSlideImage({ index: 99, data: png, extension: 'png' }), (e) => e.code === 'FILE_NOT_FOUND')
})

console.log('\n=== 9. 幻灯片表格 ===')

test('★ 读取 PowerPoint 原生表格的行列内容', () => {
  // 样本第 3 页的表格由真实 PowerPoint 创建
  const t = PptxPresentation.open(original).readSlide(2).shapes.find((s) => s.type === 'table')
  assert.ok(t, '应识别出表格形状')
  assert.ok(t.table, '应带表格内容')
  assert.equal(t.table.row_count, 3)
  assert.equal(t.table.column_count, 3)
  assert.equal(t.table.rows[0][0], '模块')
  assert.equal(t.table.rows[1][2], '46')
})

test('插入表格并读回内容', () => {
  const p = PptxPresentation.open(original)
  const result = p.addSlideTable({
    index: 0,
    rows: [
      ['模块', '状态', '用例数'],
      ['XLSX', '已完成', '46'],
      ['DOCX', '已完成', '73']
    ]
  })
  assert.equal(result.rows, 3)
  assert.equal(result.columns, 3)

  const slide = PptxPresentation.open(p.save()).readSlide(0)
  const table = slide.shapes.find((s) => s.type === 'table')
  assert.ok(table, '插入后应能读到表格')
  assert.equal(table.table.rows[0][0], '模块')
  assert.equal(table.table.rows[2][2], '73')
})

test('★ 插入表格后校验通过', () => {
  const p = PptxPresentation.open(original)
  p.addSlideTable({ index: 0, rows: [['A', 'B'], ['1', '2']] })
  const report = PptxPresentation.open(p.save()).validate()
  assert.equal(report.valid, true, JSON.stringify(report.checks.filter((c) => !c.ok)))
})

test('★ 只改目标幻灯片，其余部件逐字节不变', () => {
  const before = new Map()
  const pkgBefore = ZipPackage.open(original)
  for (const n of pkgBefore.names()) before.set(n, pkgBefore.read(n))

  const p = PptxPresentation.open(original)
  p.addSlideTable({ index: 0, rows: [['甲', '乙']] })
  const pkgAfter = ZipPackage.open(p.save())

  const changed = [...before.keys()].filter((n) => !pkgAfter.read(n).equals(before.get(n)))
  assert.deepEqual(changed, ['ppt/slides/slide1.xml'], `意外改动：${changed.join(', ')}`)
})

test('默认套用内置表格样式（否则是无边框裸文本）', () => {
  const p = PptxPresentation.open(original)
  p.addSlideTable({ index: 0, rows: [['A']] })
  const xml = ZipPackage.open(p.save()).readText('ppt/slides/slide1.xml')
  assert.ok(xml.includes('<a:tableStyleId>'), '缺少表格样式引用')
  assert.ok(xml.includes('firstRow="1"'), '首行应标记为表头')
})

test('未指定尺寸时按行列估算并居中', () => {
  const p = PptxPresentation.open(original)
  const result = p.addSlideTable({ index: 0, rows: [['A', 'B'], ['1', '2']] })
  assert.ok(result.size_emu.cx > 0 && result.size_emu.cy > 0)
  assert.ok(result.offset_emu.x >= 0 && result.offset_emu.y >= 0)
})

test('可指定位置与尺寸', () => {
  const p = PptxPresentation.open(original)
  const result = p.addSlideTable({ index: 0, rows: [['A']], leftPx: 100, topPx: 50, widthPx: 400, heightPx: 100 })
  assert.deepEqual(result.offset_emu, { x: 100 * 9525, y: 50 * 9525 })
  assert.deepEqual(result.size_emu, { cx: 400 * 9525, cy: 100 * 9525 })
})

test('参差的行按最大列数补齐', () => {
  const p = PptxPresentation.open(original)
  const result = p.addSlideTable({ index: 0, rows: [['A', 'B', 'C'], ['1']] })
  assert.equal(result.columns, 3)
  const table = PptxPresentation.open(p.save()).readSlide(0).shapes.find((s) => s.type === 'table')
  assert.equal(table.table.column_count, 3)
  assert.equal(table.table.rows[1][2], '', '缺的单元格应为空文本')
})

test('空 rows 被拒绝', () => {
  const p = PptxPresentation.open(original)
  assert.throws(() => p.addSlideTable({ index: 0, rows: [] }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => p.addSlideTable({ index: 0, rows: [[]] }), (e) => e.code === 'INVALID_REQUEST')
})

test('单元格文本转义特殊字符', () => {
  const p = PptxPresentation.open(original)
  p.addSlideTable({ index: 0, rows: [['<标签>', 'a&b', '"引号"']] })
  const table = PptxPresentation.open(p.save()).readSlide(0).shapes.find((s) => s.type === 'table')
  assert.deepEqual(table.table.rows[0], ['<标签>', 'a&b', '"引号"'])
})

console.log('\n=== 10. 复制幻灯片与文本框 ===')

/** @param {string} part - 幻灯片部件名。 @returns {string} 其关系部件名。 */
function relsOf(part) {
  return part.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
}

test('复制幻灯片插入到源幻灯片紧后面', () => {
  const p = PptxPresentation.open(original)
  const before = p.structure().slide_count
  const result = p.duplicateSlide({ index: 0 })
  assert.equal(result.source, 0)
  assert.equal(result.index, 1)
  assert.equal(result.slide_count, before + 1)

  const reopened = PptxPresentation.open(p.save())
  assert.equal(reopened.structure().slide_count, before + 1)
  // 副本内容与源页一致
  assert.deepEqual(reopened.readSlide(1).shapes.map((s) => s.text), reopened.readSlide(0).shapes.map((s) => s.text))
})

test('副本使用全新的部件名与关系 id', () => {
  const p = PptxPresentation.open(original)
  const existing = new Set(ZipPackage.open(original).names())
  const result = p.duplicateSlide({ index: 0 })
  assert.ok(!existing.has(result.part), `部件名复用了已存在的 ${result.part}`)
  const pkg = ZipPackage.open(p.save())
  assert.ok(pkg.has(result.part), '副本部件应存在')
  const presRels = pkg.readText('ppt/_rels/presentation.xml.rels')
  assert.ok(presRels.includes(`Id="${result.relationship_id}"`), '演示文稿关系应存在')
  assert.equal((presRels.match(new RegExp(`Id="${result.relationship_id}"`, 'g')) ?? []).length, 1)
})

test('★ 复制带备注的幻灯片时丢掉备注页关系（共用备注部件会让 PowerPoint 报修复）', () => {
  const p = PptxPresentation.open(original)
  const sourceRels = p.pkg.readText(relsOf('ppt/slides/slide2.xml'))
  assert.ok(sourceRels.includes('notesSlide'), '前置条件：第 2 张幻灯片应带备注页关系')

  const result = p.duplicateSlide({ index: 1 })
  assert.equal(result.notes_dropped, true)

  const pkg = ZipPackage.open(p.save())
  assert.ok(!pkg.readText(relsOf(result.part)).includes('notesSlide'), '副本不应引用备注页部件')
  // 源页的备注关系必须原样保留
  assert.equal(pkg.readText(relsOf('ppt/slides/slide2.xml')), sourceRels)
  assert.equal(PptxPresentation.open(p.save()).validate().valid, true)
})

test('复制不带备注的幻灯片时 notes_dropped 为 false', () => {
  const p = PptxPresentation.open(original)
  assert.equal(p.duplicateSlide({ index: 0 }).notes_dropped, false)
})

test('连续复制两次各自独立', () => {
  const p = PptxPresentation.open(original)
  const a = p.duplicateSlide({ index: 0 })
  const b = p.duplicateSlide({ index: 0 })
  assert.notEqual(a.part, b.part)
  assert.notEqual(a.slide_id, b.slide_id)
  assert.equal(PptxPresentation.open(p.save()).structure().slide_count, p.structure().slide_count)
})

test('★ 复制后母版、版式、主题与备注页数量不受影响', () => {
  const before = PptxPresentation.open(original).structure()
  const p = PptxPresentation.open(original)
  p.duplicateSlide({ index: 1 })
  const after = PptxPresentation.open(p.save()).structure()
  assert.equal(after.masters, before.masters)
  assert.equal(after.layouts, before.layouts)
  assert.equal(after.themes, before.themes)
  assert.equal(after.notes_slides, before.notes_slides)
})

test('★ 复制只新增部件，原有部件逐字节不变', () => {
  const before = new Map()
  const pkgBefore = ZipPackage.open(original)
  for (const n of pkgBefore.names()) before.set(n, pkgBefore.read(n))

  const p = PptxPresentation.open(original)
  p.duplicateSlide({ index: 0 })
  const pkgAfter = ZipPackage.open(p.save())

  const changed = [...before.keys()].filter((n) => !pkgAfter.read(n).equals(before.get(n)))
  assert.deepEqual(changed.sort(), ['[Content_Types].xml', 'ppt/_rels/presentation.xml.rels', 'ppt/presentation.xml'])
  const added = pkgAfter.names().filter((n) => !before.has(n))
  assert.deepEqual(added.sort(), ['ppt/slides/_rels/slide5.xml.rels', 'ppt/slides/slide5.xml'])
})

test('复制越界时给出总数', () => {
  const p = PptxPresentation.open(original)
  assert.throws(() => p.duplicateSlide({ index: 99 }), (e) => e.code === 'FILE_NOT_FOUND')
})

test('插入文本框并读回内容', () => {
  const p = PptxPresentation.open(original)
  const result = p.addTextBox({ index: 0, text: '第一行\n第二行', leftPx: 100, topPx: 120, widthPx: 420, heightPx: 110, fontSizePt: 20, bold: true, align: 'center' })
  const slide = PptxPresentation.open(p.save()).readSlide(0)
  const box = slide.shapes.find((s) => s.name === `文本框 ${result.shape_id}`)
  assert.ok(box, '应能读到文本框形状')
  assert.ok(box.text.includes('第一行') && box.text.includes('第二行'), `文本应保留两行，实际 ${box.text}`)
  assert.equal(box.text.split('\n').length, 2)
})

test('★ 形状 id 全页唯一（重复 id 会让 PowerPoint 报修复）', () => {
  const p = PptxPresentation.open(original)
  const a = p.addTextBox({ index: 0, text: 'A' })
  const b = p.addTextBox({ index: 0, text: 'B' })
  assert.notEqual(a.shape_id, b.shape_id)

  const pkg = ZipPackage.open(p.save())
  const ids = [...pkg.readText('ppt/slides/slide1.xml').matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => m[1])
  assert.equal(new Set(ids).size, ids.length, `形状 id 有重复：${ids.join(', ')}`)
})

test('★ 文本框位置尺寸按 96 DPI 换算成 EMU', () => {
  const p = PptxPresentation.open(original)
  p.addTextBox({ index: 0, text: '定位', leftPx: 96, topPx: 48, widthPx: 480, heightPx: 240 })
  const xml = ZipPackage.open(p.save()).readText('ppt/slides/slide1.xml')
  assert.ok(xml.includes('<a:off x="914400" y="457200"/>'), '偏移应为 96*9525 / 48*9525')
  assert.ok(xml.includes('<a:ext cx="4572000" cy="2286000"/>'), '尺寸应为 480*9525 / 240*9525')
})

test('★ 字号与对齐写入 a:rPr / a:pPr', () => {
  const p = PptxPresentation.open(original)
  p.addTextBox({ index: 0, text: '样式', fontSizePt: 20, bold: true, align: 'right' })
  const xml = ZipPackage.open(p.save()).readText('ppt/slides/slide1.xml')
  assert.ok(xml.includes('sz="2000"'), '字号应以百分之一磅写入')
  assert.ok(xml.includes('algn="r"'), '右对齐应写入 algn="r"')
  assert.ok(/<a:rPr[^>]* b="1"/.test(xml), '加粗应写入 b="1"')
})

test('非法 align 被拒绝', () => {
  const p = PptxPresentation.open(original)
  assert.throws(() => p.addTextBox({ index: 0, text: 'x', align: 'justify' }), (e) => e.code === 'INVALID_REQUEST')
})

test('文本框越界时给出总数', () => {
  const p = PptxPresentation.open(original)
  assert.throws(() => p.addTextBox({ index: 99, text: 'x' }), (e) => e.code === 'FILE_NOT_FOUND')
})

test('文本框转义特殊字符', () => {
  const p = PptxPresentation.open(original)
  p.addTextBox({ index: 0, text: '<标签> & "引号"' })
  const slide = PptxPresentation.open(p.save()).readSlide(0)
  assert.ok(slide.shapes.some((s) => s.text === '<标签> & "引号"'), '特殊字符应原样读回')
})

test('★ 文本框后校验通过，且只改目标幻灯片', () => {
  const before = new Map()
  const pkgBefore = ZipPackage.open(original)
  for (const n of pkgBefore.names()) before.set(n, pkgBefore.read(n))

  const p = PptxPresentation.open(original)
  p.addTextBox({ index: 2, text: '仅这一页' })
  const bytes = p.save()
  const pkgAfter = ZipPackage.open(bytes)
  const changed = [...before.keys()].filter((n) => !pkgAfter.read(n).equals(before.get(n)))
  assert.deepEqual(changed, ['ppt/slides/slide3.xml'], `意外改动：${changed.join(', ')}`)
  assert.equal(PptxPresentation.open(bytes).validate().valid, true)
})

test('复制 + 文本框组合操作后依然校验通过', () => {
  const p = PptxPresentation.open(original)
  const dup = p.duplicateSlide({ index: 1 })
  p.addTextBox({ index: dup.index, text: '副本上的文本框\n第二行' })
  const reopened = PptxPresentation.open(p.save())
  assert.equal(reopened.validate().valid, true)
  assert.equal(reopened.readSlide(dup.index).shapes.filter((s) => s.text.includes('副本上的文本框')).length, 1)
})

console.log('\n=== 11. 版式 ===')

test('列出全部版式：下标、名称、占位符数、被谁使用', () => {
  const p = PptxPresentation.open(original)
  const layouts = p.layouts()
  assert.equal(layouts.length, p.structure().layouts)
  assert.ok(layouts.every((l) => typeof l.name === 'string' && l.name.length > 0), '每个版式都应有名称')
  assert.ok(layouts.every((l) => l.part.startsWith('ppt/slideLayouts/')), '版式部件路径不对')
  // used_by 必须与实际使用情况自洽
  const used = layouts.flatMap((l) => l.used_by.map((s) => [s, l.index]))
  assert.equal(used.length, p.structure().slide_count, '每张幻灯片应恰好对应一个版式')
})

test('读回每页当前版式', () => {
  const rows = PptxPresentation.open(original).slideLayouts()
  assert.equal(rows.length, 4)
  assert.equal(rows[0].layout_name, '标题幻灯片')
  assert.equal(rows[0].layout_index, 0)
  assert.ok(rows.every((r) => r.layout_part === null || r.layout_index !== null), '版式下标应能对上')
})

test('★ 按名称切换版式，关系目标随之改变', () => {
  const p = PptxPresentation.open(original)
  const before = p.slideLayouts()[0]
  const result = p.setSlideLayout({ index: 0, name: '空白' })
  assert.equal(result.from, before.layout_part)
  assert.notEqual(result.to, before.layout_part)
  assert.equal(result.layout_name, '空白')

  const reopened = PptxPresentation.open(p.save())
  assert.equal(reopened.slideLayouts()[0].layout_name, '空白')
  // 关系部件里应写相对路径（少一层 ../ PowerPoint 找不到版式）
  const rels = ZipPackage.open(p.save()).readText('ppt/slides/_rels/slide1.xml.rels')
  assert.ok(rels.includes(`Target="../slideLayouts/${result.to.split('/').pop()}"`), `关系目标不对：${rels}`)
})

test('按下标切换版式与按名称等价', () => {
  const byName = PptxPresentation.open(original)
  const nameResult = byName.setSlideLayout({ index: 1, name: '仅标题' })
  const byIndex = PptxPresentation.open(original)
  const indexResult = byIndex.setSlideLayout({ index: 1, layout: nameResult.to ? byName.layouts().find((l) => l.part === nameResult.to).index : 0 })
  assert.equal(indexResult.to, nameResult.to)
})

test('★ 换版式不动任何形状（占位符文本与形状数不变）', () => {
  const p = PptxPresentation.open(original)
  const before = p.readSlide(0).shapes
  p.setSlideLayout({ index: 0, name: '空白' })
  const after = PptxPresentation.open(p.save()).readSlide(0).shapes
  assert.equal(after.length, before.length)
  assert.deepEqual(after.map((s) => s.text), before.map((s) => s.text))
})

test('★ 换版式只改目标幻灯片的关系部件，其余逐字节不变', () => {
  const before = new Map()
  const pkgBefore = ZipPackage.open(original)
  for (const n of pkgBefore.names()) before.set(n, pkgBefore.read(n))

  const p = PptxPresentation.open(original)
  p.setSlideLayout({ index: 2, name: '仅标题' })
  const pkgAfter = ZipPackage.open(p.save())
  const changed = [...before.keys()].filter((n) => !pkgAfter.read(n).equals(before.get(n)))
  assert.deepEqual(changed, ['ppt/slides/_rels/slide3.xml.rels'], `意外改动：${changed.join(', ')}`)
  assert.equal(PptxPresentation.open(p.save()).validate().valid, true)
})

test('★ 版式参数校验', () => {
  const p = PptxPresentation.open(original)
  assert.throws(() => p.setSlideLayout({ index: 0 }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => p.setSlideLayout({ index: 0, layout: 99 }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => p.setSlideLayout({ index: 0, name: '不存在的版式' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => p.setSlideLayout({ index: 99, name: '空白' }), (e) => e.code === 'FILE_NOT_FOUND')
})

test('母版、主题与备注不受换版式影响', () => {
  const before = PptxPresentation.open(original).structure()
  const p = PptxPresentation.open(original)
  p.setSlideLayout({ index: 0, name: '空白' })
  const after = PptxPresentation.open(p.save()).structure()
  assert.equal(after.masters, before.masters)
  assert.equal(after.themes, before.themes)
  assert.equal(after.notes_slides, before.notes_slides)
  assert.equal(after.layouts, before.layouts, '版式数量不应变化（只是引用了另一个）')
})

console.log('\n=== 12. 原生图表 ===')

const CHART_BLOCK = {
  type: 'column',
  title: '各产品金额',
  categories: ['1月', '2月', '3月'],
  series: [{ name: '金额', values: [120, 150, 180] }]
}

test('★ 插入图表：图表部件、嵌入工作簿、关系与图形框都写齐', () => {
  const p = PptxPresentation.open(original)
  const result = p.addSlideChart({ index: 1, ...CHART_BLOCK })
  assert.equal(result.chart_part, 'ppt/charts/chart1.xml')
  assert.equal(result.embedded_workbook, 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')
  assert.equal(result.series_count, 1)

  const pkg = ZipPackage.open(p.save())
  assert.ok(pkg.has('ppt/charts/chart1.xml'), '图表部件应存在')
  assert.ok(pkg.has('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'), '嵌入工作簿应存在')
  assert.ok(pkg.readText('ppt/charts/_rels/chart1.xml.rels').includes('../embeddings/Microsoft_Excel_Worksheet1.xlsx'))
  assert.ok(pkg.readText('ppt/slides/_rels/slide2.xml.rels').includes('../charts/chart1.xml'), '幻灯片应引用图表')
  assert.ok(/<Default Extension="xlsx"/.test(pkg.readText('[Content_Types].xml')), 'xlsx 扩展名要有 Default')
  assert.ok(pkg.readText('[Content_Types].xml').includes('/ppt/charts/chart1.xml'), '图表部件要有 Override')
})

test('★ externalData 必须排在 c:spPr 之后（放错位置 PowerPoint 直接拒绝打开）', () => {
  const p = PptxPresentation.open(original)
  p.addSlideChart({ index: 1, ...CHART_BLOCK })
  const xml = ZipPackage.open(p.save()).readText('ppt/charts/chart1.xml')
  const spPr = xml.indexOf('<c:spPr>')
  const external = xml.indexOf('<c:externalData')
  const close = xml.indexOf('</c:chartSpace>')
  assert.ok(spPr > 0 && external > spPr && external < close, `顺序不对：spPr=${spPr} external=${external} close=${close}`)
})

test('★ graphicFrame 的 graphicData uri 必须是 chart 命名空间（写错图形会静默消失）', () => {
  const p = PptxPresentation.open(original)
  p.addSlideChart({ index: 1, ...CHART_BLOCK })
  const slideXml = ZipPackage.open(p.save()).readText('ppt/slides/slide2.xml')
  const frame = /<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/.exec(slideXml)[0]
  assert.ok(frame.includes('uri="http://schemas.openxmlformats.org/drawingml/2006/chart"'), frame.slice(0, 200))
  assert.ok(/<c:chart [^>]*r:id="rId\d+"\/>/.test(frame), '图形框应引用图表关系')
})

test('插入后读回：形状类型为 chart，演示文稿识别出 1 张图表', () => {
  const p = PptxPresentation.open(original)
  p.addSlideChart({ index: 1, ...CHART_BLOCK })
  const reopened = PptxPresentation.open(p.save())
  assert.equal(reopened.structure().charts, 1)
  const shapes = reopened.readSlide(1).shapes
  assert.ok(shapes.some((s) => s.type === 'chart'), `应识别出 chart 形状：${JSON.stringify(shapes.map((s) => s.type))}`)
  assert.equal(reopened.validate().valid, true)
})

test('四种图表类型都可写入，且都带上标题', () => {
  const p = PptxPresentation.open(original)
  for (const [i, type] of ['column', 'bar', 'line', 'pie'].entries()) {
    const slide = p.addSlide({ layoutOf: 2 })
    p.addSlideChart({
      index: slide.index,
      type,
      title: `${type} 标题`,
      categories: ['A', 'B'],
      series: [{ name: '系列', values: [1, 2] }]
    })
    assert.ok(slide.index >= 0, `第 ${i} 个图表应插入成功`)
  }
  const reopened = PptxPresentation.open(p.save())
  assert.equal(reopened.structure().charts, 4)
  assert.equal(reopened.validate().valid, true)
})

test('★ 多系列与中文标题正确转义写入', () => {
  const p = PptxPresentation.open(original)
  p.addSlideChart({
    index: 1,
    type: 'column',
    title: '今年 vs 去年 <对比>',
    categories: ['1月', '2月'],
    series: [
      { name: '今年', values: [10, 20] },
      { name: '去年', values: [8, 15] }
    ]
  })
  const bytes = p.save()
  const xml = ZipPackage.open(bytes).readText('ppt/charts/chart1.xml')
  assert.ok(xml.includes('今年 vs 去年 &lt;对比&gt;'), '标题应转义')
  assert.ok(!xml.includes('<对比>'), '不应出现未转义的尖括号')
  assert.equal((xml.match(/<c:ser>/g) ?? []).length, 2)
  assert.equal(PptxPresentation.open(bytes).validate().valid, true)
})

test('★ 嵌入工作簿的数据与图表一致（供 PowerPoint「编辑数据」）', () => {
  const p = PptxPresentation.open(original)
  p.addSlideChart({ index: 1, ...CHART_BLOCK })
  const bytes = p.save()
  const embedded = ZipPackage.open(bytes).read('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')
  const book = Workbook.open(embedded)
  const cells = book.readSheet(book.sheetNames()[0]).cells
  const byRef = new Map(cells.map((c) => [c.ref, c.value]))
  assert.equal(byRef.get('A2'), '1月')
  assert.equal(byRef.get('B2'), 120)
  assert.equal(byRef.get('B4'), 180)
  assert.equal(byRef.get('B1'), '金额')
  // 图表里的引用要指向这个嵌入表
  const chartXml = ZipPackage.open(bytes).readText('ppt/charts/chart1.xml')
  assert.ok(chartXml.includes('<c:f>Sheet1!$B$2:$B$4</c:f>'), '数值引用应指向嵌入表')
})

test('★ 图表参数校验', () => {
  const p = PptxPresentation.open(original)
  assert.throws(() => p.addSlideChart({ index: 1, type: 'donut', series: [{ values: [1] }] }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => p.addSlideChart({ index: 1, series: [] }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => p.addSlideChart({ index: 1, series: [{ values: [] }] }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => p.addSlideChart({ index: 1, series: [{ values: ['x'] }] }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(
    () => p.addSlideChart({ index: 1, type: 'pie', series: [{ values: [1] }, { values: [2] }] }),
    (e) => e.code === 'INVALID_REQUEST'
  )
  assert.throws(() => p.addSlideChart({ index: 99, series: [{ values: [1] }] }), (e) => e.code === 'FILE_NOT_FOUND')
})

test('★ 多个图表各自独立编号，互不覆盖', () => {
  const p = PptxPresentation.open(original)
  const a = p.addSlideChart({ index: 0, ...CHART_BLOCK })
  const b = p.addSlideChart({ index: 1, ...CHART_BLOCK })
  assert.notEqual(a.chart_part, b.chart_part)
  assert.notEqual(a.embedded_workbook, b.embedded_workbook)
  const pkg = ZipPackage.open(p.save())
  assert.ok(pkg.has(a.chart_part) && pkg.has(b.chart_part))
  assert.equal(PptxPresentation.open(p.save()).structure().charts, 2)
})

test('★ 复制含图表的幻灯片会深拷贝图表部件（共用 chartN.xml 会让 PowerPoint 拒绝打开）', () => {
  const p = PptxPresentation.open(original)
  p.addSlideChart({ index: 1, ...CHART_BLOCK })
  const before = p.pkg.names().filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n)).length
  const dup = p.duplicateSlide({ index: 1 })
  const bytes = p.save()
  const pkg = ZipPackage.open(bytes)
  const after = pkg.names().filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n)).length
  assert.equal(after, before + 1, '复制后应多出一个独立的图表部件')
  // 两页引用不同的图表部件
  const relsA = pkg.readText('ppt/slides/_rels/slide2.xml.rels')
  const relsB = pkg.readText(`ppt/slides/_rels/${dup.part.split('/').pop()}.rels`)
  const chartA = /Target="(\.\.\/charts\/chart\d+\.xml)"/.exec(relsA)[1]
  const chartB = /Target="(\.\.\/charts\/chart\d+\.xml)"/.exec(relsB)[1]
  assert.notEqual(chartA, chartB, '两页不能共用同一个图表部件')
  // 嵌入工作簿也跟着复制，两页「编辑数据」互不干扰
  const embeds = pkg.names().filter((n) => /^ppt\/embeddings\/.+\.xlsx$/.test(n))
  assert.equal(embeds.length, 2, `嵌入工作簿应各一份：${embeds.join(', ')}`)
  assert.equal(PptxPresentation.open(bytes).validate().valid, true)
})

test('★ 校验器能抓出悬空的图表引用与嵌入工作簿（不是只检查图片）', () => {
  const p = PptxPresentation.open(original)
  p.addSlideChart({ index: 1, ...CHART_BLOCK })
  const good = p.save()
  assert.equal(PptxPresentation.open(good).validate().valid, true)

  const withoutChart = ZipPackage.open(good)
  withoutChart.delete('ppt/charts/chart1.xml')
  const v1 = PptxPresentation.open(withoutChart.toBuffer()).validate()
  assert.equal(v1.valid, false, '删掉图表部件后校验必须失败')
  const c1 = v1.checks.find((c) => c.name === '图片/图表引用可解析')
  assert.ok(c1 && !c1.ok && c1.detail.includes('chart1.xml'), `应指出悬空的图表引用：${JSON.stringify(v1.checks)}`)

  const withoutEmbed = ZipPackage.open(good)
  withoutEmbed.delete('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')
  const v2 = PptxPresentation.open(withoutEmbed.toBuffer()).validate()
  assert.equal(v2.valid, false, '删掉嵌入工作簿后校验必须失败')
  const c2 = v2.checks.find((c) => c.name === '图表嵌入工作簿可解析')
  assert.ok(c2 && !c2.ok, `应指出嵌入工作簿缺失：${JSON.stringify(v2.checks)}`)
})

console.log('\n=== 13. 主题 ===')

test('★ 主题清单：名称已清掉零宽字符、标明被哪个母版引用', () => {
  const p = PptxPresentation.open(original)
  const themes = p.themes()
  assert.equal(themes.length, p.structure().themes)
  assert.ok(themes.every((t) => typeof t.name === 'string' && t.name.length > 0), `主题应有名称：${JSON.stringify(themes)}`)
  assert.ok(
    themes.every((t) => !/[\u200B-\u200D\uFEFF]/.test(t.name)),
    '名称里的零宽字符必须清掉（否则按名字匹配永远匹配不上）'
  )
  assert.ok(themes.some((t) => t.used_by_masters.includes(0)), '应有主题被母版 0 使用')
})

test('★ 主题名里的零宽字符：raw_name 保留原值、name 是清理后的值', () => {
  const themes = PptxPresentation.open(original).themes()
  const withRaw = themes.find((t) => t.raw_name && t.raw_name !== t.name)
  if (!withRaw) return // 样本本身没有零宽字符时跳过
  assert.ok(withRaw.raw_name.length > withRaw.name.length, 'raw_name 应比清理后的长')
  assert.equal(withRaw.raw_name.replace(/[\u200B-\u200D\uFEFF]/g, ''), withRaw.name)
})

test('★ 主题同名时按名字取第一个，按下标则精确（不会随机挑一个）', () => {
  const p = PptxPresentation.open(original)
  const themes = p.themes()
  const firstName = themes[0].name
  const duplicates = themes.filter((t) => t.name === firstName)
  const byName = p.setTheme({ master: 0, name: firstName })
  assert.equal(byName.source_theme, themes[0].part, '同名时应取下标最小的那个')
  if (duplicates.length > 1) {
    console.log(`     ⚠️ 样本里有 ${duplicates.length} 个同名主题「${firstName}」，按名字匹配取第一个；要精确请用 theme 下标`)
  }
  const byIndex = PptxPresentation.open(original).setTheme({ master: 0, theme: 1 })
  assert.equal(byIndex.source_theme, themes[1].part, '按下标必须精确命中第 2 个主题')
  assert.notEqual(byIndex.to, byIndex.source_theme, '挂上去的应是克隆出来的新部件，而不是源主题本身')
  assert.match(byIndex.to, /^ppt\/theme\/theme\d+\.xml$/)
})

test('★ 换主题会克隆一个新主题部件（直接指向别处在用的主题会让 PowerPoint 拒绝打开）', () => {
  const before = new Map()
  const pkgBefore = ZipPackage.open(original)
  for (const n of pkgBefore.names()) before.set(n, pkgBefore.read(n))

  const p = PptxPresentation.open(original)
  const result = p.setTheme({ master: 0, theme: 1 })
  const bytes = p.save()
  const pkgAfter = ZipPackage.open(bytes)
  assert.equal(result.reused_part, false, '原主题还被 presentation.xml / 备注母版引用，必须新建部件')
  assert.notEqual(result.to, result.source_theme, '挂上去的应是克隆出来的新部件')
  assert.equal(pkgAfter.read(result.to).equals(pkgBefore.read(result.source_theme)), true, '克隆内容应与源主题一致')

  const changed = [...before.keys()].filter((n) => !pkgAfter.read(n).equals(before.get(n)))
  assert.deepEqual(
    changed.sort(),
    ['[Content_Types].xml', 'ppt/slideMasters/_rels/slideMaster1.xml.rels'].sort(),
    `意外改动：${changed.join(', ')}`
  )
  assert.deepEqual(pkgAfter.names().filter((n) => !before.has(n)), [result.to], '只应新增那一个主题部件')
  assert.ok(pkgAfter.readText('ppt/slideMasters/_rels/slideMaster1.xml.rels').includes(`../theme/${result.to.split('/').pop()}`))
  assert.equal(PptxPresentation.open(bytes).validate().valid, true)
})

test('★ 来回切换不无限堆积主题部件（当前主题只被该母版引用时就地改内容）', () => {
  const p = PptxPresentation.open(original)
  const first = p.setTheme({ master: 0, theme: 1 })
  const back = p.setTheme({ master: 0, theme: 0 })
  assert.equal(back.reused_part, true, '切回去时应就地改克隆件的内容')
  assert.equal(back.to, first.to, '部件数不应增长')
  const again = p.setTheme({ master: 0, theme: 1 })
  assert.equal(again.to, first.to)
  assert.equal(PptxPresentation.open(p.save()).themes().length, 3, '始终只多出 1 个主题部件')
})

test('切换到自己正在用的主题是空操作', () => {
  const p = PptxPresentation.open(original)
  const current = p.masterThemes()[0].theme_part
  const index = p.themes().find((t) => t.part === current).index
  const result = p.setTheme({ master: 0, theme: index })
  assert.equal(result.changed, false)
  assert.equal(result.to, current)
})

test('读回每个母版当前的主题', () => {
  const p = PptxPresentation.open(original)
  assert.equal(p.masterThemes()[0].theme_part, 'ppt/theme/theme1.xml')
  const result = p.setTheme({ master: 0, theme: 1 })
  const reopened = PptxPresentation.open(p.save())
  const after = reopened.masterThemes()[0]
  assert.equal(after.theme_part, result.to)
  assert.equal(after.theme_index, reopened.themes().findIndex((t) => t.part === result.to))
})

test('★ 主题参数校验', () => {
  const p = PptxPresentation.open(original)
  assert.throws(() => p.setTheme({ master: 0 }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => p.setTheme({ master: 0, theme: 99 }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => p.setTheme({ master: 0, name: '不存在' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => p.setTheme({ master: 9, theme: 0 }), (e) => e.code === 'FILE_NOT_FOUND')
})

test('★ 版式清单的 type 不再恒为 null（根元素要从 #document 容器里取）', () => {
  const layouts = PptxPresentation.open(original).layouts()
  assert.ok(
    layouts.some((l) => typeof l.type === 'string' && l.type.length > 0),
    `版式 type 应有值：${JSON.stringify(layouts.map((l) => [l.name, l.type]))}`
  )
})

console.log('\n=== 14. 从零生成演示文稿 ===')

test('★ 生成必需部件齐全（少一个 PowerPoint 就报修复或拒绝打开）', () => {
  const { bytes } = buildBlankPptx({ title: '标题', subtitle: '副标题' })
  const pkg = ZipPackage.open(bytes)
  for (const part of [
    '[Content_Types].xml',
    '_rels/.rels',
    'ppt/presentation.xml',
    'ppt/_rels/presentation.xml.rels',
    'ppt/slideMasters/slideMaster1.xml',
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    'ppt/slideLayouts/slideLayout1.xml',
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    'ppt/theme/theme1.xml',
    'ppt/slides/slide1.xml',
    'ppt/slides/_rels/slide1.xml.rels',
    'docProps/core.xml',
    'docProps/app.xml'
  ]) {
    assert.ok(pkg.has(part), `缺少部件：${part}`)
  }
})

test('★ 生成的演示文稿能被自己的读取器解析为 16:9 一页', () => {
  const { bytes } = buildBlankPptx({ title: '从零生成', subtitle: '副标题' })
  const p = PptxPresentation.open(bytes)
  const structure = p.structure()
  assert.equal(structure.slide_count, 1)
  assert.equal(structure.slide_size.aspect, '16:9')
  assert.equal(structure.masters, 1)
  assert.equal(structure.layouts, 1)
  assert.equal(structure.themes, 1)
  assert.equal(p.validate().valid, true)
  const shapes = p.readSlide(0).shapes
  assert.ok(shapes.some((s) => s.text === '从零生成'), `标题占位符应有文本：${JSON.stringify(shapes.map((s) => s.text))}`)
  assert.ok(shapes.some((s) => s.text === '副标题'))
})

test('★ 母版必须带 clrMap（缺了 PowerPoint 直接报内容有问题）', () => {
  const { bytes } = buildBlankPptx()
  const master = ZipPackage.open(bytes).readText('ppt/slideMasters/slideMaster1.xml')
  assert.ok(/<p:clrMap [^>]*bg1="lt1"/.test(master), '母版缺少 p:clrMap')
  for (const key of ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent6', 'hlink', 'folHlink']) {
    assert.ok(master.includes(`${key}="`), `clrMap 缺少 ${key}`)
  }
})

test('★ 回归：主题的 fmtScheme 四组样式各 3 条，且渐变停靠点的颜色直接写在 a:gs 下', () => {
  const { bytes } = buildBlankPptx()
  const themeText = ZipPackage.open(bytes).readText('ppt/theme/theme1.xml')
  const root = XmlDoc.parse(themeText).root.children.find((c) => c.type === 'element')
  const fmtScheme = findAll(root, 'fmtScheme')[0]
  assert.ok(fmtScheme, '主题缺少 fmtScheme')
  for (const [list, count] of [
    ['fillStyleLst', 3],
    ['lnStyleLst', 3],
    ['effectStyleLst', 3],
    ['bgFillStyleLst', 3]
  ]) {
    const node = findAll(fmtScheme, list)[0]
    assert.ok(node, `主题缺少 ${list}`)
    // 直接数子元素，不用正则 —— 正则会把嵌套的 a:solidFill 也算进来（踩过）
    const children = (node.children ?? []).filter((c) => c.type === 'element')
    assert.equal(children.length, count, `${list} 应有 ${count} 条，实际 ${children.length}`)
  }
  // 这条是实测踩出来的：`<a:gs>` 里包一层 `<a:solidFill>` 不是合法的颜色选择，
  // PowerPoint 会因此**拒绝打开整个文件**（用 PowerPoint 自己的 fmtScheme 替换就正常）
  assert.ok(!/<a:gs pos="\d+"><a:solidFill>/.test(themeText), '渐变停靠点的颜色不能包在 a:solidFill 里')
  assert.ok(/<a:gs pos="\d+"><a:schemeClr val="/.test(themeText), '渐变停靠点应直接写颜色')
})

test('主题与配色是本插件自己的（不复制 Office 主题内容）', () => {
  const p = PptxPresentation.open(buildBlankPptx().bytes)
  assert.equal(p.themes()[0].name, 'office-plugin')
  const theme = ZipPackage.open(buildBlankPptx().bytes).readText('ppt/theme/theme1.xml')
  assert.ok(theme.includes('office-plugin'), '主题名应带插件标识')
  assert.ok(!theme.includes('Office Theme'), '不应出现 Office 自带主题名')
})

test('layout=blank 生成无占位符的空白首页', () => {
  const { bytes, layout } = buildBlankPptx({ layout: 'blank' })
  assert.equal(layout, 'blank')
  const p = PptxPresentation.open(bytes)
  assert.equal(p.readSlide(0).shapes.length, 0, '空白版式不应有占位符')
  assert.equal(p.validate().valid, true)
})

test('非法 layout 被拒绝', () => {
  assert.throws(() => buildBlankPptx({ layout: 'fancy' }), (e) => e.code === 'INVALID_REQUEST')
})

test('★ 生成的是「活文件」：能继续增页 / 加文本框 / 加表格 / 加图表 / 换版式', () => {
  const created = PptxPresentation.open(buildBlankPptx({ title: '母版' }).bytes)
  const slide = created.addSlide({ layoutOf: 0 })
  created.addTextBox({ index: slide.index, text: '后加的文本框' })
  created.addSlideTable({ index: slide.index, rows: [['A', 'B']] })
  created.addSlideChart({
    index: slide.index,
    type: 'pie',
    title: '后加的图表',
    categories: ['一', '二'],
    series: [{ name: '占比', values: [60, 40] }]
  })
  created.setSlideLayout({ index: slide.index, layout: 0 })

  const bytes = created.save()
  const reopened = PptxPresentation.open(bytes)
  assert.equal(reopened.structure().slide_count, 2)
  assert.equal(reopened.structure().charts, 1)
  assert.equal(reopened.validate().valid, true)
  assert.ok(reopened.readSlide(1).shapes.some((s) => s.text === '后加的文本框'))
})

console.log(`\n结果：${passed} 通过，${failed} 失败\n`)
process.exit(failed === 0 ? 0 : 1)
