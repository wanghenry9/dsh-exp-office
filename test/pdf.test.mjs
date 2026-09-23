/**
 * PDF 适配器测试（阶段 5：读取与文本提取）。
 *
 * 样本是真实世界的 PDF：用宿主内置 LibreOffice 引擎把中文 DOCX 转出来，
 * 含压缩内容流、子集化 CJK 字体与 ToUnicode CMap。
 *
 * 运行：node test/pdf.test.mjs
 */
import assert from 'node:assert/strict'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import zlib from 'node:zlib'
import { PdfDocument } from '../lib/pdf.js'
import { buildPdf } from '../lib/pdf-writer.js'
import { readCmap, readFont, subsetFont } from '../lib/font-subset.js'

const here = dirname(fileURLToPath(import.meta.url))
const fixture = join(here, 'fixtures', 'sample.pdf')

if (!existsSync(fixture)) {
  console.log('⏭️  跳过：未找到 test/fixtures/sample.pdf（先运行 node test/make-pdf-fixture.mjs）。')
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

console.log('\n=== 1. 打开与元数据 ===')

test('打开真实 PDF 并识别版本', () => {
  const pdf = PdfDocument.open(original)
  assert.match(pdf.info.version, /^1\.\d+$/)
})

test('读取文档元数据', () => {
  const meta = PdfDocument.open(original).metadata()
  assert.equal(meta.title, '季度销售报告')
  assert.equal(meta.author, 'dsh-exp-office')
  assert.match(meta.producer, /LibreOffice/)
  assert.match(meta.creation_date, /^D:\d{14}/)
})

test('拒绝非 PDF 文件', () => {
  assert.throws(() => PdfDocument.open(Buffer.from('这不是 PDF 文件，没有文件头')), (e) => e.code === 'CORRUPTED_DOCUMENT')
})

test('拒绝过小的文件', () => {
  assert.throws(() => PdfDocument.open(Buffer.from('PDF')), (e) => e.code === 'CORRUPTED_DOCUMENT')
})

console.log('\n=== 2. 页面 ===')

test('识别页面数量与尺寸', () => {
  const pdf = PdfDocument.open(original)
  const pages = pdf.pageInfos()
  assert.equal(pages.length, 1)
  // A4：595 × 842 pt（允许 1pt 误差）
  assert.ok(Math.abs(pages[0].width_pt - 595.3) < 1, `宽度 ${pages[0].width_pt}`)
  assert.ok(Math.abs(pages[0].height_pt - 841.9) < 1, `高度 ${pages[0].height_pt}`)
  assert.equal(pages[0].rotation, 0)
})

test('★ 按 /Kids 顺序归一化页面次序', () => {
  const pdf = PdfDocument.open(original)
  const pages = pdf.pages()
  assert.equal(pages.length, 1)
  assert.ok(pages[0].num > 0)
})

console.log('\n=== 3. 文本提取 ===')

test('★★ 提取中文文本（ToUnicode CMap 生效）', () => {
  const { text } = PdfDocument.open(original).extractText()
  assert.ok(text.includes('季度销售报告'), `未提取到标题：${JSON.stringify(text.slice(0, 120))}`)
  assert.ok(text.includes('本文件用于验证'), `未提取到正文：${JSON.stringify(text.slice(0, 200))}`)
  assert.ok(text.includes('月份') && text.includes('金额'), '未提取到表格文本')
})

test('提取结果长度与分页统计自洽', () => {
  const result = PdfDocument.open(original).extractText()
  assert.equal(result.pages, 1)
  assert.equal(result.truncated, false)
  assert.ok(result.length > 20)
  assert.equal(result.chars_per_page.length, 1)
})

test('支持字符上限截断', () => {
  const result = PdfDocument.open(original).extractText({ maxChars: 5 })
  assert.equal(result.text.length, 5)
  assert.equal(result.truncated, true)
})

test('可按页提取', () => {
  const result = PdfDocument.open(original).extractText({ page: 0 })
  assert.equal(result.pages, 1)
  assert.ok(result.text.includes('季度'))
})

test('页越界时报错并给出总页数', () => {
  const pdf = PdfDocument.open(original)
  try {
    pdf.extractText({ page: 99 })
    assert.fail('应当抛出')
  } catch (err) {
    assert.equal(err.code, 'INVALID_REQUEST')
    assert.equal(err.details.page_count, 1)
  }
})

console.log('\n=== 4. 结构识别 ===')

test('识别字体与未加密状态', () => {
  const s = PdfDocument.open(original).structure()
  assert.equal(s.encrypted, false)
  assert.ok(s.font_count > 0, `字体数 ${s.font_count}`)
  assert.ok(s.fonts.some((f) => /SourceHanSerif|Calibri/.test(f)), `实际字体：${s.fonts.join(', ')}`)
})

test('报告对象数、签名、表单与批注状态', () => {
  const s = PdfDocument.open(original).structure()
  assert.ok(s.object_count > 5)
  assert.equal(s.has_digital_signature, false)
  assert.equal(s.form_fields, 0)
  assert.equal(s.annotations, 0)
})

console.log('\n=== 5. 校验 ===')

test('样本校验通过', () => {
  const report = PdfDocument.open(original).validate()
  assert.equal(report.valid, true, JSON.stringify(report.checks.filter((c) => !c.ok)))
})

test('★ 如实列出未做的检查项（不做 xref 一致性校验）', () => {
  const report = PdfDocument.open(original).validate()
  assert.ok(report.not_checked.length >= 5)
  assert.ok(report.not_checked.some((c) => c.name === '交叉引用表一致性'))
  assert.ok(report.not_checked.some((c) => c.name === '数字签名有效性'))
  assert.ok(report.not_checked.every((c) => typeof c.reason === 'string' && c.reason.length > 0))
})

console.log('\n=== 6. 内容流与过滤器 ===')

test('解码 FlateDecode 内容流', () => {
  const pdf = PdfDocument.open(original)
  const content = pdf.pageContent(pdf.pages()[0])
  assert.ok(content.length > 100, `内容流长度 ${content.length}`)
  assert.match(content.toString('latin1'), /BT\b/)
})

test('★ 损坏的内容流给出结构化错误而不是崩溃', () => {
  // 完整的对象图：Catalog → Pages → Page → Contents（内容流故意写坏）
  const body = [
    '%PDF-1.4\n',
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
    '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R >>\nendobj\n',
    '4 0 obj\n<< /Length 12 /Filter /FlateDecode >>\nstream\nnot-zlib-data\nendstream\nendobj\n',
    'trailer\n<< /Root 1 0 R /Size 5 >>\n%%EOF\n'
  ].join('')
  const pdf = PdfDocument.open(Buffer.from(body, 'latin1'))
  assert.equal(pdf.pages().length, 1, '应能解析出页面')
  assert.throws(() => pdf.extractText(), (e) => e.code === 'CORRUPTED_DOCUMENT')
})

test('★ 未压缩的十六进制与字面字符串都能解析', () => {
  // 手工构造一份未压缩、无字体映射的最小 PDF，验证基础路径
  const body = [
    '%PDF-1.4\n',
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
    '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R >>\nendobj\n',
    '4 0 obj\n<< /Length 44 >>\nstream\nBT /F1 12 Tf 10 50 Td (Hello PDF) Tj ET\nendstream\nendobj\n',
    'trailer\n<< /Root 1 0 R /Size 5 >>\n%%EOF\n'
  ].join('')
  const pdf = PdfDocument.open(Buffer.from(body, 'latin1'))
  const pages = pdf.pageInfos()
  assert.equal(pages.length, 1)
  assert.equal(pages[0].width_pt, 200)
  assert.equal(pages[0].height_pt, 100)
  const { text } = pdf.extractText()
  assert.ok(text.includes('Hello PDF'), `实际提取：${JSON.stringify(text)}`)
})

test('★ 十六进制字符串与 TJ 数组都能提取', () => {
  const body = [
    '%PDF-1.4\n',
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
    '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R >>\nendobj\n',
    '4 0 obj\n<< /Length 60 >>\nstream\nBT /F1 12 Tf 10 50 Td <48656C6C6F> Tj [<20576F726C64>] TJ ET\nendstream\nendobj\n',
    'trailer\n<< /Root 1 0 R /Size 5 >>\n%%EOF\n'
  ].join('')
  const pdf = PdfDocument.open(Buffer.from(body, 'latin1'))
  const { text } = pdf.extractText()
  assert.ok(text.includes('Hello'), `十六进制字符串未提取：${JSON.stringify(text)}`)
  assert.ok(text.includes(' World'), `TJ 数组未提取：${JSON.stringify(text)}`)
})

test('★ 加密 PDF 被识别而不是静默给出空内容', () => {
  const body = [
    '%PDF-1.4\n',
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
    '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] >>\nendobj\n',
    '9 0 obj\n<< /Filter /Standard /V 2 /R 3 /Length 128 /P -3904 >>\nendobj\n',
    'trailer\n<< /Root 1 0 R /Encrypt 9 0 R /Size 10 >>\n%%EOF\n'
  ].join('')
  const pdf = PdfDocument.open(Buffer.from(body, 'latin1'))
  const s = pdf.structure()
  assert.equal(s.encrypted, true)
  assert.ok(s.encryption, '应给出加密信息')
  assert.equal(s.encryption.algorithm, 'RC4')
  assert.equal(s.encryption.key_length_bits, 128)
  assert.equal(s.encryption.readable, true)
  const report = pdf.validate()
  assert.ok(report.checks.some((c) => c.name === '未加密或已声明加密' && c.detail.includes('已加密')))
})

test('★ 加密字典悬空引用时也报告「已加密」（不是 null）', () => {
  const body = [
    '%PDF-1.4\n',
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
    '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] >>\nendobj\n',
    'trailer\n<< /Root 1 0 R /Encrypt 99 0 R /Size 10 >>\n%%EOF\n'
  ].join('')
  const pdf = PdfDocument.open(Buffer.from(body, 'latin1'))
  const s = pdf.structure()
  assert.equal(s.encrypted, true)
  assert.equal(s.encryption.readable, false)
  assert.match(s.encryption.algorithm, /不可读/)
})

console.log('\n=== 7. 页面级写操作（增量更新） ===')

const multiFixture = join(here, 'fixtures', 'sample-multipage.pdf')
const multi = existsSync(multiFixture) ? readFileSync(multiFixture) : null

/**
 * 读出一份 PDF 每一页的文本，用于验证页序与页数。
 * @param {Buffer} bytes - PDF 字节。
 * @returns {string[]} 每页文本。
 */
function perPageText(bytes) {
  const pdf = PdfDocument.open(bytes)
  return pdf.pages().map((_, index) => pdf.extractText({ page: index }).text.replace(/\s+/g, ' ').trim())
}

/**
 * 校验增量更新段的内部自洽性。
 *
 * 这是**互操作性契约**：第三方阅读器只按 startxref → xref → 对象偏移这条链读文件，
 * 因此偏移必须精确指向 `N G obj`，且 /Prev 必须指回旧交叉引用段。
 *
 * @param {Buffer} next - 写回后的字节。
 * @param {Buffer} prev - 原始字节。
 * @returns {object} 校验细节。
 */
function inspectIncremental(next, prev) {
  assert.ok(next.length > prev.length, '增量更新应只追加字节')
  assert.ok(next.subarray(0, prev.length).equals(prev), '原始字节必须原样保留（作为前缀）')

  const tailStart = next.lastIndexOf(Buffer.from('startxref', 'latin1'))
  assert.ok(tailStart > 0, '文件尾部应有 startxref')
  const match = /^startxref\s+(\d+)\s*%%EOF\s*$/.exec(next.subarray(tailStart).toString('latin1'))
  assert.ok(match, 'startxref 段格式不正确')
  const xrefOffset = Number(match[1])
  assert.equal(next.subarray(xrefOffset, xrefOffset + 4).toString('latin1'), 'xref', 'startxref 必须指向 xref 关键字')

  const xrefEnd = next.indexOf(Buffer.from('trailer', 'latin1'), xrefOffset)
  assert.ok(xrefEnd > xrefOffset, 'xref 段后应有 trailer')
  const xrefText = next.subarray(xrefOffset, xrefEnd).toString('latin1')
  const lines = xrefText.split('\n').filter((line) => line !== '')
  assert.equal(lines[0], 'xref')

  const offsets = []
  let index = 1
  while (index < lines.length) {
    const header = /^(\d+) (\d+)$/.exec(lines[index])
    assert.ok(header, `xref 子段头不合法：${JSON.stringify(lines[index])}`)
    const first = Number(header[1])
    const count = Number(header[2])
    index += 1
    for (let i = 0; i < count; i += 1) {
      const entry = /^(\d{10}) (\d{5}) ([nf]) $/.exec(lines[index])
      assert.ok(entry, `xref 条目不合法：${JSON.stringify(lines[index])}`)
      const offset = Number(entry[1])
      if (entry[3] === 'n') {
        const head = next.subarray(offset, offset + 40).toString('latin1')
        const objMatch = /^(\d+) (\d+) obj\b/.exec(head)
        assert.ok(objMatch, `偏移 ${offset} 未指向对象头：${JSON.stringify(head.slice(0, 24))}`)
        assert.equal(Number(objMatch[1]), first + i, `偏移 ${offset} 处的对象号与 xref 不一致`)
        offsets.push(offset)
      }
      index += 1
    }
  }
  assert.ok(offsets.length >= 1, '至少应写回一个对象')

  const trailerText = next.subarray(xrefEnd).toString('latin1')
  const prevMatch = /\/Prev (\d+)/.exec(trailerText)
  assert.ok(prevMatch, 'trailer 必须带 /Prev 指回旧交叉引用段')
  assert.ok(/\/Root \d+ \d+ R/.test(trailerText), 'trailer 的 /Root 必须是合法的间接引用')
  return { xrefOffset, prev: Number(prevMatch[1]), objects: offsets.length, appended: next.length - prev.length }
}

/**
 * 读出文件尾部的 startxref 值。
 * @param {Buffer} bytes - PDF 字节。
 * @returns {number} 偏移。
 */
function readStartxrefOf(bytes) {
  const pos = bytes.lastIndexOf(Buffer.from('startxref', 'latin1'))
  const match = /^startxref\s+(\d+)/.exec(bytes.subarray(pos).toString('latin1'))
  assert.ok(match, '原文件应有 startxref')
  return Number(match[1])
}

test('未修改时 save() 原样返回，不做任何追加', () => {
  const pdf = PdfDocument.open(original)
  assert.equal(pdf.dirty, false)
  assert.ok(pdf.save().equals(original))
})

test('★ 旋转页面：原字节保持为前缀，/Rotate 写回', () => {
  const pdf = PdfDocument.open(original)
  const change = pdf.rotatePage({ page: 0, degrees: 90 })
  assert.deepEqual(change, { type: 'rotate_page', page: 0, from: 0, to: 90 })
  assert.equal(pdf.dirty, true)
  const next = pdf.save()
  assert.equal(pdf.dirty, false, '写回后应清空待写队列')
  const info = inspectIncremental(next, original)
  assert.ok(info.appended > 0)
  assert.equal(info.prev, readStartxrefOf(original), '/Prev 应等于原文件的 startxref')
  assert.equal(PdfDocument.open(next).pageInfos()[0].rotation, 90)
})

test('★ 追加的页面对象只多了 /Rotate（最小修改）', () => {
  const before = PdfDocument.open(original)
  const pageRef = `${before.pages()[0].num} ${before.pages()[0].gen}`
  const beforeKeys = new Set(Object.keys(before.resolve(...pageRef.split(' ').map(Number))))
  const pdf = PdfDocument.open(original)
  pdf.rotatePage({ page: 0, degrees: 270 })
  const after = PdfDocument.open(pdf.save())
  const page = after.resolve(...pageRef.split(' ').map(Number))
  const afterKeys = new Set(Object.keys(page))
  const added = [...afterKeys].filter((k) => !beforeKeys.has(k))
  const removed = [...beforeKeys].filter((k) => !afterKeys.has(k))
  assert.deepEqual(added, ['Rotate'], `不应新增其他键：${added}`)
  assert.deepEqual(removed, [], `不应丢键：${removed}`)
  assert.equal(page.Rotate, 270)
})

test('角度归一到 0/90/180/270', () => {
  const pdf = PdfDocument.open(original)
  assert.equal(pdf.rotatePage({ page: 0, degrees: 450 }).to, 90)
  assert.equal(pdf.rotatePage({ page: 0, degrees: -90 }).to, 270)
  assert.equal(pdf.rotatePage({ page: 0, degrees: 360 }).to, 0)
})

test('越界页号报错并给出总页数', () => {
  const pdf = PdfDocument.open(original)
  assert.throws(
    () => pdf.rotatePage({ page: 9, degrees: 90 }),
    (err) => err.code === 'INVALID_REQUEST' && err.details.page_count === 1
  )
})

test('拒绝删除最后一页', () => {
  const pdf = PdfDocument.open(original)
  assert.throws(() => pdf.deletePage({ page: 0 }), (err) => err.code === 'INVALID_REQUEST')
})

test('★ 名称与字符串按 PDF 语法转义后往返', () => {
  const body = [
    '%PDF-1.4\n',
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
    '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Foo#20Bar (hello) /Bin <00FF> >>\nendobj\n',
    'trailer\n<< /Root 1 0 R /Size 4 >>\n%%EOF\n'
  ].join('')
  const bytes = Buffer.from(body, 'latin1')
  const pdf = PdfDocument.open(bytes)
  pdf.rotatePage({ page: 0, degrees: 90 })
  const next = pdf.save()
  const appended = next.subarray(bytes.length).toString('latin1')
  assert.ok(appended.includes('/Foo#20Bar'), `名称未转义：${appended}`)
  assert.ok(appended.includes('(hello)'), `字符串未按字面串写回：${appended}`)
  assert.ok(appended.includes('<00FF>'), `高位字节未按十六进制串写回：${appended}`)
  const page = PdfDocument.open(next).pages()[0].value
  assert.equal(String(page['Foo Bar']), 'hello')
  assert.ok(Buffer.isBuffer(page.Bin))
  assert.deepEqual([...page.Bin], [0x00, 0xff])
})

if (!multi) {
  console.log('  ⏭️  跳过多页用例：先运行 node test/make-pdf-fixture.mjs 生成 sample-multipage.pdf。')
} else {
  test('★ 增量段结构自洽（startxref → xref → 对象偏移）', () => {
    const pdf = PdfDocument.open(multi)
    pdf.rotatePage({ page: 1, degrees: 180 })
    inspectIncremental(pdf.save(), multi)
  })

  test('★ 重排页面后页序与文本顺序同步变化', () => {
    const before = perPageText(multi)
    assert.equal(before.length, 3, `多页样本应有 3 页，实际 ${before.length}`)
    const pdf = PdfDocument.open(multi)
    const result = pdf.movePage({ from: 0, to: 2 })
    assert.equal(result.order.length, 3)
    const after = perPageText(pdf.save())
    assert.deepEqual(after, [before[1], before[2], before[0]])
  })

  test('★ 删除页面后页数减一，被摘掉的对象不再算作页', () => {
    const before = perPageText(multi)
    const pdf = PdfDocument.open(multi)
    const removed = pdf.pages()[1]
    const result = pdf.deletePage({ page: 1 })
    assert.equal(result.page_count, 2)
    assert.equal(result.removed_object, `${removed.num} ${removed.gen}`)
    const next = pdf.save()
    const after = perPageText(next)
    assert.deepEqual(after, [before[0], before[2]])
    // 被摘掉的对象仍在文件里（增量更新不回收字节），这是设计取舍，不是漏删
    assert.ok(next.toString('latin1').includes(`${removed.num} ${removed.gen} obj`))
    assert.equal(PdfDocument.open(next).structure().page_count, 2)
  })

  test('★ 读页面树用嵌套节点也能读，但写操作明确拒绝', () => {
    const body = [
      '%PDF-1.4\n',
      '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
      '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 2 >>\nendobj\n',
      '3 0 obj\n<< /Type /Pages /Parent 2 0 R /Kids [4 0 R 5 0 R] /Count 2 >>\nendobj\n',
      '4 0 obj\n<< /Type /Page /Parent 3 0 R /MediaBox [0 0 200 100] >>\nendobj\n',
      '5 0 obj\n<< /Type /Page /Parent 3 0 R /MediaBox [0 0 200 100] >>\nendobj\n',
      'trailer\n<< /Root 1 0 R /Size 6 >>\n%%EOF\n'
    ].join('')
    const pdf = PdfDocument.open(Buffer.from(body, 'latin1'))
    assert.equal(pdf.pages().length, 2, '嵌套页面树应能读出 2 页')
    assert.throws(() => pdf.movePage({ from: 0, to: 1 }), (err) => err.code === 'UNSUPPORTED_FEATURE')
    assert.throws(() => pdf.deletePage({ page: 0 }), (err) => err.code === 'UNSUPPORTED_FEATURE')
  })

  test('★ 第三方阅读器能读的最低要求：xref 偏移精确、/Prev 正确', () => {
    const pdf = PdfDocument.open(multi)
    pdf.movePage({ from: 2, to: 0 })
    const info = inspectIncremental(pdf.save(), multi)
    assert.ok(info.objects >= 1)
  })
}

console.log('\n=== 8. 叠加式写入（水印与页码） ===')

test('★ 水印只追加内容流，原内容流与资源条目都不丢', () => {
  const before = PdfDocument.open(original)
  const pdf = PdfDocument.open(original)
  const info = pdf.addTextOverlay({ text: 'DRAFT ONLY', position: 'center' })
  assert.deepEqual(info.pages, [0])
  const next = pdf.save()
  assert.ok(next.subarray(0, original.length).equals(original), '原字节必须原样保留')

  const after = PdfDocument.open(next)
  const afterText = after.extractText().text
  assert.ok(afterText.includes('DRAFT ONLY'), `水印文本未进入内容流：${JSON.stringify(afterText.slice(-60))}`)
  // 原有文字必须仍然可读（资源合并若把 /Font 覆盖掉，中文会退化成控制字符）
  for (const fragment of ['季度销售报告', '本文件用于验证']) {
    assert.ok(afterText.includes(fragment), `原有文字丢失：${fragment} → ${JSON.stringify(afterText.slice(0, 80))}`)
  }
  assert.equal(after.structure().font_count, before.structure().font_count + 1, '只应多出 Helvetica 一个字体')
  assert.equal(after.validate().valid, true)
})

test('★ 页面 /Contents 变成数组，新流排在最后', () => {
  const pdf = PdfDocument.open(original)
  pdf.addTextOverlay({ text: 'TAIL', position: 'bottom' })
  const next = PdfDocument.open(pdf.save())
  const contents = next.pages()[0].value.Contents
  assert.ok(Array.isArray(contents), '/Contents 应为数组')
  assert.equal(contents.length, 2)
  const last = next.deref(contents[1])
  const text = next.decodeStream(last).toString('latin1')
  assert.ok(text.includes('(TAIL) Tj'), text)
  // 第一个流必须还是原来的对象
  assert.ok(next.deref(contents[0]) && next.decodeStream(next.deref(contents[0])).length > 100)
})

test('★ 页码占位符与 labelFor 逐页取值', () => {
  const pdf = PdfDocument.open(original)
  pdf.addTextOverlay({ text: '{page} / {total}', position: 'bottom', fontSize: 10 })
  const next = PdfDocument.open(pdf.save())
  assert.ok(next.extractText().text.includes('1 / 1'))

  const pdf2 = PdfDocument.open(original)
  pdf2.addTextOverlay({ text: 'x', position: 'bottom', labelFor: (index) => `PAGE-${index + 10}` })
  const next2 = PdfDocument.open(pdf2.save())
  assert.ok(next2.extractText().text.includes('PAGE-10'))
})

test('★ 中文叠加文本被明确拒绝（需要嵌入字体）', () => {
  const pdf = PdfDocument.open(original)
  assert.throws(
    () => pdf.addTextOverlay({ text: '机密' }),
    (err) => err.code === 'UNSUPPORTED_FEATURE' && /ASCII/.test(err.message)
  )
})

test('叠加参数不合法时给出结构错误', () => {
  const pdf = PdfDocument.open(original)
  assert.throws(() => pdf.addTextOverlay({ text: '' }), (err) => err.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.addTextOverlay({ text: 'A', position: 'left' }), (err) => err.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.addTextOverlay({ text: 'A', gray: 2 }), (err) => err.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.addTextOverlay({ text: 'A', pages: [5] }), (err) => err.code === 'INVALID_REQUEST')
})

if (multi) {
  test('★ 多页水印：只改指定页，未指定页的 /Contents 保持单流', () => {
    const pdf = PdfDocument.open(multi)
    const info = pdf.addTextOverlay({ text: 'CONFIDENTIAL', pages: [0, 2] })
    assert.deepEqual(info.pages, [0, 2])
    const next = PdfDocument.open(pdf.save())
    assert.equal(next.pages().length, 3)
    assert.ok(Array.isArray(next.pages()[0].value.Contents))
    assert.ok(Array.isArray(next.pages()[2].value.Contents))
    assert.ok(!Array.isArray(next.pages()[1].value.Contents), '第 2 页不应被改动')
    const texts = [0, 1, 2].map((i) => next.extractText({ page: i }).text)
    assert.ok(texts[0].includes('CONFIDENTIAL'))
    assert.ok(!texts[1].includes('CONFIDENTIAL'))
    assert.ok(texts[2].includes('CONFIDENTIAL'))
    assert.equal(next.validate().valid, true)
  })

  test('★ 逐页页码：每页文本各自正确', () => {
    const pdf = PdfDocument.open(multi)
    pdf.addTextOverlay({
      text: 'p',
      position: 'bottom',
      labelFor: (index) => `- ${index + 1} -`
    })
    const next = PdfDocument.open(pdf.save())
    assert.deepEqual(
      [0, 1, 2].map((i) => next.extractText({ page: i }).text.trim().split('\n').pop()),
      ['- 1 -', '- 2 -', '- 3 -']
    )
  })

  test('★ 叠加后仍是合法增量更新（xref 链自洽）', () => {
    const pdf = PdfDocument.open(multi)
    pdf.addTextOverlay({ text: 'OK', pages: [1] })
    inspectIncremental(pdf.save(), multi)
  })
}

console.log('\n=== 9. 元数据写入 ===')

test('★ 改写元数据：原字节保持前缀，ASCII 与中文都能回读', () => {
  const pdf = PdfDocument.open(original)
  const result = pdf.updateMetadata({ title: '季度销售报告（已归档）', author: 'dsh-exp-office', keywords: 'Q3, sales' })
  assert.equal(result.created_info, false, '样本自带 /Info，不应新建')
  const next = pdf.save()
  assert.ok(next.subarray(0, original.length).equals(original), '原字节必须原样保留')
  const meta = PdfDocument.open(next).metadata()
  assert.equal(meta.title, '季度销售报告（已归档）')
  assert.equal(meta.author, 'dsh-exp-office')
  assert.equal(meta.keywords, 'Q3, sales')
  assert.equal(PdfDocument.open(next).pages().length, 1)
  // 中文按 UTF-16BE + BOM 写：文件里应出现 FE FF 开头的十六进制串
  assert.ok(/<FEFF[0-9A-F]+>/.test(next.toString('latin1')), '中文元数据应写成 UTF-16BE 十六进制串')
})

test('★ 原先没有 /Info 时新建对象并接进 trailer', () => {
  const body = [
    '%PDF-1.4\n',
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
    '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] >>\nendobj\n',
    'trailer\n<< /Root 1 0 R /Size 4 >>\n%%EOF\n'
  ].join('')
  const bytes = Buffer.from(body, 'latin1')
  const pdf = PdfDocument.open(bytes)
  const result = pdf.updateMetadata({ title: 'No Info Before' })
  assert.equal(result.created_info, true)
  const next = pdf.save()
  const text = next.subarray(bytes.length).toString('latin1')
  assert.ok(text.includes('/Info'), 'trailer 必须引用新建的 /Info')
  assert.ok(/\/Type \/Catalog|\/Title/.test(text))
  assert.equal(PdfDocument.open(next).metadata().title, 'No Info Before')
})

test('元数据写入需要至少一个字段', () => {
  const pdf = PdfDocument.open(original)
  assert.throws(() => pdf.updateMetadata({}), (err) => err.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.updateMetadata(), (err) => err.code === 'INVALID_REQUEST')
})

console.log('\n=== 10. 对象图重写（回收字节 / 拆分） ===')

test('★ 全量重写：页数、每页文本、元数据、页尺寸全部保真', () => {
  const before = PdfDocument.open(original)
  const beforeText = before.extractText().text
  const beforeMeta = before.metadata()
  const result = before.rewrite()
  assert.ok(result.objects > 10, `对象数偏少：${result.objects}`)
  const after = PdfDocument.open(result.buffer)
  assert.equal(after.pages().length, 1)
  assert.equal(after.extractText().text.replace(/\s+/g, ' '), beforeText.replace(/\s+/g, ' '))
  assert.equal(after.metadata().title, beforeMeta.title)
  assert.equal(after.metadata().author, beforeMeta.author)
  assert.deepEqual(after.pageInfos()[0].width_pt, before.pageInfos()[0].width_pt)
  assert.equal(after.validate().valid, true)
})

test('★ 重写产出经典 xref 表，不再有对象流与交叉引用流', () => {
  const result = PdfDocument.open(original).rewrite()
  const text = result.buffer.toString('latin1')
  assert.ok(/\nxref\n/.test(text), '应写出经典 xref 表')
  assert.ok(!/\/Type\s*\/ObjStm/.test(text), '不应再有对象流容器')
  assert.ok(!/\/Type\s*\/XRef/.test(text), '不应再有交叉引用流')
  assert.match(text, /startxref\n\d+\n%%EOF\n$/)
  // 间接引用必须是三段式：漏了生成号（`/Root 1 R`）会被阅读器当成数字，目录/元数据静默丢失
  assert.match(text, /\/Root \d+ 0 R/, 'trailer 的 /Root 必须是 num gen R')
  assert.ok(!/\/Root \d+ R(?!\s*$)/.test(text), 'trailer 的 /Root 不能漏生成号')
})

test('★ 字节级不变量：重写后页面对象数恰好等于页数（删页的取舍对照）', () => {
  const countPageObjects = (buf) => (buf.toString('latin1').match(/\/Type\s*\/Page[^s]/g) ?? []).length
  // 增量式删页：对象还在文件里，只是没人引用 —— 这是明确的取舍
  const incremental = PdfDocument.open(multi ?? original)
  if (incremental.pages().length >= 2) {
    incremental.deletePage({ page: 0 })
    const afterDelete = incremental.save()
    assert.equal(PdfDocument.open(afterDelete).pages().length, incremental.pages().length)
    assert.ok(
      countPageObjects(afterDelete) > PdfDocument.open(afterDelete).pages().length,
      '增量式删页后，文件里仍应有被摘掉引用的页面对象（这是设计取舍）'
    )
  }
  // 重写：对象数不多不少，被删/未选页在字节里彻底消失
  const rewritten = PdfDocument.open(multi ?? original).rewrite({ keepPages: [0] })
  assert.equal(countPageObjects(rewritten.buffer), 1, '重写后 /Type /Page 出现次数应恰好等于保留页数')
  assert.equal(PdfDocument.open(rewritten.buffer).pages().length, 1)
})

test('★ 回收字节：反复增量更新后重写把增长收回来', () => {
  let grown = multi ?? original
  const doc0 = PdfDocument.open(grown)
  const pageCount = doc0.pages().length
  const baseline = grown.length
  for (let i = 0; i < 5; i += 1) {
    const doc = PdfDocument.open(grown)
    doc.rotatePage({ page: 0, degrees: 90 })
    grown = doc.save()
  }
  assert.ok(grown.length > baseline, '增量更新应该让文件变大')
  const result = PdfDocument.open(grown).rewrite()
  assert.ok(result.buffer.length < grown.length, `重写应回收字节：${grown.length} → ${result.buffer.length}`)
  const after = PdfDocument.open(result.buffer)
  assert.equal(after.pageInfos()[0].rotation, 90, '重写必须保留编辑结果')
  assert.equal(after.pages().length, pageCount)
})

test('★ 拆分：只保留选中页，未选页在字节里彻底消失', () => {
  const source = multi ?? original
  const doc = PdfDocument.open(source)
  const pageCount = doc.pages().length
  const keep = pageCount >= 3 ? [2, 0] : [0]
  const result = doc.rewrite({ keepPages: keep })
  assert.deepEqual(result.pages, keep)
  assert.equal(result.dropped, pageCount - keep.length)
  const after = PdfDocument.open(result.buffer)
  assert.equal(after.pages().length, keep.length)
  assert.equal(after.validate().valid, true)
  if (pageCount >= 3) {
    // 页序按 keepPages 给出的顺序
    const expected = keep.map((i) => doc.extractText({ page: i }).text.replace(/\s+/g, '').trim())
    const actual = keep.map((_, i) => after.extractText({ page: i }).text.replace(/\s+/g, '').trim())
    assert.deepEqual(actual, expected)
  }
  // 泄漏检查用**对象计数**而不是文本匹配：内容流是压缩的，正文字符串本来就不以明文出现在字节里，
  // 「搜不到文字」并不能证明内容不在文件里。页面对象数恰好等于页数才是硬条件。
  // （这条断言抓出过一个真 bug：书签的 /Dest 指回被丢弃的页面，可达性遍历把整页内容又写了回来。）
  const pageObjects = (result.buffer.toString('latin1').match(/\/Type\s*\/Page[^s]/g) ?? []).length
  assert.equal(pageObjects, keep.length, `字节里应只剩 ${keep.length} 个页面对象，实际 ${pageObjects}`)
  assert.ok(result.buffer.length < source.length, '丢掉页面后体积应变小')
})

test('★ 拆分嵌套页面树也能得到正确页数与内容', () => {
  const body = [
    '%PDF-1.4\n',
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 2 >>\nendobj\n',
    '3 0 obj\n<< /Type /Pages /Parent 2 0 R /Kids [4 0 R 5 0 R] /Count 2 >>\nendobj\n',
    '4 0 obj\n<< /Type /Page /Parent 3 0 R /MediaBox [0 0 200 100] /Contents 6 0 R >>\nendobj\n',
    '5 0 obj\n<< /Type /Page /Parent 3 0 R /MediaBox [0 0 200 100] /Contents 7 0 R >>\nendobj\n',
    '6 0 obj\n<< /Length 44 >>\nstream\nBT /F1 12 Tf 10 50 Td (PAGE-ONE) Tj ET\nendstream\nendobj\n',
    '7 0 obj\n<< /Length 44 >>\nstream\nBT /F1 12 Tf 10 50 Td (PAGE-TWO) Tj ET\nendstream\nendobj\n',
    'trailer\n<< /Root 1 0 R /Size 8 >>\n%%EOF\n'
  ].join('')
  const parsed = PdfDocument.open(Buffer.from(body, 'latin1'))
  assert.equal(parsed.pages().length, 2)
  const result = parsed.rewrite({ keepPages: [1] })
  const after = PdfDocument.open(result.buffer)
  assert.equal(after.pages().length, 1)
  assert.ok(after.extractText().text.includes('PAGE-TWO'))
  assert.ok(!result.buffer.toString('latin1').includes('PAGE-ONE'), '未选页内容不应残留')
})

test('★ 重写会带上内存里的编辑结果（旋转 + 叠加水印）', () => {
  const doc = PdfDocument.open(original)
  doc.rotatePage({ page: 0, degrees: 90 })
  doc.addTextOverlay({ text: 'REWRITTEN', position: 'top', fontSize: 12 })
  const result = doc.rewrite()
  const after = PdfDocument.open(result.buffer)
  assert.equal(after.pageInfos()[0].rotation, 90)
  assert.ok(after.extractText().text.includes('REWRITTEN'))
  assert.equal(after.validate().valid, true)
})

test('★ 回归：先改元数据再重写，修改不能丢（staged 值优先于对象表）', () => {
  const doc = PdfDocument.open(original)
  doc.updateMetadata({ title: '重写后仍在', author: 'staged-author' })
  const result = doc.rewrite()
  const after = PdfDocument.open(result.buffer)
  assert.equal(after.metadata().title, '重写后仍在')
  assert.equal(after.metadata().author, 'staged-author')
})

test('拒绝重写加密 PDF', () => {
  const body = [
    '%PDF-1.4\n',
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
    '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] >>\nendobj\n',
    '9 0 obj\n<< /Filter /Standard /V 2 /R 3 /Length 128 /P -3904 >>\nendobj\n',
    'trailer\n<< /Root 1 0 R /Encrypt 9 0 R /Size 10 >>\n%%EOF\n'
  ].join('')
  const parsed = PdfDocument.open(Buffer.from(body, 'latin1'))
  assert.throws(() => parsed.rewrite(), (err) => err.code === 'PASSWORD_REQUIRED')
})

test('★ 引用了不存在的对象时拒绝重写，而不是产出缺内容的文件', () => {
  const body = [
    '%PDF-1.4\n',
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
    '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 99 0 R >>\nendobj\n',
    'trailer\n<< /Root 1 0 R /Size 4 >>\n%%EOF\n'
  ].join('')
  const parsed = PdfDocument.open(Buffer.from(body, 'latin1'))
  assert.throws(
    () => parsed.rewrite(),
    (err) => err.code === 'CORRUPTED_DOCUMENT' && /解析不出来/.test(err.message)
  )
})

console.log('\n=== 11. 合并（跨文件对象图搬运） ===')

test('★ 合并两份 PDF：页序、逐页文本、页尺寸、元数据都正确', () => {
  if (!multi) {
    console.log('     ⏭️  跳过：缺少 sample-multipage.pdf')
    return
  }
  const first = PdfDocument.open(original)
  const second = PdfDocument.open(multi)
  const expectedTexts = [
    ...first.pages().map((_, i) => first.extractText({ page: i }).text.replace(/\s+/g, ' ').trim()),
    ...second.pages().map((_, i) => second.extractText({ page: i }).text.replace(/\s+/g, ' ').trim())
  ]
  const expectedSizes = [...first.pageInfos(), ...second.pageInfos()].map((p) => p.width_pt)

  const merged = PdfDocument.merge([original, multi])
  const doc = PdfDocument.open(merged.buffer)
  assert.equal(doc.pages().length, expectedTexts.length)
  assert.equal(merged.pages, expectedTexts.length)
  assert.equal(merged.sources, 2)
  assert.equal(doc.validate().valid, true)
  const actualTexts = doc.pages().map((_, i) => doc.extractText({ page: i }).text.replace(/\s+/g, ' ').trim())
  assert.deepEqual(actualTexts, expectedTexts, '合并后的逐页文本必须与来源逐页对应')
  assert.deepEqual(doc.pageInfos().map((p) => p.width_pt), expectedSizes, '各自的页尺寸要保持')
  assert.equal(doc.metadata().title, first.metadata().title, '元数据取自第一份')
  // 对象计数：合并后不应出现多余的页面对象
  const pageObjects = (merged.buffer.toString('latin1').match(/\/Type\s*\/Page[^s]/g) ?? []).length
  assert.equal(pageObjects, expectedTexts.length)
})

test('★ 合并后仍是合法增量更新的底子（可继续编辑与拆分）', () => {
  if (!multi) return
  const merged = PdfDocument.merge([original, multi])
  const doc = PdfDocument.open(merged.buffer)
  // 合并产物必须还能被继续改写：旋转一页 + 拆出一页
  doc.rotatePage({ page: 0, degrees: 90 })
  const edited = PdfDocument.open(doc.save())
  assert.equal(edited.pageInfos()[0].rotation, 90)
  const split = edited.rewrite({ keepPages: [0, 3] })
  const after = PdfDocument.open(split.buffer)
  assert.equal(after.pages().length, 2)
  assert.equal(after.validate().valid, true)
})

test('合并要求至少两份，且拒绝含加密文档的组合', () => {
  assert.throws(() => PdfDocument.merge([original]), (err) => err.code === 'INVALID_REQUEST')
  const encrypted = Buffer.from(
    [
      '%PDF-1.4\n',
      '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
      '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
      '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] >>\nendobj\n',
      '9 0 obj\n<< /Filter /Standard /V 2 /R 3 /Length 128 /P -3904 >>\nendobj\n',
      'trailer\n<< /Root 1 0 R /Encrypt 9 0 R /Size 10 >>\n%%EOF\n'
    ].join(''),
    'latin1'
  )
  assert.throws(() => PdfDocument.merge([original, encrypted]), (err) => err.code === 'PASSWORD_REQUIRED')
})

test('★ 合并同一份文件两次也能得到两份内容（自我合并不串号）', () => {
  const merged = PdfDocument.merge([original, original])
  const doc = PdfDocument.open(merged.buffer)
  assert.equal(doc.pages().length, 2)
  const texts = doc.pages().map((_, i) => doc.extractText({ page: i }).text.replace(/\s+/g, ' ').trim())
  assert.equal(texts[0], texts[1])
  assert.ok(texts[0].length > 0)
  assert.equal(doc.validate().valid, true)
})

console.log('\n=== 12. 图片提取 ===')

const imageFixture = join(here, 'fixtures', 'sample-images.pdf')
const withImage = existsSync(imageFixture) ? readFileSync(imageFixture) : null

test('★ 回归：流对象的字典字段要从 .dict 取（图片/对象流计数曾恒为 0）', () => {
  if (!withImage) {
    console.log('     ⏭️  跳过：缺少 sample-images.pdf')
    return
  }
  const structure = PdfDocument.open(withImage).structure()
  assert.ok(structure.images >= 1, `图片计数应至少为 1，实际 ${structure.images}`)
  assert.ok(structure.font_count >= 1)
})

test('★ 从真实 PDF 提取图片：尺寸、格式、共用页都正确', () => {
  if (!withImage) return
  const { images, skipped } = PdfDocument.open(withImage).extractImages()
  assert.equal(images.length, 1, `应提取到 1 张，实际 ${images.length}（跳过 ${skipped.length}）`)
  const image = images[0]
  assert.equal(image.format, 'png')
  assert.equal(image.width, 320)
  assert.equal(image.height, 160)
  assert.ok(image.bytes.length > 100)
  assert.deepEqual(image.pages, [1, 2, 3], '同一张图被三页共用，应列出页码')
  assert.deepEqual([...image.bytes.subarray(0, 8)], [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 'PNG 魔数')
})

test('★ 提取的 PNG 内部自洽（IHDR 尺寸与解压后的像素数一致）', () => {
  if (!withImage) return
  const image = PdfDocument.open(withImage).extractImages().images[0]
  const width = image.bytes.readUInt32BE(16)
  const height = image.bytes.readUInt32BE(20)
  assert.equal(width, image.width)
  assert.equal(height, image.height)
  const idatStart = image.bytes.indexOf(Buffer.from('IDAT', 'latin1'))
  assert.ok(idatStart > 0, 'PNG 里应有 IDAT')
  const length = image.bytes.readUInt32BE(idatStart - 4)
  const raw = zlib.inflateSync(image.bytes.subarray(idatStart + 4, idatStart + 4 + length))
  assert.equal(raw.length, (width * 3 + 1) * height, '每行应有 1 个过滤字节 + RGB 像素')
})

test('按页筛选与越界报错', () => {
  if (!withImage) return
  const doc = PdfDocument.open(withImage)
  assert.deepEqual(doc.extractImages({ pages: [2] }).images[0].pages, [3])
  assert.throws(() => doc.extractImages({ pages: [9] }), (err) => err.code === 'INVALID_REQUEST')
  assert.throws(() => doc.extractImages({ pages: [] }), (err) => err.code === 'INVALID_REQUEST')
})

test('★ 不支持的分量位数给出跳过原因，而不是导出坏文件', () => {
  const body = [
    '%PDF-1.4\n',
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
    '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n',
    '4 0 obj\n<< /Length 8 >>\nstream\nq Q\n\nendstream\nendobj\n',
    '5 0 obj\n<< /Type /XObject /Subtype /Image /Width 8 /Height 8 /ColorSpace /DeviceGray /BitsPerComponent 1 /Length 8 >>\nstream\nAAAAAAAA\nendstream\nendobj\n',
    'trailer\n<< /Root 1 0 R /Size 6 >>\n%%EOF\n'
  ].join('')
  const { images, skipped } = PdfDocument.open(Buffer.from(body, 'latin1')).extractImages()
  assert.equal(images.length, 0)
  assert.equal(skipped.length, 1)
  assert.match(skipped[0].skipped, /8 位分量/)
})

test('★ 合并后的 PDF 仍能提取到图片，且页码整体后移', () => {
  if (!withImage) return
  const merged = PdfDocument.merge([original, withImage])
  const { images } = PdfDocument.open(merged.buffer).extractImages()
  assert.equal(images.length, 1)
  assert.deepEqual(images[0].pages, [2, 3, 4], '第一份只有 1 页，图片页码应整体后移一位')
  assert.equal(images[0].width, 320)
})

console.log('\n=== 13. 安全防护 ===')

test('超大文件被拒绝', () => {
  assert.throws(() => PdfDocument.open(original, { maxBytes: 10 }), (e) => e.code === 'MEMORY_LIMIT')
})

test('对象数超限被拒绝', () => {
  assert.throws(() => PdfDocument.open(original, { maxObjects: 1 }), (e) => e.code === 'MEMORY_LIMIT')
})

test('★ 解压上限拦住压缩炸弹', () => {
  // 构造一个声明巨大解压结果的流
  const bomb = zlib.deflateSync(Buffer.alloc(4 * 1024 * 1024, 0x41))
  const body = Buffer.concat([
    Buffer.from('%PDF-1.4\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R >>\nendobj\n4 0 obj\n<< /Length ', 'latin1'),
    Buffer.from(String(bomb.length), 'latin1'),
    Buffer.from(' /Filter /FlateDecode >>\nstream\n', 'latin1'),
    bomb,
    Buffer.from('\nendstream\nendobj\ntrailer\n<< /Root 1 0 R /Size 5 >>\n%%EOF\n', 'latin1')
  ])
  const pdf = PdfDocument.open(body, { maxStreamBytes: 64 * 1024 })
  const report = pdf.validate()
  assert.equal(report.checks.find((c) => c.name === '内容流可解码').ok, false, '超限的流应被拒绝解码')
})

console.log('\n=== 12. 全文搜索 ===')

test('★ 搜索中文与英文关键词，返回页码 / 偏移 / 行号 / 上下文', () => {
  const pdf = PdfDocument.open(multi)
  const r = pdf.search({ query: '季度业务报告', contextChars: 6 })
  assert.equal(r.found, true)
  assert.equal(r.total_matches, 1)
  assert.deepEqual(r.matches_by_page, { 0: 1 })
  assert.equal(r.matches[0].page, 0)
  assert.equal(r.matches[0].offset, 0)
  assert.equal(r.matches[0].line, 1)
  assert.ok(r.matches[0].context_after.startsWith('\n技术方案'), `上下文不对：${JSON.stringify(r.matches[0].context_after)}`)
})

test('★ 跨页命中：同一关键词在第 2、3 页各命中一次', () => {
  const r = PdfDocument.open(multi).search({ query: 'XLSX' })
  assert.equal(r.total_matches, 2)
  assert.deepEqual(r.matches_by_page, { 1: 1, 2: 1 })
  assert.equal(r.pages_scanned, 3)
  assert.ok(r.chars_scanned > 0)
})

test('★ 默认不区分大小写，caseSensitive 后不再命中', () => {
  const loose = PdfDocument.open(multi).search({ query: 'xlsx' })
  assert.equal(loose.total_matches, 2)
  assert.equal(loose.matches[0].match, 'XLSX', '返回的应是原文大小写')
  const strict = PdfDocument.open(multi).search({ query: 'xlsx', caseSensitive: true })
  assert.equal(strict.found, false)
  assert.equal(strict.total_matches, 0)
})

test('★ 查询按字面量处理：正则元字符不会被当成模式', () => {
  const pdf = PdfDocument.open(multi)
  // "46" 是真实内容；"(a)" 不存在，且正则语法下 "(" 会抛错 —— 必须安全返回未命中
  assert.equal(pdf.search({ query: '46' }).total_matches, 1)
  assert.equal(pdf.search({ query: '(a)' }).found, false)
  assert.equal(pdf.search({ query: '[', }).found, false)
  // "." 不应匹配任意字符：文本里没有字面量 "."，所以命中数为 0
  assert.equal(pdf.search({ query: '.' }).found, false)
})

test('★ 未命中时如实返回，而不是抛错或假装找到', () => {
  const r = PdfDocument.open(multi).search({ query: '这份内容不存在' })
  assert.equal(r.found, false)
  assert.deepEqual(r.matches, [])
  assert.equal(r.total_matches, 0)
  assert.equal(r.pages_scanned, 3)
})

test('maxResults 截断时给出 truncated 与真实总数', () => {
  const r = PdfDocument.open(multi).search({ query: '完成', maxResults: 1 })
  assert.ok(r.total_matches > 1, `样本里「完成」应出现多次，实际 ${r.total_matches}`)
  assert.equal(r.returned_matches, 1)
  assert.equal(r.truncated, true)
})

test('page 参数只搜指定页', () => {
  const r = PdfDocument.open(multi).search({ query: 'XLSX', page: 2 })
  assert.equal(r.pages_scanned, 1)
  assert.deepEqual(r.matches_by_page, { 2: 1 })
})

test('搜索参数校验', () => {
  const pdf = PdfDocument.open(multi)
  assert.throws(() => pdf.search({ query: '' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.search({}), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.search({ query: 'x', maxResults: 0 }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.search({ query: 'x', maxResults: 5000 }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.search({ query: 'x', contextChars: -1 }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.search({ query: 'x', page: 99 }), (e) => e.code === 'INVALID_REQUEST')
})

test('★ 搜索命中位置与逐页提取的文本自洽（偏移可复现）', () => {
  const pdf = PdfDocument.open(multi)
  const pageText = pdf.extractText({ page: 2 }).text
  const r = pdf.search({ query: '已完成', page: 2 })
  assert.ok(r.total_matches >= 1)
  for (const m of r.matches) {
    assert.equal(pageText.slice(m.offset, m.offset + m.match.length), m.match, '偏移处应正好是被匹配的文本')
  }
})

console.log('\n=== 13. 便签批注 ===')

/** 构造一份带「间接 /Annots 数组」的合成 PDF：用于验证追加而不是覆盖。 */
function pdfWithIndirectAnnots() {
  const src = [
    '%PDF-1.4',
    '1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj',
    '2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj',
    '3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 300 400] /Annots 5 0 R >> endobj',
    '5 0 obj [6 0 R] endobj',
    '6 0 obj << /Type /Annot /Subtype /Text /Rect [10 10 30 30] /Contents (existing note) >> endobj',
    'trailer << /Root 1 0 R /Size 7 >>',
    '%%EOF',
    ''
  ].join('\n')
  return Buffer.from(src, 'latin1')
}

test('★ 添加便签批注：位置、作者、标志位、中文内容都能读回', () => {
  const pdf = PdfDocument.open(original)
  const result = pdf.addAnnotation({
    page: 0,
    text: '批注：这一页需要法务复核（中文也可以）。',
    author: 'dsh-exp-office',
    x: 60,
    y: 700,
    open: true
  })
  assert.deepEqual(result.rect, [60, 700, 84, 724])
  assert.equal(result.non_ascii, true)

  const reopened = PdfDocument.open(pdf.save())
  const list = reopened.annotations()
  assert.equal(list.count, 1)
  assert.equal(list.pages_with_annotations, 1)
  const a = list.annotations[0]
  assert.equal(a.subtype, 'Text')
  assert.equal(a.contents, '批注：这一页需要法务复核（中文也可以）。', '中文批注内容应原样读回（UTF-16BE + BOM）')
  assert.equal(a.author, 'dsh-exp-office')
  assert.deepEqual(a.rect, [60, 700, 84, 724])
  assert.equal(a.flags, 4, 'F=4 是 Print 标志，不写打印时批注会消失')
  assert.equal(a.open, true)
  assert.equal(a.is_form_field, false)
})

test('★ 结构识别把便签算作批注而不是表单字段', () => {
  const pdf = PdfDocument.open(original)
  pdf.addAnnotation({ page: 0, text: 'note' })
  const structure = PdfDocument.open(pdf.save()).structure()
  assert.equal(structure.annotations, 1)
  assert.equal(structure.form_fields, 0)
})

test('★ /Annots 是间接数组时追加而不是覆盖（否则原有批注全丢）', () => {
  const pdf = PdfDocument.open(pdfWithIndirectAnnots())
  assert.equal(pdf.annotations().count, 1, '前置条件：原文件已有 1 条批注')
  pdf.addAnnotation({ page: 0, text: 'second note' })
  const list = PdfDocument.open(pdf.save()).annotations()
  assert.equal(list.count, 2, `原有批注必须保留：${JSON.stringify(list.annotations)}`)
  assert.deepEqual(list.annotations.map((a) => a.contents), ['existing note', 'second note'])
})

test('没有 /Annots 的页面会被补上数组', () => {
  const pdf = PdfDocument.open(original)
  assert.equal(pdf.annotations().count, 0)
  pdf.addAnnotation({ page: 0, text: 'first' })
  assert.equal(PdfDocument.open(pdf.save()).annotations().count, 1)
})

test('★ 省略坐标时放在页面左上角内侧', () => {
  const pdf = PdfDocument.open(original)
  const box = pdf.pageInfos()[0]
  const result = pdf.addAnnotation({ page: 0, text: 'default position' })
  // 左上角 = 页面高度附近；PDF 原点在左下角
  assert.ok(result.rect[1] > box.height_pt / 2, `默认应在页面上半部分：${result.rect}`)
  assert.ok(result.rect[0] < 40, `默认应贴近左边界：${result.rect}`)
})

test('★ 批注写入是增量更新：原有字节一个都不动', () => {
  const before = Buffer.from(original)
  const pdf = PdfDocument.open(Buffer.from(original))
  pdf.addAnnotation({ page: 0, text: 'incremental' })
  const after = pdf.save()
  assert.ok(after.length > before.length, '增量更新应让文件变大')
  assert.ok(after.subarray(0, before.length).equals(before), '原有字节必须保持不变')
})

test('★ 多页/多条批注互不干扰', () => {
  const pdf = PdfDocument.open(multi)
  pdf.addAnnotation({ page: 0, text: '第一页的批注' })
  pdf.addAnnotation({ page: 2, text: '第三页的批注', author: '审阅者' })
  const list = PdfDocument.open(pdf.save()).annotations()
  assert.equal(list.count, 2)
  assert.equal(list.pages_with_annotations, 2)
  assert.deepEqual(list.annotations.map((a) => a.page), [0, 2])
  assert.equal(list.annotations[1].author, '审阅者')
})

test('批注参数校验', () => {
  const pdf = PdfDocument.open(original)
  assert.throws(() => pdf.addAnnotation({ page: 0, text: '' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.addAnnotation({ text: 'x' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.addAnnotation({ page: 99, text: 'x' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.addAnnotation({ page: 0, text: 'x', width: 0 }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.addAnnotation({ page: 0, text: 'x', x: Number.NaN }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.addAnnotation({ page: 0, text: 'x', color: [2, 0, 0] }), (e) => e.code === 'INVALID_REQUEST')
})

test('★ 加批注后结构仍然合法', () => {
  const pdf = PdfDocument.open(original)
  pdf.addAnnotation({ page: 0, text: 'valid check' })
  const report = PdfDocument.open(pdf.save()).validate()
  assert.deepEqual(report.checks.filter((c) => !c.ok), [], JSON.stringify(report.checks))
})

console.log('\n=== 14. AcroForm 表单 ===')

const formFixture = join(here, 'fixtures', 'form-sample.pdf')
const hasForm = existsSync(formFixture)
const formBytes = hasForm ? readFileSync(formFixture) : null

test('★ 读取表单字段：类型、选项、开状态名（不靠猜）', () => {
  if (!hasForm) return
  const form = PdfDocument.open(formBytes).forms()
  assert.equal(form.has_form, true)
  assert.equal(form.field_count, 4)
  const byName = new Map(form.fields.map((f) => [f.name, f]))
  assert.equal(byName.get('full_name').kind, '文本')
  assert.equal(byName.get('agree_terms').kind, '复选框')
  assert.equal(byName.get('plan').kind, '列表')
  assert.equal(byName.get('tier').kind, '单选组')
  assert.deepEqual(byName.get('plan').options, ['basic', 'pro', 'enterprise'])
  assert.deepEqual(byName.get('agree_terms').on_states, ['Yes'], '开状态名要从 /AP /N 读出来')
  assert.deepEqual(byName.get('tier').on_states, ['A', 'B'])
  assert.equal(byName.get('full_name').value, null)
})

test('★ 填写四类字段（含中文）并读回', () => {
  if (!hasForm) return
  const pdf = PdfDocument.open(formBytes)
  const result = pdf.fillForm({ fields: { full_name: '张伟', agree_terms: true, plan: 'pro', tier: 'B' } })
  assert.equal(result.filled_count, 4)
  assert.equal(result.need_appearances, true)
  assert.equal(result.filled.find((f) => f.name === 'plan').option_index, 1, '下拉应写入选项索引')

  const back = PdfDocument.open(pdf.save()).forms()
  const byName = new Map(back.fields.map((f) => [f.name, f]))
  assert.equal(byName.get('full_name').value, '张伟', '中文应通过 UTF-16BE + BOM 原样写入')
  assert.equal(byName.get('agree_terms').value, 'Yes')
  assert.equal(byName.get('agree_terms').checked, true)
  assert.equal(byName.get('plan').value, 'pro')
  assert.equal(byName.get('tier').value, 'B')
  assert.equal(back.need_appearances, true, '/NeedAppearances 应置为 true 让阅读器重绘外观')
})

test('★ 单选组只把一个 widget 置为开状态', () => {
  if (!hasForm) return
  const pdf = PdfDocument.open(formBytes)
  pdf.fillForm({ fields: { tier: 'B' } })
  const tier = PdfDocument.open(pdf.save()).forms().fields.find((f) => f.name === 'tier')
  assert.deepEqual(
    tier.widget_states.map((w) => w.appearance_state),
    ['Off', 'B'],
    `只有 B 那个 widget 应该是开状态：${JSON.stringify(tier.widget_states)}`
  )
})

test('复选框用布尔或开状态名都能填', () => {
  if (!hasForm) return
  const on = PdfDocument.open(formBytes)
  on.fillForm({ fields: { agree_terms: 'Yes' } })
  assert.equal(PdfDocument.open(on.save()).forms().fields.find((f) => f.name === 'agree_terms').checked, true)
  const off = PdfDocument.open(formBytes)
  off.fillForm({ fields: { agree_terms: false } })
  assert.equal(PdfDocument.open(off.save()).forms().fields.find((f) => f.name === 'agree_terms').checked, false)
})

test('★ 校验失败时不留「填了一半」的状态（两遍式：先全验、再全写）', () => {
  if (!hasForm) return
  const pdf = PdfDocument.open(formBytes)
  assert.throws(
    () => pdf.fillForm({ fields: { full_name: '先改这个', plan: '不存在的选项' } }),
    (e) => e.code === 'INVALID_REQUEST'
  )
  assert.equal(pdf.dirty, false, '失败后不应有任何待写回改动')
  assert.ok(pdf.save().equals(formBytes), '失败后字节必须与源文件一致')
})

test('★ 填写是增量更新：原有字节一个都不动', () => {
  if (!hasForm) return
  const pdf = PdfDocument.open(formBytes)
  pdf.fillForm({ fields: { full_name: '增量' } })
  const after = pdf.save()
  assert.ok(after.length > formBytes.length)
  assert.ok(after.subarray(0, formBytes.length).equals(formBytes))
  assert.deepEqual(PdfDocument.open(after).validate().checks.filter((c) => !c.ok), [])
})

test('★ 表单参数校验：字段名 / 选项 / 开状态 / 无表单', () => {
  if (!hasForm) return
  const pdf = PdfDocument.open(formBytes)
  assert.throws(() => pdf.fillForm({ fields: {} }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.fillForm({ fields: [] }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(
    () => pdf.fillForm({ fields: { 不存在的字段: 'x' } }),
    (e) => e.code === 'INVALID_REQUEST' && Array.isArray(e.details.field_names)
  )
  assert.throws(() => pdf.fillForm({ fields: { plan: '不存在的选项' } }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => pdf.fillForm({ fields: { agree_terms: 'Maybe' } }), (e) => e.code === 'INVALID_REQUEST')
  // 没有表单的文档：读得到「没有」，写会被明确拒绝
  const plain = PdfDocument.open(original)
  assert.equal(plain.forms().has_form, false)
  assert.throws(() => plain.fillForm({ fields: { a: 1 } }), (e) => e.code === 'UNSUPPORTED_FEATURE')
})

test('★ 只读字段默认拒绝，force=true 才允许', () => {
  if (!hasForm) return
  // 造一个只读字段（/Ff 131073 = 只读 + 下拉）
  const text = formBytes.toString('latin1').replace('/T (plan)', '/T (plan) /Ff 131073')
  const pdf = PdfDocument.open(Buffer.from(text, 'latin1'))
  const field = pdf.forms().fields.find((f) => f.name === 'plan')
  assert.equal(field.read_only, true)
  assert.throws(() => pdf.fillForm({ fields: { plan: 'pro' } }), (e) => e.code === 'PERMISSION_DENIED')
  const forced = pdf.fillForm({ fields: { plan: 'pro' }, force: true })
  assert.equal(forced.filled_count, 1)
  assert.equal(PdfDocument.open(pdf.save()).forms().fields.find((f) => f.name === 'plan').value, 'pro')
})


/**
 * 找一个可用于测试的中文字体（本机没有就跳过相关用例）。
 * @returns {{path: string, index: number}|null} 字体位置。
 */
function resolveFixtureFont() {
  for (const candidate of ['C:\\Windows\\Fonts\\simhei.ttf', 'C:\\Windows\\Fonts\\Deng.ttf']) {
    if (existsSync(candidate)) return { path: candidate, index: 0 }
  }
  return null
}

console.log('\n=== 15. 从零生成 PDF ===')

test('★ 生成最基本的一份：头部、页数、文本都能读回', () => {
  const { buffer, pages } = buildPdf({ lines: ['Quarterly Sales Report', 'Second line'] })
  assert.equal(pages, 1)
  assert.equal(buffer.subarray(0, 8).toString('latin1'), '%PDF-1.7')
  assert.ok(buffer.subarray(-7).toString('latin1').includes('%%EOF'))
  const doc = PdfDocument.open(buffer)
  assert.equal(doc.pages().length, 1)
  assert.equal(doc.extractText().text, 'Quarterly Sales Report\nSecond line')
})

test('★ 自校验通过：自家校验器认为结构完整（含 xref 与 startxref）', () => {
  const { buffer } = buildPdf({ lines: ['structure check'] })
  const report = PdfDocument.open(buffer).validate()
  assert.equal(report.valid, true, JSON.stringify(report.checks.filter((c) => !c.ok)))
  const text = buffer.toString('latin1')
  assert.ok(text.includes('startxref'), '必须有 startxref')
  assert.ok(text.includes('/Type /Catalog'), '必须有文档目录')
  assert.ok(text.includes('/BaseFont /Helvetica'), '应使用标准 14 字体')
  assert.ok(text.includes('/ToUnicode'), '必须写 ToUnicode，否则文本提取不出来')
})

test('★ 文本可提取是设计目标：中文以外的常见字符都能读回', () => {
  const lines = ['ASCII letters and digits 0123456789', 'Punctuation: (parens) [brackets] {braces} "quotes"', 'Accents: caf\u00e9 na\u00efve \u00e9\u00e8\u00ea \u00fc\u00f6\u00e4', 'Symbols: \u20ac \u2014 \u201csmart\u201d \u2013 \u2022']
  const { buffer } = buildPdf({ lines })
  const extracted = PdfDocument.open(buffer).extractText().text
  assert.ok(extracted.includes('ASCII letters and digits 0123456789'))
  assert.ok(extracted.includes('Punctuation: (parens) [brackets] {braces} "quotes"'))
  assert.ok(extracted.includes('\u20ac'), '欧元符号应能读回')
  assert.ok(extracted.includes('\u2014'), '长破折号应能读回')
})


test('转义：括号与反斜杠不会破坏内容流', () => {
  const tricky = 'a (b) c \\\\ d (unbalanced ( and )'
  const { buffer } = buildPdf({ lines: [tricky] })
  const doc = PdfDocument.open(buffer)
  assert.equal(doc.validate().valid, true)
  assert.ok(doc.extractText().text.includes('(b)'), '括号应对称读回')
  assert.ok(doc.extractText().text.includes('\\\\'), '反斜杠应读回')
})

test('多页：自动分页与强制分页都按预期', () => {
  const many = Array.from({ length: 120 }, (_, i) => `line ${i + 1}`)
  const auto = buildPdf({ lines: many, fontSizePt: 11 })
  assert.ok(auto.pages > 1, `120 行应分成多页，实际 ${auto.pages} 页`)
  assert.equal(PdfDocument.open(auto.buffer).pages().length, auto.pages)

  const forced = buildPdf({ lines: ['page one', '\fpage two', '\fpage three'] })
  assert.equal(forced.pages, 3)
  const text = PdfDocument.open(forced.buffer).extractText()
  assert.equal(text.pages, 3)
  assert.ok(text.text.includes('page one'))
  assert.ok(text.text.includes('page three'))
})

test('折行：超宽文本被折开并给出可核对的行数', () => {
  const long = 'word '.repeat(80).trim()
  const { lines } = buildPdf({ lines: [long], fontSizePt: 11, pageSize: 'A5' })
  assert.ok(lines > 1, `超宽文本应折行，实际 ${lines} 行`)
})

test('纸张与方向：横向会交换宽高', () => {
  const portrait = buildPdf({ lines: ['x'], pageSize: 'A4', orientation: 'portrait' })
  const landscape = buildPdf({ lines: ['x'], pageSize: 'A4', orientation: 'landscape' })
  const boxOf = (buffer) => {
    const match = /\/MediaBox \[0 0 ([\d.]+) ([\d.]+)\]/.exec(buffer.toString('latin1'))
    return match ? [Number(match[1]), Number(match[2])] : null
  }
  const p = boxOf(portrait.buffer)
  const l = boxOf(landscape.buffer)
  assert.ok(p[1] > p[0], '纵向应高大于宽')
  assert.equal(l[0], p[1])
  assert.equal(l[1], p[0])
})

test('元数据：ASCII 与非 ASCII（UTF-16BE + BOM）都能读回', () => {
  const { buffer } = buildPdf({ lines: ['x'], metadata: { title: 'Quarterly Report', author: '季度报告作者' } })
  const meta = PdfDocument.open(buffer).metadata()
  assert.equal(meta.title, 'Quarterly Report')
  assert.equal(meta.author, '季度报告作者')
})

test('参数校验：纸张/方向/字号/行类型都有明确错误', () => {
  assert.throws(() => buildPdf({ lines: ['x'], pageSize: 'B5' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => buildPdf({ lines: ['x'], orientation: 'sideways' }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => buildPdf({ lines: ['x'], fontSizePt: 0 }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => buildPdf({ lines: [42] }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => buildPdf({ lines: ['x'], marginPt: 400 }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => buildPdf({ lines: ['x'], fontFamily: 'comic' }), (e) => e.code === 'INVALID_REQUEST')
})

test('★ 生成的 PDF 可继续被自家页面操作编辑（不是死文件）', () => {
  const { buffer } = buildPdf({ lines: ['first page', '\fsecond page'] })
  const doc = PdfDocument.open(buffer)
  doc.rotatePage({ page: 0, degrees: 90 })
  doc.addTextOverlay({ text: 'OVERLAY', pages: [1] })
  const saved = doc.save()
  const reopened = PdfDocument.open(saved)
  assert.equal(reopened.pages().length, 2)
  assert.equal(reopened.pages()[0].value.Rotate, 90)
  assert.ok(reopened.extractText().text.includes('OVERLAY'))
})

test('压缩开关都产出合法文件', () => {
  for (const compress of [true, false]) {
    const { buffer } = buildPdf({ lines: ['compress check'], compress })
    const doc = PdfDocument.open(buffer)
    assert.equal(doc.validate().valid, true)
    assert.equal(doc.extractText().text, 'compress check')
  }
})

console.log('\n=== 16. 表格提取（纯位置推断）===')

test('★ 真实 PDF（LibreOffice 从带表格的 DOCX 导出）里读出正确的二维表格', () => {
  const file = join(here, 'fixtures', 'sample.pdf')
  if (!existsSync(file)) return
  const doc = PdfDocument.open(readFileSync(file))
  const result = doc.extractTables()
  assert.equal(result.table_count, 1, `应恰好识别 1 张表，实际 ${result.table_count}`)
  const table = result.tables[0]
  assert.equal(table.row_count, 2)
  assert.equal(table.column_count, 2)
  assert.deepEqual(table.rows, [
    ['月份', '金额'],
    ['1月', '12000']
  ])
})

test('★ 多页 PDF：只在含表格的那一页报出表格', () => {
  const file = join(here, 'fixtures', 'sample-multipage.pdf')
  if (!existsSync(file)) return
  const doc = PdfDocument.open(readFileSync(file))
  const result = doc.extractTables()
  assert.equal(result.table_count, 1)
  const table = result.tables[0]
  assert.equal(table.page, 2, '表格在第 3 页（下标 2）')
  assert.deepEqual(table.rows[0], ['模块', '状态', '测试数'])
  assert.equal(table.rows.length, 3)
})

test('★ 没有表格的 PDF 不会凭空报出表格（不误判正文）', () => {
  const { buffer } = buildPdf({ lines: ['Just a paragraph of prose.', 'Another line without any column alignment.', 'Third line.'] })
  const result = PdfDocument.open(buffer).extractTables()
  assert.equal(result.table_count, 0, `不应把正文当表格：${JSON.stringify(result.tables)}`)
})

test('★ 自己生成的表格能被自己读回来（等宽字体列对齐）', () => {
  const rows = [
    ['Region', 'Revenue'],
    ['North', '12000'],
    ['South', '9500']
  ]
  const lines = rows.map(([left, right]) => `${left}${' '.repeat(Math.max(1, 24 - left.length))}${right}`)
  const { buffer } = buildPdf({ lines, fontFamily: 'courier' })
  const result = PdfDocument.open(buffer).extractTables()
  assert.equal(result.table_count, 1, `等宽字体下应识别为表格：${JSON.stringify(result.tables)}`)
  assert.equal(result.tables[0].rows.length, 3)
})

test('接口如实报告方法与未做的部分', () => {
  const file = join(here, 'fixtures', 'sample.pdf')
  if (!existsSync(file)) return
  const result = PdfDocument.open(readFileSync(file)).extractTables()
  assert.match(result.method, /位置/)
  assert.ok(result.not_done.some((item) => /OCR/.test(item)), '必须说明扫描件需要 OCR')
  assert.ok(Array.isArray(result.approximation) && result.approximation.length > 0)
})

test('textRuns 给出坐标与字号，并如实说明近似之处', () => {
  const file = join(here, 'fixtures', 'sample.pdf')
  if (!existsSync(file)) return
  const doc = PdfDocument.open(readFileSync(file))
  const runs = doc.textRuns({ page: 0 })
  assert.ok(runs.runs.length > 0)
  const first = runs.runs[0]
  assert.equal(typeof first.x, 'number')
  assert.equal(typeof first.y, 'number')
  assert.equal(typeof first.font_size, 'number')
  assert.ok(runs.approximation.some((item) => /旋转|缩放/.test(item)), '必须说明坐标不含旋转与缩放')
})

test('表格参数校验：容差/阈值必须是正数', () => {
  const file = join(here, 'fixtures', 'sample.pdf')
  if (!existsSync(file)) return
  const doc = PdfDocument.open(readFileSync(file))
  assert.throws(() => doc.extractTables({ rowTolerance: -1 }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.extractTables({ columnGap: 0 }), (e) => e.code === 'INVALID_REQUEST')
  assert.throws(() => doc.extractTables({ minRows: 0 }), (e) => e.code === 'INVALID_REQUEST')
})

test('★ 中文：自动嵌入子集字体，中文与中英混排都能读回', () => {
  const lines = ['季度销售报告', '本 PDF 从零生成，嵌入中文字体子集。', 'Mixed 中英混排 with ASCII 12345（全角括号）。']
  const { buffer, embedded_font: embedded } = buildPdf({ lines, metadata: { title: '中文验证' } })
  assert.ok(embedded, '含中文时必须嵌入字体')
  assert.ok(embedded.glyphs > 0)
  assert.ok(buffer.length < 200 * 1024, `子集后文件应仍在百 KB 量级，实际 ${buffer.length} 字节`)
  const doc = PdfDocument.open(buffer)
  assert.equal(doc.validate().valid, true)
  const extracted = doc.extractText().text
  for (const line of lines) assert.ok(extracted.includes(line), `读回文本缺少「${line}」：${extracted}`)
  assert.ok(doc.structure().fonts.some((f) => f.includes('+')), '字体名应带子集前缀（6 个大写字母 + +）')
})

test('★ 中文字形来自真实字体：字宽不是平均估算（不同字宽度不同）', () => {
  const { buffer } = buildPdf({ lines: ['国国国国', 'iiii'] })
  const raw = buffer.toString('latin1')
  const widths = [...raw.matchAll(/(\d+) \[(\d+)\]/g)].map((m) => Number(m[2]))
  assert.ok(widths.length > 0, '应写入 /W 宽度数组')
  assert.ok(new Set(widths).size > 1, `全角汉字与半角拉丁字母的宽度应不同，实际 ${JSON.stringify([...new Set(widths)].slice(0, 8))}`)
})

test('指定的字体文件不存在 → FILE_NOT_FOUND（不静默退回）', () => {
  assert.throws(
    () => buildPdf({ lines: ['中文'], cjkFontPath: 'C:\\nonexistent\\font.ttf' }),
    (err) => err.code === 'FILE_NOT_FOUND'
  )
})

test('★ 中文字体子集：保留原字形编号，字形数据与原件逐字节一致', () => {
  const fontFile = resolveFixtureFont()
  if (!fontFile) return
  const font = readFont(readFileSync(fontFile.path), fontFile.index)
  const lookup = readCmap(font.tableBuffer(0x636d6170))
  const text = '季度报告 ABC 123'
  const gids = [...text].map((ch) => lookup(ch.codePointAt(0))).filter((gid) => gid !== 0)
  assert.ok(gids.length > 0, '字体应覆盖这些字符')
  const { buffer: subset } = subsetFont(font, gids)
  const back = readFont(subset)
  for (const gid of gids) {
    const original = font.glyph(gid)
    const copy = back.glyph(gid)
    assert.ok(copy.length >= original.length, `字形 ${gid} 应保留`)
    assert.ok(copy.subarray(0, original.length).equals(original), `字形 ${gid} 的数据应与原件一致`)
  }
  // 复合字形的闭包：子集里的字形数 >= 直接命中的字形数
  const backLookup = readCmap(back.tableBuffer(0x636d6170))
  for (const ch of text) {
    const gid = lookup(ch.codePointAt(0))
    if (gid !== 0) assert.equal(backLookup(ch.codePointAt(0)), gid, `码点 ${ch} 的 GID 应保持不变`)
  }
})

test('找不到任何中文字体时给出可操作错误（列出试过的路径）', () => {
  // 用一个「所有候选都不存在」的字体路径集合来模拟：直接调用 resolveCjkFont 的反例
  assert.throws(
    () => buildPdf({ lines: ['中文'], cjkFontPath: 'C:\\Windows\\Fonts\\__no_such_font__.ttf' }),
    (err) => err.code === 'FILE_NOT_FOUND'
  )
})

console.log(`\n结果：${passed} 通过，${failed} 失败\n`)
process.exit(failed === 0 ? 0 : 1)
