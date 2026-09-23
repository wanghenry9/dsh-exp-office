// 生成带 AcroForm 的 PDF 表单夹具：文本 / 复选框 / 下拉 / 单选组 四类字段齐全。
// 手工拼 PDF（不依赖 Office），因此夹具可复现、可逐字节对照。
// 用法：node test/make-pdf-form-fixture.mjs
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const here = path.dirname(fileURLToPath(import.meta.url))
const target = path.join(here, 'fixtures', 'form-sample.pdf')

const objects = [
  '1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm 10 0 R >>\nendobj\n',
  '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
  '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /Font << /Helv 30 0 R >> >> ' +
    '/Annots [4 0 R 5 0 R 6 0 R 7 0 R 8 0 R] /Contents 40 0 R >>\nendobj\n',
  // 文本字段
  '4 0 obj\n<< /Type /Annot /Subtype /Widget /FT /Tx /T (full_name) /Rect [50 240 250 262] /F 4 ' +
    '/DA (/Helv 10 Tf 0 g) /AP << /N 20 0 R >> >>\nendobj\n',
  // 复选框（开状态名故意用 Yes，验证「读出来而不是猜」）
  '5 0 obj\n<< /Type /Annot /Subtype /Widget /FT /Btn /T (agree_terms) /Rect [50 205 70 225] /F 4 ' +
    '/AP << /N << /Yes 21 0 R /Off 22 0 R >> >> >>\nendobj\n',
  // 下拉（可编辑关闭，带三个选项）
  '6 0 obj\n<< /Type /Annot /Subtype /Widget /FT /Ch /T (plan) /Rect [50 165 200 185] /F 4 ' +
    '/Opt [(basic) (pro) (enterprise)] /AP << /N 23 0 R >> >>\nendobj\n',
  // 单选组：两个 widget 各有自己的开状态名（A / B），父字段是 9 0 R
  '7 0 obj\n<< /Type /Annot /Subtype /Widget /FT /Btn /Ff 32768 /T (tier) /Rect [50 125 70 145] /F 4 ' +
    '/AP << /N << /A 24 0 R /Off 25 0 R >> >> /Parent 9 0 R >>\nendobj\n',
  '8 0 obj\n<< /Type /Annot /Subtype /Widget /FT /Btn /Ff 32768 /T (tier) /Rect [110 125 130 145] /F 4 ' +
    '/AP << /N << /B 26 0 R /Off 27 0 R >> >> /Parent 9 0 R >>\nendobj\n',
  '9 0 obj\n<< /FT /Btn /Ff 32768 /T (tier) /Kids [7 0 R 8 0 R] /V /Off >>\nendobj\n',
  '10 0 obj\n<< /Fields [4 0 R 5 0 R 6 0 R 9 0 R] /NeedAppearances false /DA (/Helv 10 Tf 0 g) ' +
    '/DR << /Font << /Helv 30 0 R >> >> >>\nendobj\n',
  // 外观流：内容无关紧要，但 /AP /N 必须存在，否则读不出开状态名
  '20 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n',
  '21 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n',
  '22 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n',
  '23 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n',
  '24 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n',
  '25 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n',
  '26 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n',
  '27 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n',
  '30 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj\n',
  '40 0 obj\n<< /Length 44 >>\nstream\nBT /Helv 12 Tf 50 280 Td (Form sample) Tj ET\nendstream\nendobj\n'
]

let pdf = '%PDF-1.4\n'
const offsets = []
for (const obj of objects) {
  offsets.push(pdf.length)
  pdf += obj
}
const xrefStart = pdf.length
pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
for (const offset of offsets) pdf += `${String(offset).padStart(10, '0')} 00000 n \n`
pdf += `trailer\n<< /Root 1 0 R /Size ${objects.length + 1} >>\nstartxref\n${xrefStart}\n%%EOF\n`

fs.writeFileSync(target, Buffer.from(pdf, 'latin1'))

const { PdfDocument } = await import('../lib/pdf.js')
const reopened = PdfDocument.open(fs.readFileSync(target))
const form = reopened.forms()
console.log(`FIXTURE: ${target}`)
console.log(`HAS_FORM: ${form.has_form}  字段数: ${form.field_count}`)
for (const field of form.fields) {
  console.log(`  ${field.name} | ${field.kind} | 值=${JSON.stringify(field.value)} | 选项=${JSON.stringify(field.options)} | 开状态=${JSON.stringify(field.on_states)}`)
}
console.log(`VALID: ${reopened.validate().checks.every((c) => c.ok)}`)

// 第二份：填好值的版本（供真机验证「填写后文件仍能被第三方解析器读取」）
const filledTarget = path.join(here, 'fixtures', 'form-filled.pdf')
const filled = PdfDocument.open(fs.readFileSync(target))
filled.fillForm({ fields: { full_name: '张伟', agree_terms: true, plan: 'pro', tier: 'B' } })
fs.writeFileSync(filledTarget, filled.save())
const back = PdfDocument.open(fs.readFileSync(filledTarget)).forms()
console.log(`FIXTURE: ${filledTarget}`)
console.log(`FILLED: ${back.fields.map((f) => `${f.name}=${JSON.stringify(f.value)}`).join(' ')}`)
console.log(`NEED_APPEARANCES: ${back.need_appearances}`)
