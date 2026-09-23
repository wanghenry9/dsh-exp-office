/**
 * 生成「段落样式 / 字符样式 / 分页符」的验证样本，供真实 Word 核对。
 *
 * 产物 test/fixtures/report-styled.docx：第一段改 Heading2、第二段加字符样式 Strong、
 * 第一段之后插分页符。Word 应读到：第一段样式变成 Heading2、页数从 1 变成 2。
 *
 * 运行：node test/make-docx-style-fixture.mjs
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { DocxDocument } from '../lib/docx.js'

const here = dirname(fileURLToPath(import.meta.url))
const source = join(here, 'fixtures', 'report.docx')
const target = join(here, 'fixtures', 'report-styled.docx')

const doc = DocxDocument.open(readFileSync(source))
const paragraph = doc.setParagraphStyle({ index: 0, styleId: 'Heading2' })
const character = doc.setCharacterStyle({ index: 1, styleId: 'Strong' })
const brk = doc.insertPageBreak({ after: 0 })
const bytes = doc.save()
writeFileSync(target, bytes)
console.log(
  `已生成：test/fixtures/report-styled.docx（${bytes.length} 字节）段落样式 ${paragraph.from}→${paragraph.style_id}、字符样式 ${character.runs} 个 run、分页符插在第 ${brk.index} 段`
)
