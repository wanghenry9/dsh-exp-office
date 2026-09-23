/**
 * 用「从零生成 PDF」能力造一份样本，供真实第三方解析器（Word 的 PDF 重排）独立校验。
 *
 * 文本刻意全用 ASCII：标准 14 字体只支持 WinAnsi，而 Word 的重排读取是最可靠的
 * 第三方信号（本机没有 pdftotext / pdfinfo）。中文生成的边界见 README。
 *
 * 运行：node test/make-pdf-created-fixture.mjs
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { buildPdf } from '../lib/pdf-writer.js'
import { PdfDocument } from '../lib/pdf.js'

const here = dirname(fileURLToPath(import.meta.url))
const outDir = join(here, 'fixtures')
mkdirSync(outDir, { recursive: true })

const lines = [
  'Quarterly Sales Report',
  '',
  'This PDF was generated from scratch by dsh-exp-office (no template, no third-party library).',
  'It uses the standard 14 fonts, so nothing has to be embedded.',
  'Marker: PDF-FROM-SCRATCH-MARKER',
  '',
  'Escaping check: (parentheses) and a backslash \\ are written literally.',
  '\fPage two starts here, forced by a form-feed marker in the input.',
  'Text extraction works because a /ToUnicode CMap is written as well.'
]

const { buffer, pages, lines: rendered } = buildPdf({
  lines,
  pageSize: 'A4',
  fontSizePt: 12,
  metadata: { title: 'From-scratch PDF check', author: 'dsh-exp-office' }
})

const out = join(outDir, 'created-from-scratch.pdf')
writeFileSync(out, buffer)

const doc = PdfDocument.open(buffer)
const extracted = doc.extractText()
console.log(`已生成：${out}`)
console.log(`规模：${pages} 页 / 排版 ${rendered} 行 / ${buffer.length} 字节`)
console.log(`自读：页数 ${doc.pages().length}｜提取字符 ${extracted.length}｜校验 ${doc.validate().valid ? '通过' : '失败'}`)
console.log(`自读的标记行是否出现：${extracted.text.includes('PDF-FROM-SCRATCH-MARKER')}`)
