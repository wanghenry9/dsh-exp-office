/**
 * 用「从零生成 PDF」的中文能力造样本，并用**第三方解析器**（Word 的 PDF 重排）独立核对文本。
 *
 * 为什么要单独跑这一步：自读能过只证明「自己写的自己认」，而嵌入字体子集这件事
 * 最容易出的错就是「别的阅读器根本不认这份字体」—— 实测就踩过一次：
 * `/W` 数组被当成 PDF 名称写出去，字体字典非法，Word 读出来 **0 个字**（自读却完全正常）。
 *
 * 用法：
 *   node test/make-pdf-cjk-fixture.mjs                 # 只生成样本并自读
 *   node test/make-pdf-cjk-fixture.mjs --verify        # 生成后调用 Word 独立读取并逐字比对
 *
 * 比对通道：Word 解析出的文本写进临时文件（UTF-8），再按**字符**比对 ——
 * 中文经 stdout 回传会被控制台编码弄坏，写文件才可靠。
 */
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { spawnSync } from 'node:child_process'
import { PdfDocument } from '../lib/pdf.js'
import { buildPdf } from '../lib/pdf-writer.js'

const here = dirname(fileURLToPath(import.meta.url))
const outDir = join(here, 'fixtures')
mkdirSync(outDir, { recursive: true })

const LINES = [
  '季度销售报告（中文 PDF 从零生成验证）',
  '',
  '本文件由 dsh-exp-office 从零生成：页面树、内容流、字体资源与 xref 都是自己写的。',
  '正文含中文时，会自动嵌入一份**只带用到的字形**的 TrueType 子集，文件仍只有几十 KB。',
  'Mixed 中英混排 with ASCII, numbers 12345 and punctuation（全角括号）。',
  '\f第二页：分页、折行与文本提取都正常。',
  'Marker: CJK-PDF-MARKER'
]

const { buffer, pages, lines, embedded_font: embedded } = buildPdf({
  lines: LINES,
  pageSize: 'A4',
  fontSizePt: 12,
  metadata: { title: '中文 PDF 从零生成', author: 'dsh-exp-office', subject: '字体子集嵌入验证' }
})

const out = join(outDir, 'created-cjk.pdf')
writeFileSync(out, buffer)

const doc = PdfDocument.open(buffer)
const extracted = doc.extractText().text
console.log(`已生成：${out}`)
console.log(`规模：${pages} 页 / 排版 ${lines} 行 / ${buffer.length} 字节（${(buffer.length / 1024).toFixed(1)} KB）`)
console.log(`嵌入字体：${embedded ? `${embedded.name}｜${embedded.characters} 个字符 / ${embedded.glyphs} 个字形｜子集未压缩 ${(embedded.subset_bytes / 1024).toFixed(1)} KB` : '无（纯拉丁）'}`)
console.log(`自读：页数 ${doc.pages().length}｜提取 ${extracted.length} 字符｜校验 ${doc.validate().valid ? '通过' : '失败'}`)
console.log(`自读是否含中文标记：${extracted.includes('中文 PDF 从零生成验证')}`)

if (!process.argv.includes('--verify')) {
  console.log('\n（加 --verify 会用 Word 的 PDF 解析器再独立核对一遍）')
  process.exit(0)
}

const work = mkdtempSync(join(tmpdir(), 'dsh-office-cjk-pdf-'))
const textOut = join(work, 'word-text.txt')
console.log('\n=== 第三方解析器（Word PDF 重排）独立读取 ===')
const result = spawnSync(
  'pwsh',
  ['-NoProfile', '-File', join(here, 'pdf-open-check.ps1'), '-Path', out, '-TextOut', textOut],
  { encoding: 'utf8' }
)
const stdout = `${result.stdout ?? ''}`
for (const line of stdout.split('\n')) {
  if (/REFLOW_PAGES|WORDS|TEXT_OUT|ERROR/.test(line)) console.log(`  ${line.trim()}`)
}
if (result.status !== 0) {
  console.log('❌ Word 读取失败（本机没有 Word 时属正常，本用例需要 Office）')
  rmSync(work, { recursive: true, force: true })
  process.exit(1)
}

const wordText = readFileSync(textOut, 'utf8')
const checks = [
  ['季度销售报告（中文 PDF 从零生成验证）', '中文标题'],
  ['本文件由 dsh-exp-office 从零生成', '中文正文 + 拉丁混排'],
  ['Mixed 中英混排 with ASCII, numbers 12345', '中英混排与数字'],
  ['第二页：分页、折行与文本提取都正常。', '第二页中文'],
  ['CJK-PDF-MARKER', 'ASCII 标记']
]
let failed = 0
for (const [needle, label] of checks) {
  const hit = wordText.includes(needle)
  if (!hit) failed += 1
  console.log(`  ${hit ? '✅' : '❌'} ${label}：${hit ? 'Word 独立读到' : `Word 未读到「${needle}」`}`)
}

rmSync(work, { recursive: true, force: true })
console.log(`\n结论：${failed === 0 ? '✅ 第三方解析器能完整读出中文与中英混排文本' : `❌ 有 ${failed} 项第三方解析器读不到`}`)
process.exit(failed === 0 ? 0 : 1)
