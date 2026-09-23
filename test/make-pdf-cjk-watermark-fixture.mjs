/**
 * 中文水印 / 页码样本：用「叠加式写入 + 嵌入字体子集」给真实 PDF 加中文水印，
 * 并用**第三方解析器**（Word 的 PDF 重排）独立核对「原文本还在、水印中文也读得到」。
 *
 * 为什么必须用第三方解析器：自读只证明自己写的自己认。这一步还顺手证明了另一件事 ——
 * **Word 的重排对「整份文档都盖着居中大水印」会退化**（实测 3 页全覆盖时它只吐出 31 个词、
 * 且把正文当成页眉页脚以外的 story），而「贴边（bottom）水印」它能完整读出 66 个词
 * 并把水印文字放进 footer story。所以这里用 bottom 位置做验收，并把这条观察写进 README。
 *
 * 用法：
 *   node test/make-pdf-cjk-watermark-fixture.mjs            # 只生成 + 自读
 *   node test/make-pdf-cjk-watermark-fixture.mjs --verify    # 再用 Word 独立核对
 */
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { spawnSync } from 'node:child_process'
import { PdfDocument } from '../lib/pdf.js'

const here = dirname(fileURLToPath(import.meta.url))
const outDir = join(here, 'fixtures')
mkdirSync(outDir, { recursive: true })

const source = join(outDir, 'sample-multipage.pdf')
if (!existsSync(source)) {
  console.log(`⏭️  跳过：缺少 ${source}（先运行 node test/make-pdf-fixture.mjs）。`)
  process.exit(0)
}

const original = readFileSync(source)
const doc = PdfDocument.open(original)
const pages = doc.pages().length
const overlay = doc.addTextOverlay({
  text: '机密 · 中文水印 {page}/{total}',
  position: 'bottom',
  fontSize: 10,
  gray: 0.3
})
const saved = doc.save()
const out = join(outDir, 'sample-multipage-cjk-watermark.pdf')
writeFileSync(out, saved)

const back = PdfDocument.open(saved)
const text = back.extractText().text
console.log(`已生成：${out}`)
console.log(`叠加：${pages} 页｜嵌入字体 ${overlay.embedded_font ? `${overlay.embedded_font.name}（${overlay.embedded_font.characters} 字符 / ${overlay.embedded_font.glyphs} 字形）` : '无'}`)
console.log(`与原文件对比：原 ${original.length} 字节 → 新 ${saved.length} 字节（叠加式：原字节应作为前缀保持不变）`)
console.log(`原字节前缀未变：${saved.subarray(0, original.length).equals(original)}`)
console.log(`自读页数 ${back.pages().length}｜校验 ${back.validate().valid ? '通过' : '失败'}｜水印中文可提取：${text.includes('机密 · 中文水印')}`)

if (!process.argv.includes('--verify')) {
  console.log('\n（加 --verify 会用 Word 的 PDF 解析器再独立核对一遍）')
  process.exit(0)
}

const work = mkdtempSync(join(tmpdir(), 'dsh-office-cjk-wm-'))
const textOut = join(work, 'word.txt')
console.log('\n=== 第三方解析器（Word PDF 重排）独立读取 ===')
const result = spawnSync('pwsh', ['-NoProfile', '-File', join(here, 'pdf-open-check.ps1'), '-Path', out, '-TextOut', textOut], {
  encoding: 'utf8'
})
for (const line of `${result.stdout ?? ''}`.split('\n')) {
  if (/REFLOW_PAGES|WORDS|TEXT_OUT|ERROR/.test(line)) console.log(`  ${line.trim()}`)
}
if (result.status !== 0 || !existsSync(textOut)) {
  console.log('❌ Word 读取失败（本机没有 Word 时属正常，本用例需要 Office）')
  rmSync(work, { recursive: true, force: true })
  process.exit(1)
}

const body = readFileSync(textOut, 'utf8')
const stories = existsSync(`${textOut}.stories`) ? readFileSync(`${textOut}.stories`, 'utf8') : ''
const checks = [
  ['原始正文仍在（第 1 页标题）', body.includes('季度业务报告')],
  ['原始正文仍在（表格文字）', body.includes('测试数') && body.includes('XLSX')],
  ['中文水印被独立读到（footer story）', stories.includes('机密 · 中文水印 1/')],
  // 逐页编号（1/3、2/3、3/3）由自家读取器逐页提取验证（见 test/pdf.test.mjs 的「中文页码」用例）；
  // Word 的重排会把多页 footer 折叠成一个 story，只保留第一页的写法，所以这里不拿它判逐页。
  ['水印占位符已替换成具体页码', /机密 · 中文水印 1\/\d/.test(stories)]
]
let failed = 0
for (const [label, ok] of checks) {
  if (!ok) failed += 1
  console.log(`  ${ok ? '✅' : '❌'} ${label}`)
}
rmSync(work, { recursive: true, force: true })
console.log(`\n结论：${failed === 0 ? '✅ 原内容未受影响，且第三方解析器读到了我们叠加的中文水印' : `❌ 有 ${failed} 项未通过`}`)
process.exit(failed === 0 ? 0 : 1)
void copyFileSync
