/**
 * 生成 PDF 测试样本：用宿主内置 LibreOffice 引擎把 DOCX 转成 PDF。
 *
 * 这样得到的是**真实世界的 PDF**（PDF 1.7、压缩内容流、子集化 CJK 字体、
 * 带 ToUnicode CMap），比手搓一个最小 PDF 更能验证解析器。
 *
 * 运行：node test/make-pdf-fixture.mjs
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { tmpdir } from 'node:os'
import { mkdtempSync } from 'node:fs'
import { buildBasicDocx } from '../lib/docx.js'
import { convertToPdf, disposeConverter } from '../lib/convert.js'

const here = dirname(fileURLToPath(import.meta.url))
const fixtureDir = join(here, 'fixtures')
mkdirSync(fixtureDir, { recursive: true })

const work = mkdtempSync(join(tmpdir(), 'dsh-pdf-fixture-'))
const docxPath = join(work, 'source.docx')
const pdfPath = join(work, 'source.pdf')

writeFileSync(docxPath, buildBasicDocx({ title: '季度销售报告', paragraph: '本文件用于验证 PDF 解析器。' }))

const result = await convertToPdf({ inputPath: docxPath, outputPath: pdfPath })

// 第二份样本：把 4 页演示文稿（含 1 页隐藏）导出成多页 PDF，
// 用于验证页面级写操作（旋转 / 重排 / 删除）能产出真实阅读器可读的文件。
const deckPath = join(here, 'fixtures', 'deck.pptx')
const multiPath = join(work, 'multipage.pdf')
const multi = await convertToPdf({ inputPath: deckPath, outputPath: multiPath })

// 第三份样本：演示文稿里有一张插入的图片，导出后用于验证「PDF 图片提取」
// （真实引擎写出的图像流：可能是 FlateDecode + PNG 预测器，也可能是 DCTDecode）
const imageDeckPath = join(here, 'fixtures', 'deck-image.pptx')
const imagePath = join(work, 'with-image.pdf')
const withImage = await convertToPdf({ inputPath: imageDeckPath, outputPath: imagePath })
await disposeConverter()

const { readFileSync, rmSync } = await import('node:fs')
const bytes = readFileSync(pdfPath)
writeFileSync(join(fixtureDir, 'sample.pdf'), bytes)
const multiBytes = readFileSync(multiPath)
writeFileSync(join(fixtureDir, 'sample-multipage.pdf'), multiBytes)
const imageBytes = readFileSync(imagePath)
writeFileSync(join(fixtureDir, 'sample-images.pdf'), imageBytes)
rmSync(work, { recursive: true, force: true })

console.log(`已生成：test/fixtures/sample.pdf（${bytes.length} 字节，${result.pages} 页，引擎 ${result.backend}）`)
console.log(
  `已生成：test/fixtures/sample-multipage.pdf（${multiBytes.length} 字节，${multi.pages} 页，引擎 ${multi.backend}）`
)
console.log(
  `已生成：test/fixtures/sample-images.pdf（${imageBytes.length} 字节，${withImage.pages} 页，引擎 ${withImage.backend}）`
)

