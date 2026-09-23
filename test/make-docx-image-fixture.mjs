/**
 * 生成「图片缩放/删除」的验证样本，供真实 Word 核对尺寸与形状数。
 *
 * 产物：
 *   test/fixtures/report-image-resized.docx —— 图片缩放到 200×90 像素（Word 应读到 150×67.5pt）
 *   test/fixtures/report-image-deleted.docx —— 图片连同空段落一起删除（Word 应读到 0 个内联形状）
 *
 * 运行：node test/make-docx-image-fixture.mjs
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { DocxDocument } from '../lib/docx.js'

const here = dirname(fileURLToPath(import.meta.url))
const source = join(here, 'fixtures', 'report-image.docx')
const original = readFileSync(source)

const resizedDoc = DocxDocument.open(original)
const resized = resizedDoc.resizeImage({ index: 0, widthPx: 200, heightPx: 90 })
const resizedBytes = resizedDoc.save()
writeFileSync(join(here, 'fixtures', 'report-image-resized.docx'), resizedBytes)

const deletedDoc = DocxDocument.open(original)
const deleted = deletedDoc.deleteImage({ index: 0 })
const deletedBytes = deletedDoc.save()
writeFileSync(join(here, 'fixtures', 'report-image-deleted.docx'), deletedBytes)

console.log(
  `已生成 report-image-resized.docx（${resizedBytes.length} 字节，${resized.from.width}×${resized.from.height} → ${resized.to.width}×${resized.to.height}px）`
)
console.log(
  `已生成 report-image-deleted.docx（${deletedBytes.length} 字节，删除整段：${deleted.removed_paragraph}，图片数 ${DocxDocument.open(deletedBytes).images().length}）`
)
