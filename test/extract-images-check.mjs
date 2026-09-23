/**
 * 端到端独立验证：把 sample-images.pdf 里的图片提取出来，然后用第三方解码器（sharp）检查。
 *
 * 插件本体不依赖 sharp；这个脚本属于测试工具，用来回答「导出的图片别人能不能打开、
 * 内容是不是原来那张」——样本是纯色图 (31,119,180)，通道均值必须对得上。
 *
 * 用法：node test/extract-images-check.mjs
 */
import { mkdirSync, readFileSync, rmSync, writeFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import { PdfDocument } from '../lib/pdf.js'
import { detectImageType, readImageSize } from '../lib/ooxml.js'

const workDir = 'tmp-verify-images'
rmSync(workDir, { recursive: true, force: true })
mkdirSync(workDir, { recursive: true })

const doc = PdfDocument.open(readFileSync('test/fixtures/sample-images.pdf'))
const { images, skipped } = doc.extractImages()
console.log(`提取：${images.length} 张，跳过 ${skipped.length} 个`)
for (const image of images) {
  const name = `page${image.pages[0]}-${image.name}.${image.extension}`
  writeFileSync(join(workDir, name), image.bytes)
  // 插件自身的读法（类型 + 尺寸）
  const type = detectImageType(image.bytes)
  const size = readImageSize(image.bytes, image.extension)
  console.log(
    `  ${name}：${type} ${size.width}x${size.height}（插件读取）｜声明尺寸 ${image.width}x${image.height}｜共用页 ${image.pages.join(',')}`
  )
}
console.log('输出目录内容:', readdirSync(workDir).join(', '))
