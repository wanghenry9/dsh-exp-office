/**
 * 生成「含图片」的 DOCX 样本，供真实 Word 验证。
 * 运行：node test/make-image-fixture.mjs
 */
import zlib from 'node:zlib'
import { writeFileSync, mkdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { DocxDocument, buildBasicDocx } from '../lib/docx.js'

/**
 * 生成一张合法的 RGB PNG。
 * @param {number} width - 宽。
 * @param {number} height - 高。
 * @param {number[]} rgb - 颜色。
 * @returns {Buffer} PNG 字节。
 */
function makePng(width, height, rgb = [31, 119, 180]) {
  const raw = Buffer.alloc((width * 3 + 1) * height)
  let o = 0
  for (let y = 0; y < height; y += 1) {
    raw[o] = 0
    o += 1
    for (let x = 0; x < width; x += 1) {
      raw[o] = rgb[0]
      raw[o + 1] = rgb[1]
      raw[o + 2] = rgb[2]
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

const here = dirname(fileURLToPath(import.meta.url))
const fixtureDir = join(here, 'fixtures')
mkdirSync(fixtureDir, { recursive: true })

const wb = DocxDocument.open(buildBasicDocx({ title: '季度销售报告' }))
// 320×160 的图，显示宽度 240 px（按比例缩小到 240×120）
const result = wb.insertImage({
  after: 0,
  data: makePng(320, 160),
  extension: 'png',
  widthPx: 240,
  altText: '季度销售额柱状图占位图'
})
const out = join(fixtureDir, 'report-image.docx')
writeFileSync(out, wb.save())
console.log(`已生成：${out}`)
console.log(`媒体部件 ${result.media_part}｜关系 ${result.relationship_id}｜显示 ${result.display_size_px.width}×${result.display_size_px.height} px`)
