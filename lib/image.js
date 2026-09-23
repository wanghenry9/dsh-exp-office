/**
 * 图像编码的格式无关能力。
 *
 * 目前只有 PNG 编码：PDF 里用 FlateDecode 存的是**原始像素**，要变成图片查看器能打开的
 * 独立文件就得包一层 PNG（IHDR + IDAT + IEND，每行前置一个过滤类型字节）。
 * 放在独立模块而不是某个适配器里，因为它不属于任何一种办公文档格式。
 *
 * @module dsh-exp-office/image
 */

import zlib from 'node:zlib'

/** PNG 颜色类型。 */
export const PNG_COLOR_TYPE = Object.freeze({ GRAY: 0, RGB: 2, INDEXED: 3, GRAY_ALPHA: 4, RGBA: 6 })

/** 每种颜色类型的通道数。 */
const CHANNELS = Object.freeze({ 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 })

/**
 * 编码一张 8 位 PNG。
 *
 * 每行都写「过滤类型 0（None）」：这里只求正确，不做逐行启发式挑选 ——
 * PDF 侧的预测器已经在别处还原过了，两者是独立的两件事。
 *
 * @param {object} args - 参数。
 * @param {number} args.width - 宽（像素）。
 * @param {number} args.height - 高（像素）。
 * @param {number} args.colorType - `PNG_COLOR_TYPE` 之一。
 * @param {Buffer} args.data - 像素字节（按行连续，不含过滤字节）。
 * @param {Buffer} [args.palette] - 索引色调色板（每色 3 字节），仅索引色需要。
 * @param {Buffer} [args.alpha] - 8 位灰度透明度；给出时灰度/RGB 会合成成 GrayAlpha/RGBA。
 * @returns {Buffer} PNG 字节。
 */
export function encodePng({ width, height, colorType, data, palette = null, alpha = null }) {
  if (!CHANNELS[colorType]) throw new Error(`不支持的 PNG 颜色类型：${colorType}`)
  const expected = width * height * CHANNELS[colorType]
  if (data.length !== expected) {
    throw new Error(`像素字节长度不符：期望 ${expected}，实际 ${data.length}`)
  }
  if (alpha && alpha.length !== width * height) {
    throw new Error(`透明度字节长度不符：期望 ${width * height}，实际 ${alpha.length}`)
  }

  let pixels = data
  let outType = colorType
  if (alpha && colorType === 0) {
    const mixed = Buffer.alloc(width * height * 2)
    for (let i = 0; i < width * height; i += 1) {
      mixed[i * 2] = data[i]
      mixed[i * 2 + 1] = alpha[i]
    }
    pixels = mixed
    outType = 4
  } else if (alpha && colorType === 2) {
    const mixed = Buffer.alloc(width * height * 4)
    for (let i = 0; i < width * height; i += 1) {
      mixed[i * 4] = data[i * 3]
      mixed[i * 4 + 1] = data[i * 3 + 1]
      mixed[i * 4 + 2] = data[i * 3 + 2]
      mixed[i * 4 + 3] = alpha[i]
    }
    pixels = mixed
    outType = 6
  }

  const stride = width * CHANNELS[outType]
  const raw = Buffer.alloc((stride + 1) * height)
  for (let y = 0; y < height; y += 1) {
    raw[y * (stride + 1)] = 0
    pixels.copy(raw, y * (stride + 1) + 1, y * stride, (y + 1) * stride)
  }

  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(width, 0)
  ihdr.writeUInt32BE(height, 4)
  ihdr[8] = 8
  ihdr[9] = outType
  const chunks = [Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr)]
  if (outType === 3 && palette) chunks.push(chunk('PLTE', palette))
  chunks.push(chunk('IDAT', zlib.deflateSync(raw, { level: 6 })))
  chunks.push(chunk('IEND', Buffer.alloc(0)))
  return Buffer.concat(chunks)
}

/**
 * 组一个 PNG 块（长度 + 类型 + 数据 + CRC）。
 * @param {string} type - 块类型（4 字节 ASCII）。
 * @param {Buffer} payload - 数据。
 * @returns {Buffer} 块字节。
 */
function chunk(type, payload) {
  const length = Buffer.alloc(4)
  length.writeUInt32BE(payload.length)
  const typeBuf = Buffer.from(type, 'latin1')
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(zlib.crc32(Buffer.concat([typeBuf, payload])) >>> 0)
  return Buffer.concat([length, typeBuf, payload, crc])
}
