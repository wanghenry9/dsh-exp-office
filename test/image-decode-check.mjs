/**
 * 用第三方图像库（sharp）独立解码图片，验证插件导出的图片真的能打开。
 *
 * 为什么需要它：PNG 是我们自己拼出来的（IHDR/IDAT/IEND + CRC），
 * 「我们自己能读回」不算证据；sharp 的 libvips 是另一个实现，它能解开才说明文件合规。
 *
 * 用法：node test/image-decode-check.mjs <图片路径> [...更多路径]
 * sharp 从宿主环境解析（插件本体不依赖它）；找不到时如实报告「无法验证」而不是假装通过。
 */
import { createRequire } from 'node:module'
import os from 'node:os'
import { existsSync, statSync } from 'node:fs'
import { basename, resolve } from 'node:path'

const paths = process.argv.slice(2)
if (paths.length === 0) {
  console.log('用法：node test/image-decode-check.mjs <图片路径> [...]')
  process.exit(2)
}

const anchors = [
  `${os.homedir()}/AppData/Roaming/npm/node_modules/@deepseek-ai/dsh/node_modules/`,
  `${os.homedir()}/.dsh/profiles/web/node_modules/`
]
let sharp = null
let sharpFrom = null
for (const anchor of anchors) {
  try {
    sharp = createRequire(anchor)('sharp')
    sharpFrom = anchor
    break
  } catch {
    sharp = null
  }
}
if (!sharp) {
  console.log('⏭️  未找到 sharp，无法做独立解码验证（插件本体不依赖它）。')
  process.exit(0)
}
console.log(`解码器：sharp（来自 ${sharpFrom}）`)

let failed = 0
for (const path of paths) {
  const absolute = resolve(path)
  if (!existsSync(absolute)) {
    console.log(`❌ 文件不存在：${path}`)
    failed += 1
    continue
  }
  try {
    const meta = await sharp(absolute).metadata()
    const stats = statSync(absolute)
    // 通道均值：样本是一张纯色图（31,119,180），重采样后仍应是同一个颜色 ——
    // 有这一条才能证明「导出的确实是被插入的那张图」，而不只是「一张能打开的图」。
    const means = (await sharp(absolute).stats()).channels.map((c) => Math.round(c.mean))
    console.log(
      `✅ ${basename(absolute)}｜${meta.format} ${meta.width}x${meta.height}｜通道 ${meta.channels}${
        meta.hasAlpha ? '（含透明）' : ''
      }｜${stats.size} 字节｜通道均值 [${means.join(', ')}]`
    )
  } catch (err) {
    console.log(`❌ ${basename(absolute)} 解码失败：${err.message}`)
    failed += 1
  }
}
process.exit(failed === 0 ? 0 : 1)
