// 生成 PPTX 主题夹具：克隆一套**配色不同**的主题（accent1 改成红色）再挂到母版上。
// 这样 PowerPoint 能用 COM 读回真实颜色，证明「换主题」确实生效，而不只是文件能打开。
// 用法：node test/make-pptx-theme-fixture.mjs
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { PptxPresentation } from '../lib/pptx.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const source = path.join(here, 'fixtures', 'deck.pptx')
const target = path.join(here, 'fixtures', 'deck-theme.pptx')
const RED = 'FF0000'

const pres = PptxPresentation.open(fs.readFileSync(source))
const before = pres.masterThemes()
const themes = pres.themes()
const other = themes.find((t) => t.part !== before[0].theme_part) ?? themes[1]
if (!other) {
  console.error('样本里只有一个主题，换不出效果。')
  process.exit(1)
}

// 把「源主题」的 accent1 改成红色：换完主题后 PowerPoint 读回的颜色应当随之改变
const themeXml = pres.pkg.readText(other.part)
const patched = themeXml.replace(/(<a:accent1>\s*<a:srgbClr val=")[0-9A-Fa-f]{6}(")/, `$1${RED}$2`)
if (patched === themeXml) {
  console.error('没能改写 accent1，无法验证换色效果。')
  process.exit(1)
}
pres.pkg.write(other.part, patched)

const result = pres.setTheme({ master: 0, theme: other.index })
fs.writeFileSync(target, pres.save())

const reopened = PptxPresentation.open(fs.readFileSync(target))
console.log(`FIXTURE: ${target}`)
console.log(`THEMES: ${themes.length}  名称: ${themes.map((t) => t.name).join(' / ')}`)
console.log(`MASTER0: ${before[0].theme_part} → ${reopened.masterThemes()[0].theme_part}（源主题 ${other.part}，accent1=#${RED}）`)
console.log(`REUSED_PART: ${result.reused_part}  VALID: ${reopened.validate().valid}`)
