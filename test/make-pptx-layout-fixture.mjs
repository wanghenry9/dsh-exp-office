// 生成 PPTX 版式夹具：把第 1 张幻灯片换成「空白」版式，并把第 2 张换成「仅标题」。
// 用法：node test/make-pptx-layout-fixture.mjs
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { PptxPresentation } from '../lib/pptx.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const source = path.join(here, 'fixtures', 'deck.pptx')
const target = path.join(here, 'fixtures', 'deck-layout.pptx')

const pres = PptxPresentation.open(fs.readFileSync(source))
const before = pres.slideLayouts()
const blank = pres.layouts().find((l) => l.name === '空白')
const titleOnly = pres.layouts().find((l) => l.name === '仅标题')
if (!blank || !titleOnly) {
  console.error('样本里没有「空白」/「仅标题」版式，无法生成夹具。')
  process.exit(1)
}
const first = pres.setSlideLayout({ index: 0, layout: blank.index })
const second = pres.setSlideLayout({ index: 1, layout: titleOnly.index })

fs.writeFileSync(target, pres.save())
const reopened = PptxPresentation.open(fs.readFileSync(target))
const after = reopened.slideLayouts()

console.log(`FIXTURE: ${target}`)
console.log(`LAYOUTS: ${reopened.layouts().length}`)
console.log(`SLIDE1: ${before[0].layout_name} → ${after[0].layout_name}`)
console.log(`SLIDE2: ${before[1].layout_name} → ${after[1].layout_name}`)
console.log(`REL_TARGET: ${first.to.split('/').pop()} / ${second.to.split('/').pop()}`)
console.log(`VALID: ${reopened.validate().valid}`)
