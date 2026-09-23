// 生成 PPTX 幻灯片操作夹具：复制一张带备注的幻灯片，并在副本上插入文本框。
// 用法：node test/make-pptx-slide-ops-fixture.mjs
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { PptxPresentation } from '../lib/pptx.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const source = path.join(here, 'fixtures', 'deck.pptx')
const target = path.join(here, 'fixtures', 'deck-slide-ops.pptx')

const pres = PptxPresentation.open(fs.readFileSync(source))
const before = pres.structure().slide_count

// 下标 1 的幻灯片带 notesSlide 关系：复制时应当丢掉备注关系而不是共用备注部件。
const dup = pres.duplicateSlide({ index: 1 })
const box = pres.addTextBox({
  index: dup.index,
  text: '复制幻灯片验证标记\n第二行文本',
  leftPx: 96,
  topPx: 420,
  widthPx: 520,
  heightPx: 120,
  fontSizePt: 24,
  bold: true,
  align: 'center'
})

fs.writeFileSync(target, pres.save())
const reopened = PptxPresentation.open(fs.readFileSync(target))
console.log(`FIXTURE: ${target}`)
console.log(`SLIDES_BEFORE: ${before}`)
console.log(`SLIDES_AFTER: ${reopened.structure().slide_count}`)
console.log(`DUPLICATE_INDEX: ${dup.index}`)
console.log(`NOTES_DROPPED: ${dup.notes_dropped}`)
console.log(`TEXT_BOX_SHAPE_ID: ${box.shape_id}`)
console.log(`VALID: ${reopened.validate().valid}`)
