// 生成 DOCX 图片环绕夹具：把图片从行内改成四周型（square），并给一个偏移。
// 用法：node test/make-docx-wrap-fixture.mjs
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { DocxDocument } from '../lib/docx.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const source = path.join(here, 'fixtures', 'report-image.docx')
const target = path.join(here, 'fixtures', 'report-image-wrapped.docx')

const doc = DocxDocument.open(fs.readFileSync(source))
const before = doc.images()[0]
const result = doc.setImageWrap({ index: 0, wrap: 'square', offsetXEmu: 114300, offsetYEmu: 57150 })
const back = DocxDocument.open(doc.save())
fs.writeFileSync(target, doc.save())

const after = back.images()[0]
console.log(`FIXTURE: ${target}`)
console.log(`WRAP_BEFORE: ${before.wrap} / WRAP_AFTER: ${after.wrap}`)
console.log(`SIZE: ${after.width_px}x${after.height_px} px`)
console.log(`CHANGED: ${result.changed}  VALID: ${back.validate().valid}`)
