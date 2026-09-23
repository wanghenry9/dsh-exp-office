// 生成带中文便签批注的 PDF 夹具。
// 用法：node test/make-pdf-annotation-fixture.mjs
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { PdfDocument } from '../lib/pdf.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const source = path.join(here, 'fixtures', 'sample.pdf')
const target = path.join(here, 'fixtures', 'sample-annotated.pdf')

const pdf = PdfDocument.open(fs.readFileSync(source))
const before = pdf.annotations().count
const result = pdf.addAnnotation({
  page: 0,
  text: '批注：这一页需要法务复核，签字后归档。',
  author: '审阅者',
  x: 60,
  y: 700,
  open: true
})
fs.writeFileSync(target, pdf.save())

const reopened = PdfDocument.open(fs.readFileSync(target))
const list = reopened.annotations()
console.log(`FIXTURE: ${target}`)
console.log(`ANNOTATIONS_BEFORE: ${before}  AFTER: ${list.count}`)
console.log(`SUBTYPE: ${list.annotations[0].subtype}  RECT: ${list.annotations[0].rect.join(',')}`)
console.log(`CONTENTS: ${list.annotations[0].contents}`)
console.log(`AUTHOR: ${list.annotations[0].author}  FLAGS: ${list.annotations[0].flags}`)
console.log(`NON_ASCII: ${result.non_ascii}  VALID: ${reopened.validate().checks.every((c) => c.ok)}`)
console.log(`STRUCTURE_ANNOTATIONS: ${reopened.structure().annotations}`)
