// 生成 DOCX 书签夹具：在标题段与正文段各插一个书签。
// 用法：node test/make-docx-bookmark-fixture.mjs
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { DocxDocument, buildBasicDocx } from '../lib/docx.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const target = path.join(here, 'fixtures', 'report-bookmarks.docx')

const doc = DocxDocument.open(buildBasicDocx({ title: '书签验证文档' }))
const a = doc.insertBookmark({ paragraph: 0, name: 'title_mark' })
const b = doc.insertBookmark({ paragraph: 1, name: '正文_1' })

fs.writeFileSync(target, doc.save())
const reopened = DocxDocument.open(fs.readFileSync(target))
console.log(`FIXTURE: ${target}`)
console.log(`BOOKMARKS: ${reopened.bookmarks().length}`)
for (const bm of reopened.bookmarks()) console.log(`  ${bm.name} id=${bm.id} para=${bm.paragraph} end=${bm.has_end}`)
console.log(`IDS: ${a.id}, ${b.id}`)
console.log(`VALID: ${reopened.validate().valid}`)
