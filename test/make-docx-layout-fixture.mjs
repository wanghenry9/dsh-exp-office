// 生成 DOCX 版式夹具：A3 横向 + 自定义页边距 + 页脚居中页码域。
// 用法：node test/make-docx-layout-fixture.mjs
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { DocxDocument, buildBasicDocx } from '../lib/docx.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const target = path.join(here, 'fixtures', 'report-layout.docx')

const doc = DocxDocument.open(buildBasicDocx({ title: '版式验证报告' }))
const layout = doc.setPageLayout({ paper: 'A3', orientation: 'landscape' })
const margins = doc.setMargins({ topCm: 2, bottomCm: 2, leftCm: 2.5, rightCm: 2.5, headerCm: 1.2, footerCm: 1.2 })
const page = doc.insertPageNumber({ position: 'footer', align: 'center', prefix: '第 ', suffix: ' 页' })

fs.writeFileSync(target, doc.save())
const reopened = DocxDocument.open(fs.readFileSync(target))
console.log(`FIXTURE: ${target}`)
console.log(`PAGE: ${layout.to.width_cm} x ${layout.to.height_cm} cm / ${layout.to.orientation}`)
console.log(`MARGINS: top=${margins.to.top_cm} left=${margins.to.left_cm} footer=${margins.to.footer_cm} cm`)
console.log(`PAGE_NUMBER_PART: ${page.part}`)
console.log(`VALID: ${reopened.validate().valid}`)
console.log(`SECTPR: ${JSON.stringify(reopened.structure().section_properties)}`)
