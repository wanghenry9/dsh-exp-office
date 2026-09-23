// 生成 DOCX 导出夹具：一份逐字节一致的 docx 副本 + 一份带 BOM 的纯文本。
// 用法：node test/make-docx-export-fixture.mjs
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { DocxDocument } from '../lib/docx.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const source = path.join(here, 'fixtures', 'report.docx')
const target = path.join(here, 'fixtures', 'report-exported.docx')
const textTarget = path.join(here, 'fixtures', 'report-exported.txt')

const sourceBytes = fs.readFileSync(source)
// 导出路径与工具实现一致：docx 是原样复制，text 是带 BOM 的 UTF-8
fs.writeFileSync(target, sourceBytes)
const { text } = DocxDocument.open(sourceBytes).text()
fs.writeFileSync(textTarget, Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(text, 'utf8')]))

const copyBytes = fs.readFileSync(target)
console.log(`FIXTURE: ${target}`)
console.log(`FIXTURE: ${textTarget}`)
console.log(`BYTE_IDENTICAL: ${copyBytes.equals(sourceBytes)}  字节: ${copyBytes.length}`)
console.log(`TEXT_CHARS: ${text.length}  首行: ${text.split('\n')[0]}`)
