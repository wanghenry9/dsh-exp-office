/**
 * 生成「插件建的表格 + 套用样式」的样本，供真实 Word 验证。
 *
 * 产物 test/fixtures/report-tables-built.docx 会被 test/word-open-check.ps1 打开，
 * 验证：表格数量、行列数、样式名（Word 读 table.Style）与单元格内容。
 *
 * 运行：node test/make-docx-table-fixture.mjs
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { DocxDocument } from '../lib/docx.js'

const here = dirname(fileURLToPath(import.meta.url))
const source = join(here, 'fixtures', 'report.docx')
const target = join(here, 'fixtures', 'report-tables-built.docx')

const doc = DocxDocument.open(readFileSync(source))
const created = doc.createTable({
  after: 0,
  rows: [
    ['产品', '数量', '金额'],
    ['笔记本', '3', '17998.5'],
    ['显示器', '10', '12990']
  ],
  header: true,
  styleId: 'TableGrid'
})
const styled = doc.setTableStyle({ table: 0, styleId: 'TableGrid', firstRow: true, bandedRows: true })
const bytes = doc.save()
writeFileSync(target, bytes)
console.log(
  `已生成：test/fixtures/report-tables-built.docx（${bytes.length} 字节）新建 ${created.rows}×${created.columns} 表格，样式 ${styled.style_id}`
)
