/**
 * 生成「带 Excel 表格对象」的样本：用插件自己的 API 在 sales.xlsx 上建表格。
 *
 * 产物 test/fixtures/sales-table.xlsx 会被 test/office-open-check.ps1 拿去用真实 Excel 打开，
 * 验证 ListObject 被 Excel 认出来（Excel 对表格 XML 很严格，顺序/接线错一处就会要求修复文件）。
 *
 * 运行：node test/make-table-fixture.mjs
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Workbook } from '../lib/xlsx.js'

const here = dirname(fileURLToPath(import.meta.url))
const source = join(here, 'fixtures', 'sales.xlsx')
const target = join(here, 'fixtures', 'sales-table.xlsx')

const wb = Workbook.open(readFileSync(source))
const info = wb.createTable('销售数据', { ref: 'A1:D5', name: '销售表', style: 'TableStyleMedium9' })
const bytes = wb.save()
writeFileSync(target, bytes)
console.log(
  `已生成：test/fixtures/sales-table.xlsx（${bytes.length} 字节）表格「${info.display_name}」${info.ref}，${info.columns.length} 列，部件 ${info.part}`
)
