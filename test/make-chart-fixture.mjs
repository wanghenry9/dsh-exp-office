/**
 * 生成「带图表」的样本：用插件自己的 API 在 sales.xlsx 上建一个柱状图。
 *
 * 产物 test/fixtures/sales-chart.xlsx 会被 test/office-open-check.ps1 / wps-open-check.ps1
 * 拿去用真实 Excel 与 WPS 表格打开，验证图表对象被认出来（Excel 对图表 XML 的顺序极敏感，
 * 错一处就会要求修复文件）。
 *
 * 运行：node test/make-chart-fixture.mjs
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Workbook } from '../lib/xlsx.js'

const here = dirname(fileURLToPath(import.meta.url))
const source = join(here, 'fixtures', 'sales.xlsx')
const target = join(here, 'fixtures', 'sales-chart.xlsx')

const wb = Workbook.open(readFileSync(source))
const info = wb.createChart('销售数据', {
  type: 'column',
  title: '各产品金额',
  categories: 'A2:A4',
  series: [{ values: 'D2:D4', nameRef: 'D1' }],
  anchor: 'F2',
  widthPx: 480,
  heightPx: 300
})
const bytes = wb.save()
writeFileSync(target, bytes)
console.log(
  `已生成：test/fixtures/sales-chart.xlsx（${bytes.length} 字节）${info.type} 图表，锚点 ${info.anchor}，系列 ${info.series_count} 个，部件 ${info.chart_part}`
)
