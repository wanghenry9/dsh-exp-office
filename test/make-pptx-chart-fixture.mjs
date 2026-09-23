// 生成 PPTX 图表夹具：四种图表类型各插一张，另用一页插入双系列柱状图。
// 用法：node test/make-pptx-chart-fixture.mjs
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { PptxPresentation } from '../lib/pptx.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const source = path.join(here, 'fixtures', 'deck.pptx')
const target = path.join(here, 'fixtures', 'deck-charts.pptx')

const pres = PptxPresentation.open(fs.readFileSync(source))
const before = pres.structure().charts

const blocks = [
  { type: 'column', title: '柱状图：各产品金额', series: [{ name: '金额', values: [120, 150, 180] }] },
  { type: 'bar', title: '条形图：各产品金额', series: [{ name: '金额', values: [120, 150, 180] }] },
  { type: 'line', title: '折线图：月度趋势', series: [{ name: '趋势', values: [10, 25, 18, 32] }] },
  { type: 'pie', title: '饼图：占比', series: [{ name: '占比', values: [40, 35, 25] }] },
  {
    type: 'column',
    title: '双系列对比',
    series: [
      { name: '今年', values: [120, 150, 180] },
      { name: '去年', values: [100, 130, 160] }
    ]
  }
]

const results = []
for (const block of blocks) {
  const slide = pres.addSlide({ layoutOf: 2 }) // 「空白」版式
  const categories = block.series[0].values.map((_, i) => `${i + 1}月`)
  const result = pres.addSlideChart({
    index: slide.index,
    type: block.type,
    title: block.title,
    categories,
    series: block.series
  })
  results.push({ slide: slide.index, type: block.type, chart: result.chart_part, series: block.series.length })
}

fs.writeFileSync(target, pres.save())
const reopened = PptxPresentation.open(fs.readFileSync(target))
console.log(`FIXTURE: ${target}`)
console.log(`SLIDES: ${reopened.structure().slide_count}  CHARTS_BEFORE: ${before}  CHARTS_AFTER: ${reopened.structure().charts}`)
for (const item of results) console.log(`  slide ${item.slide}: ${item.type} → ${item.chart}（${item.series} 个系列）`)
console.log(`VALID: ${reopened.validate().valid}`)
