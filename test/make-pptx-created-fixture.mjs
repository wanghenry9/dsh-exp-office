// 从零生成演示文稿夹具，并在它上面继续编辑（证明是「活文件」而不是死文件）。
// 用法：node test/make-pptx-created-fixture.mjs
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { PptxPresentation } from '../lib/pptx.js'
import { buildBlankPptx } from '../lib/pptx-template.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const target = path.join(here, 'fixtures', 'created-deck.pptx')
const editedTarget = path.join(here, 'fixtures', 'created-deck-edited.pptx')

const built = buildBlankPptx({ title: '从零生成的演示文稿', subtitle: '由 dsh-exp-office 生成（无模板文件）' })
fs.writeFileSync(target, built.bytes)

const created = PptxPresentation.open(built.bytes)
const structure = created.structure()
console.log(`FIXTURE: ${target}`)
console.log(`BYTES: ${built.bytes.length}  SLIDES: ${structure.slide_count}  尺寸: ${structure.slide_size.aspect}`)
console.log(`MASTERS: ${structure.masters}  LAYOUTS: ${structure.layouts}  THEMES: ${structure.themes}`)
console.log(`THEME_NAME: ${created.themes()[0]?.name}  LAYOUT: ${created.layouts()[0]?.name} (${created.layouts()[0]?.type})`)
console.log(`PLACEHOLDERS: ${created.readSlide(0).shapes.map((s) => `${s.name}=${JSON.stringify(s.text)}`).join(' ')}`)
console.log(`VALID: ${created.validate().valid}`)

// 在它上面继续编辑：新增页 → 文本框 → 表格 → 图表 → 换版式
const slide = created.addSlide({ layoutOf: 0 })
created.addTextBox({ index: slide.index, text: '这一页是后加的', leftPx: 80, topPx: 60, widthPx: 600, heightPx: 80, fontSizePt: 28, bold: true })
created.addSlideTable({ index: slide.index, rows: [['项目', '状态'], ['主题生成', '完成']], topPx: 200, widthPx: 600 })
created.addSlideChart({
  index: slide.index,
  type: 'column',
  title: '编辑后的图表',
  categories: ['一', '二', '三'],
  series: [{ name: '数值', values: [3, 1, 2] }],
  topPx: 360,
  widthPx: 600,
  heightPx: 240
})
created.setSlideLayout({ index: slide.index, layout: 0 })
fs.writeFileSync(editedTarget, created.save())

const reopened = PptxPresentation.open(fs.readFileSync(editedTarget))
console.log(`FIXTURE: ${editedTarget}`)
console.log(`EDITED: 页数=${reopened.structure().slide_count} 图表=${reopened.structure().charts} 形状=${reopened.readSlide(1).shapes.length}`)
console.log(`EDITED_VALID: ${reopened.validate().valid}`)
