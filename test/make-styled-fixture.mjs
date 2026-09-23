/**
 * 生成一个「样式全覆盖」样本，用于真实 Excel/WPS 验证样式写入是否被接受。
 * 运行：node test/make-styled-fixture.mjs
 */
import { writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Workbook } from '../lib/xlsx.js'
import { buildEmptyWorkbook } from '../lib/tools.js'

const here = dirname(fileURLToPath(import.meta.url))

const wb = Workbook.open(buildEmptyWorkbook('样式测试'))
wb.writeCells('样式测试', [
  { ref: 'A1', value: '加粗红字14号' },
  { ref: 'B1', value: 1234.5 },
  { ref: 'C1', value: '黄底居中' },
  { ref: 'D1', value: 3.14159 },
  { ref: 'E1', value: '斜体下划线' },
  { ref: 'A2', value: '普通文本对照' }
])

wb.applyCellStyle('样式测试', [
  { ref: 'A1', bold: true, fontColor: '#FF0000', fontSize: 14 },
  { ref: 'B1', numberFormat: '#,##0.00' },
  { ref: 'C1', fillColor: '#FFFF00', horizontal: 'center' },
  { ref: 'D1', numberFormat: '0.0000"元"' },
  { ref: 'E1', italic: true, underline: true, fontName: 'Arial' }
])

const out = join(here, 'fixtures', 'styled.xlsx')
writeFileSync(out, wb.save())
console.log(`样式样本已生成：${out}`)
