/**
 * 逐格读数：把插件读到的值与 Office/WPS 读到的值对照（E5 交叉验证的一半）。
 *
 * 用法：
 *   node test/read-cells-check.mjs <工作簿> <工作表> A1 B2 C2 J50000
 *
 * 输出 `REF: VALUE`，可以直接与 `test/office-open-check.ps1 -CellRefs` 的输出逐行对照。
 */
import { readFileSync } from 'node:fs'
import { Workbook } from '../lib/xlsx.js'

const [file, sheet, ...refs] = process.argv.slice(2)
if (!file || !sheet || refs.length === 0) {
  console.error('用法：node test/read-cells-check.mjs <工作簿> <工作表> <单元格引用...>')
  process.exit(2)
}

const wb = Workbook.open(readFileSync(file))
console.log(`FILE: ${file}`)
console.log(`SHEET: ${sheet}`)
console.log(`REF_COUNT: ${refs.length}`)
for (const ref of refs) {
  const cell = wb.readRange(sheet, `${ref}:${ref}`).rows[0]?.[0]
  const shown = cell === null || cell === undefined ? '' : String(cell.value)
  console.log(`${ref}: ${shown}${cell?.formula ? `   [公式 ${cell.formula}]` : ''}`)
}
