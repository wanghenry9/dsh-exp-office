/**
 * 大文件样本生成器：直接拼工作表 XML，不经过 DOM。
 *
 * 为什么不用 `wb.writeCells()` 生成大样本：
 * 那条路每写一批都要建/改整表 DOM，20 万单元格就已经要 600 MB 级内存，
 * 100 万单元格直接 OOM（本机 Node 默认堆约 2 GB）。生成器只拼字符串，
 * 因此 100 万单元格也能在几十 MB 内存里造出来。
 *
 * 用法：
 *   node test/make-xlsx-large-fixture.mjs                 # 50 万单元格（5 万行 × 10 列）
 *   node test/make-xlsx-large-fixture.mjs --rows 50000 --cols 20 --wide   # 50 MB 级
 *   node test/make-xlsx-large-fixture.mjs --rows 1000000 --cols 1         # 100 万单元格
 *   node test/make-xlsx-large-fixture.mjs --out D:\tmp\big.xlsx
 *
 * 产出默认落在 `test/fixtures/large-<行>x<列>.xlsx`，已在 .gitignore 里排除。
 */
import { mkdirSync, statSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { ZipPackage } from '../lib/ooxml.js'

const here = dirname(fileURLToPath(import.meta.url))

/**
 * 取命令行参数值。
 * @param {string} name - 参数名。
 * @param {string|null} fallback - 默认值。
 * @returns {string|null} 参数值。
 */
function arg(name, fallback = null) {
  const index = process.argv.indexOf(`--${name}`)
  return index === -1 ? fallback : process.argv[index + 1]
}

const rows = Number(arg('rows', '50000'))
const cols = Number(arg('cols', '10'))
const wide = process.argv.includes('--wide')
const random = process.argv.includes('--random')
const padding = Number(arg('padding', wide ? '180' : '0'))
const out = arg('out', join(here, 'fixtures', `large-${rows}x${cols}${wide ? '-wide' : ''}${random ? '-rand' : ''}.xlsx`))

if (!Number.isInteger(rows) || rows < 1 || !Number.isInteger(cols) || cols < 1) {
  console.error('用法：node test/make-xlsx-large-fixture.mjs [--rows N] [--cols N] [--wide] [--random] [--padding N] [--out 路径]')
  process.exit(2)
}

/**
 * 1 基列号转列字母。
 * @param {number} index - 1 基列号。
 * @returns {string} 列字母。
 */
function col(index) {
  let n = index
  let out = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(65 + rem) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

/** 共享字符串表内容（少量字符串被大量单元格复用，才需要 t="s"）。 */
const palette = ['甲', '乙', '丙', '丁', '戊', '己', '庚', '辛']

/**
 * 可复现的伪随机数发生器（mulberry32）。
 *
 * `--random` 用它生成长文本：随机内容几乎压不动，才能造出真正的 50 MB 级文件；
 * 固定种子保证同一条命令每次生成同样的文件。
 *
 * @param {number} seed - 种子。
 * @returns {() => number} `[0, 1)` 上的伪随机数。
 */
function mulberry32(seed) {
  let a = seed >>> 0
  return () => {
    a = (a + 0x6d2b79f5) >>> 0
    let t = a
    t = Math.imul(t ^ (t >>> 15), t | 1)
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

const nextRandom = mulberry32(20260923)
const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'

/**
 * 生成一个单元格的长文本。
 * @param {number} r - 行号（混进内容里，保证各行互不相同）。
 * @returns {string} 文本。
 */
function cellText(r) {
  if (padding <= 0) return ''
  if (!random) return `${'长文本'.repeat(Math.ceil(padding / 3)).slice(0, padding)}${r}`
  let text = ''
  for (let i = 0; i < padding; i += 1) text += alphabet[Math.floor(nextRandom() * alphabet.length)]
  return `${text}${r}`
}

const started = Date.now()
const parts = []
for (let r = 1; r <= rows; r += 1) {
  const cells = []
  for (let c = 1; c <= cols; c += 1) {
    const ref = `${col(c)}${r}`
    if (c === 1) cells.push(`<c r="${ref}" t="s"><v>${r % palette.length}</v></c>`)
    else if (padding > 0 && c === cols) cells.push(`<c r="${ref}" t="inlineStr"><is><t>${cellText(r)}</t></is></c>`)
    else if (c % 3 === 0) cells.push(`<c r="${ref}"><f>${col(c)}${Math.max(1, r - 1)}*2</f><v>${(r * c) % 100000}</v></c>`)
    else cells.push(`<c r="${ref}"><v>${(r * c) % 100000}</v></c>`)
  }
  parts.push(`<row r="${r}">${cells.join('')}</row>`)
}
const sheetData = parts.join('')

const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { maxEntries: 64, maxEntryBytes: 1 << 30, maxTotalBytes: 1 << 31, maxRatio: 1000 })
pkg.write(
  '[Content_Types].xml',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`
)
pkg.write(
  '_rels/.rels',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
)
pkg.write(
  'xl/workbook.xml',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="大数据" sheetId="1" r:id="rId1"/></sheets></workbook>`
)
pkg.write(
  'xl/_rels/workbook.xml.rels',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`
)
pkg.write(
  'xl/worksheets/sheet1.xml',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:${col(cols)}${rows}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>${sheetData}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`
)
pkg.write(
  'xl/styles.xml',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`
)
pkg.write(
  'xl/sharedStrings.xml',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${rows}" uniqueCount="${palette.length}">${palette.map((s) => `<si><t>${s}</t></si>`).join('')}</sst>`
)

const buffer = pkg.toBuffer()
mkdirSync(dirname(out), { recursive: true })
writeFileSync(out, buffer)
const size = statSync(out).size
console.log(`已生成：${out}`)
console.log(
  `规模：${rows} 行 × ${cols} 列 = ${(rows * cols).toLocaleString('en-US')} 个单元格｜文件 ${(size / 1048576).toFixed(1)} MB｜耗时 ${((Date.now() - started) / 1000).toFixed(1)}s｜RSS ${(process.memoryUsage().rss / 1048576).toFixed(0)} MB`
)
