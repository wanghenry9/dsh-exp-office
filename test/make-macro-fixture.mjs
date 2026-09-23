/**
 * 生成「宏启用工作簿」测试样本（.xlsm）。
 *
 * 关于这个样本的诚实边界：
 *   - 本机构建 `vbaProject.bin` 有两条路：Excel COM 写 VBA 需要打开
 *     「信任对 VBA 工程对象模型的访问」（默认关闭，且属于改动用户的 Office 安全设置，
 *     不擅自开启）；系统里也没有现成的宏文件样本可复用。
 *   - 因此这里**合成**一个带 OLE 复合文档头的 `vbaProject.bin` 部件。
 *     它能证明：① 宏检测路径识别得出「这是宏文件」；
 *     ② 编辑普通单元格后该部件逐字节不变（宏按原字节保留）。
 *     它**不能**证明：Excel 的 VBA 编辑器仍能运行其中的宏 —— 那需要一个真实的
 *     宏工程，属于需要用户授权的验证项。
 *
 * 运行：node test/make-macro-fixture.mjs
 */
import { writeFileSync, mkdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { ZipPackage, XmlDoc, findAll, attr } from '../lib/ooxml.js'
import { buildEmptyWorkbook } from '../lib/tools.js'

/** OLE 复合文档签名：让合成部件在结构上可被识别为 VBA 工程的容器形态。 */
const OLE_SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])

/**
 * 合成一个具备 OLE 头的 vbaProject.bin。
 * @param {number} [size] - 总字节数。
 * @returns {Buffer} 部件字节。
 */
export function syntheticVbaProject(size = 4096) {
  const buf = Buffer.alloc(size)
  OLE_SIGNATURE.copy(buf, 0)
  // 填充可辨识的内容，便于在失败时定位
  buf.write('DSH-EXP-OFFICE synthetic vbaProject for byte-preservation test', 8, 'latin1')
  for (let i = 64; i < size; i += 1) buf[i] = (i * 31) % 251
  return buf
}

/**
 * 把一份 .xlsx 升级为带宏部件与相应声明的 .xlsm。
 * @param {Buffer} xlsxBytes - 源工作簿字节。
 * @param {Buffer} vbaBytes - vbaProject.bin 字节。
 * @returns {Buffer} .xlsm 字节。
 */
export function toMacroEnabled(xlsxBytes, vbaBytes) {
  const pkg = ZipPackage.open(xlsxBytes)
  pkg.write('xl/vbaProject.bin', vbaBytes)

  // 内容类型：主部件改为 macroEnabled，并声明 bin 的默认类型
  const ctDoc = XmlDoc.parse(pkg.readText('[Content_Types].xml'))
  const ctRoot = ctDoc.root.children.find((c) => c.type === 'element')
  ctDoc.appendChild(ctRoot, '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>')
  for (const node of findAll(ctRoot, 'Override')) {
    if (attr(node, 'PartName') === '/xl/workbook.xml') {
      ctDoc.setAttr(node, 'ContentType', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
    }
  }
  pkg.write('[Content_Types].xml', ctDoc.toString())

  // 关系：workbook → vbaProject
  const relsDoc = XmlDoc.parse(pkg.readText('xl/_rels/workbook.xml.rels'))
  const relsRoot = relsDoc.root.children.find((c) => c.type === 'element')
  let max = 0
  for (const rel of findAll(relsRoot, 'Relationship')) {
    const m = /^rId(\d+)$/.exec(attr(rel, 'Id') ?? '')
    if (m) max = Math.max(max, Number(m[1]))
  }
  relsDoc.appendChild(
    relsRoot,
    `<Relationship Id="rId${max + 1}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>`
  )
  pkg.write('xl/_rels/workbook.xml.rels', relsDoc.toString())

  return pkg.toBuffer()
}

const here = dirname(fileURLToPath(import.meta.url))
const fixtureDir = join(here, 'fixtures')
mkdirSync(fixtureDir, { recursive: true })

const vba = syntheticVbaProject()
const xlsm = toMacroEnabled(buildEmptyWorkbook('数据'), vba)
const out = join(fixtureDir, 'macro.xlsm')
writeFileSync(out, xlsm)

const rawXlsx = join(fixtureDir, 'macro-source.xlsx')
writeFileSync(rawXlsx, buildEmptyWorkbook('数据'))

console.log(`已生成：${out}（${xlsm.length} 字节，含合成 vbaProject.bin ${vba.length} 字节）`)
