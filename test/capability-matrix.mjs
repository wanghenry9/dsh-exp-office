/**
 * 能力矩阵：**从代码生成**工具总表，避免文档与实现漂移。
 *
 * 分工：
 *   - 本脚本只负责「事实部分」——工具名/用途（读 `dsh.plugin.json`）、能力与非能力清单
 *     （读 `lib/capabilities.js`）、套件清单（读 `test/all.mjs`）、脚本清单（读 `package.json`）；
 *   - `docs/能力矩阵.md` 里的散文部分（验证矩阵、验收清单、验证缺口）由人维护，
 *     生成器只替换 `<!-- GENERATED:... -->` 标记之间的内容。
 *
 * 用法：
 *   node test/capability-matrix.mjs            # 重新生成（写入 docs/能力矩阵.md）
 *   node test/capability-matrix.mjs --check    # 只校验有没有漂移（交付前自检用；漂移退出 3）
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const root = join(here, '..')
const docPath = join(root, 'docs', '能力矩阵.md')
const checkOnly = process.argv.includes('--check')

const manifest = JSON.parse(readFileSync(join(root, 'dsh.plugin.json'), 'utf8'))
const packageJson = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'))
const capabilitiesSource = readFileSync(join(root, 'lib', 'capabilities.js'), 'utf8')
const suitesSource = readFileSync(join(root, 'test', 'all.mjs'), 'utf8')

// 工具名与用途以**运行时定义**为准（manifest 里只有名字），顺带核对两份清单是否一致
const { createOfficeTools } = await import('../lib/tools.js')
const { resolveDefineTool } = await import('../lib/define-tool.js')
const { mkdtempSync, rmSync } = await import('node:fs')
const { tmpdir } = await import('node:os')
const workspace = mkdtempSync(join(tmpdir(), 'dsh-office-matrix-'))
const { defineTool } = await resolveDefineTool({})
const definitions = createOfficeTools({ defineTool, config: { workspaceRoot: workspace } })
rmSync(workspace, { recursive: true, force: true })
const declaredNames = [...manifest.contributes.tools].sort()
const runtimeNames = definitions.map((d) => d.name).sort()
const nameDrift =
  declaredNames.length === runtimeNames.length && declaredNames.every((name, index) => name === runtimeNames[index])
    ? null
    : {
        only_in_manifest: declaredNames.filter((name) => !runtimeNames.includes(name)),
        only_in_runtime: runtimeNames.filter((name) => !declaredNames.includes(name))
      }

/**
 * 从 capabilities.js 里抽出一个导出数组的字面量项。
 * @param {string} name - 导出名（如 CAPABILITIES）。
 * @returns {string[]} 项列表。
 */
function readStringArray(name) {
  const match = new RegExp(`export const ${name} = Object.freeze\\(\\[([\\s\\S]*?)\\]\\)`).exec(capabilitiesSource)
  if (!match) throw new Error(`capabilities.js 里找不到 ${name}`)
  return [...match[1].matchAll(/'([^']+)'/g)].map((m) => m[1])
}

/**
 * 从 all.mjs 里抽出套件清单。
 * @returns {{title: string, file: string}[]} 套件。
 */
function readSuites() {
  return [...suitesSource.matchAll(/\['([^']+)',\s*'([^']+)'\]/g)].map((m) => ({ title: m[1], file: m[2] }))
}

/**
 * 把工具按名字前缀归到「格式/功能域」。
 * @param {string} name - 工具名。
 * @returns {string} 分组名。
 */
function groupOf(name) {
  if (/xlsx|workbook|sheet|cell|range|styles|chart/.test(name)) return 'XLSX'
  if (/docx|paragraph|bookmark|footnote/.test(name)) return 'DOCX'
  if (/pptx|slide|presentation|theme|shape/.test(name)) return 'PPTX'
  if (/pdf/.test(name)) return 'PDF'
  if (/engines|recalculate|rerender|automation/.test(name)) return '本地联动（可选）'
  if (/task|preview_operation|execute_operation|rollback/.test(name)) return '任务五件套'
  if (/convert/.test(name)) return '跨格式'
  return '文件与通用'
}

/**
 * 取一句话用途（描述里的第一句，去掉 markdown 强调符号）。
 * @param {string} description - 工具描述。
 * @returns {string} 短说明。
 */
function brief(description) {
  const cleaned = String(description ?? '')
    .replace(/\*\*/g, '')
    .replace(/\s+/g, ' ')
    .trim()
  const sentence = cleaned.split(/(?<=[。；;])/)[0] ?? cleaned
  return sentence.length > 58 ? `${sentence.slice(0, 57)}…` : sentence
}

const tools = definitions.map((d) => ({ name: d.name, description: d.description }))
const capabilities = readStringArray('CAPABILITIES')
const notImplemented = readStringArray('NOT_IMPLEMENTED')
const suites = readSuites()

const groups = new Map()
for (const tool of tools) {
  const group = groupOf(tool.name)
  if (!groups.has(group)) groups.set(group, [])
  groups.get(group).push(tool)
}
const order = ['文件与通用', 'XLSX', 'DOCX', 'PPTX', 'PDF', '跨格式', '本地联动（可选）', '任务五件套']
const groupNames = [...groups.keys()].sort((a, b) => order.indexOf(a) - order.indexOf(b))

const toolSection = [
  `共 **${tools.length}** 个工具（清单来自 \`dsh.plugin.json\`，本表由 \`npm run gen:capabilities\` 生成，不要手改）。`,
  ''
]
for (const group of groupNames) {
  toolSection.push(`### ${group}（${groups.get(group).length}）`, '')
  toolSection.push('| 工具 | 用途 |')
  toolSection.push('|---|---|')
  for (const tool of groups.get(group).sort((a, b) => a.name.localeCompare(b.name))) {
    toolSection.push(`| \`${tool.name}\` | ${brief(tool.description)} |`)
  }
  toolSection.push('')
}

const factSection = [
  '| 事实 | 值 |',
  '|---|---|',
  `| 插件版本 | \`${manifest.version}\` |`,
  `| 工具数 | **${tools.length}**（\`dsh.plugin.json\` 与运行时定义${nameDrift ? '**不一致**' : '一致'}） |`,
  `| 声明支持的能力 | **${capabilities.length}** 项（\`lib/capabilities.js\`） |`,
  `| 明确不做的能力 | **${notImplemented.length}** 项 |`,
  `| 测试套件 | **${suites.length}** 个 |`,
  `| 运行时依赖 | ${Object.keys(packageJson.dependencies ?? {}).length} 个（peer 依赖 ${Object.keys(packageJson.peerDependencies ?? {}).length} 个，由宿主提供） |`,
  `| Node 版本要求 | \`${packageJson.engines?.node ?? '未声明'}\` |`,
  ''
]
if (nameDrift) {
  factSection.push(
    '> ⚠️ `dsh.plugin.json` 与运行时定义的工具清单不一致：',
    `> 只在 manifest：${nameDrift.only_in_manifest.join('、') || '（无）'}；只在运行时：${nameDrift.only_in_runtime.join('、') || '（无）'}`,
    ''
  )
}

const suiteSection = ['| 套件 | 文件 |', '|---|---|']
for (const suite of suites) suiteSection.push(`| ${suite.title} | \`test/${suite.file}\` |`)
suiteSection.push('')

/** 明确不做的能力：一条一项，附一句「为什么不做」。 */
const NOT_DONE_NOTES = {
  'xlsx.chart.advanced': '只做原生图表创建与读取，不做高级图表（组合图/趋势线/误差线等）编辑',
  'xlsx.pivot.modify': '透视表只识别与报告，不修改',
  'xlsx.macro.execute': 'VBA 按原字节保留、只检测，从不执行',
  'pptx.validate.render': '不渲染，因此无法判定文本溢出、元素重叠、空白页',
  'pptx.animation.edit': '动画与切换只报告，不编辑',
  'pdf.text.edit': '不做 PDF 原位文本改写（需要重建内容流与字体，风险不可控）',
  'pdf.ocr': '本机无命令行 OCR、运行时也够不着 WinRT（核查见 docs/OCR能力核查.md）',
  'pdf.image.rasterize': '不做页面光栅化（没有可用的渲染器）',
  'pdf.to_office': 'PDF → Office 反向转换需要版面重建，规格 §二十三 自己列为不建议第一版做',
  'pdf.visual_regression': '视觉回归需要渲染成图片，本机没有光栅化器'
}
const notImplementedSection = notImplemented.map(
  (item) => `- \`${item}\`：${NOT_DONE_NOTES[item] ?? '（见 README 与 docs/能力与验证.md 的说明）'}`
)
notImplementedSection.push('')

/**
 * 用生成内容替换标记之间的部分。
 * @param {string} text - 原文。
 * @param {string} key - 标记名。
 * @param {string} body - 新内容。
 * @returns {string} 替换后的文本。
 */
function replaceBlock(text, key, body) {
  const begin = `<!-- GENERATED:${key}:BEGIN -->`
  const end = `<!-- GENERATED:${key}:END -->`
  const start = text.indexOf(begin)
  const stop = text.indexOf(end)
  if (start === -1 || stop === -1) throw new Error(`docs/能力矩阵.md 缺少标记 ${key}`)
  return `${text.slice(0, start + begin.length)}\n${body.join('\n')}\n${text.slice(stop)}`
}

if (!readFileSync('docs/能力矩阵.md', 'utf8').length && !checkOnly) {
  throw new Error('docs/能力矩阵.md 不存在')
}
const original = readFileSync(docPath, 'utf8')
let updated = replaceBlock(original, 'FACTS', factSection)
updated = replaceBlock(updated, 'TOOLS', toolSection)
updated = replaceBlock(updated, 'SUITES', suiteSection)
updated = replaceBlock(updated, 'NOT_IMPLEMENTED', notImplementedSection)

if (checkOnly) {
  if (nameDrift) {
    console.log('❌ `dsh.plugin.json` 与运行时定义的工具清单不一致')
    console.log(`   只在 manifest：${nameDrift.only_in_manifest.join('、') || '（无）'}`)
    console.log(`   只在运行时：${nameDrift.only_in_runtime.join('、') || '（无）'}`)
    process.exit(3)
  }
  if (updated !== original) {
    console.log('❌ docs/能力矩阵.md 与代码不一致（工具/能力/套件清单有变动）')
    console.log('   修复：npm run gen:capabilities')
    process.exit(3)
  }
  console.log(
    `✅ 能力矩阵与代码一致（${tools.length} 个工具 / ${capabilities.length} 项能力 / ${notImplemented.length} 项不做 / ${suites.length} 个套件）`
  )
  process.exit(0)
}

if (updated === original) {
  console.log(`能力矩阵已是最新（${tools.length} 个工具 / ${suites.length} 个套件）`)
} else {
  writeFileSync(docPath, updated, 'utf8')
  console.log(`已更新 docs/能力矩阵.md（${tools.length} 个工具 / ${suites.length} 个套件）`)
}
