/**
 * 测试总入口：依次运行全部测试套件并汇总结果。
 * 运行：npm test
 */
import { spawnSync } from 'node:child_process'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))

const SUITES = [
  ['引擎：ZIP 容器 / XML 最小修改 / 安全防护', 'engine.test.mjs'],
  ['XLSX 适配器', 'xlsx.test.mjs'],
  ['DOCX 适配器（读取 / 写入 / 校验）', 'docx.test.mjs'],
  ['PPTX 适配器（读取 / 写入 / 幻灯片管理）', 'pptx.test.mjs'],
  ['PDF 适配器（读取与文本提取）', 'pdf.test.mjs'],
  ['保真度：真实 Excel 多特性样本', 'fidelity.test.mjs'],
  ['保真度：真实 Word 复杂样本（批注/修订/目录）', 'word-fidelity.test.mjs'],
  ['宏启用工作簿（检测与字节保留）', 'macro.test.mjs'],
  ['插件集成：注册 / 端到端 / 事务 / 任务五件套', 'plugin.test.mjs'],
  ['并发与文件锁', 'concurrency.test.mjs']
]

let failedSuites = 0
const started = Date.now()

for (const [title, file] of SUITES) {
  console.log(`\n${'─'.repeat(70)}\n▶ ${title}  (${file})\n${'─'.repeat(70)}`)
  const result = spawnSync(process.execPath, [join(here, file)], { stdio: 'inherit' })
  if (result.status !== 0) failedSuites += 1
}

const seconds = ((Date.now() - started) / 1000).toFixed(1)
console.log(`\n${'═'.repeat(70)}`)
console.log(failedSuites === 0 ? `✅ 全部 ${SUITES.length} 个测试套件通过（${seconds}s）` : `❌ ${failedSuites}/${SUITES.length} 个测试套件失败（${seconds}s）`)
console.log('═'.repeat(70))
process.exit(failedSuites === 0 ? 0 : 1)
