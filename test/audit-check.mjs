/**
 * 审计日志的**独立**验证：用 .NET 的 JSON 解析器（PowerShell `ConvertFrom-Json`）逐行读回来，
 * 再用与插件无关的方式断言「一调用一行、字段齐、无绝对路径」。
 *
 * 为什么不用 Node 自己读：Node 读自己的 JSON 只证明「自己写的自己认」。
 * 这里换一个实现（.NET 的 System.Text.Json）当第三方裁判 —— 顺带证明日志是**标准 JSONL**，
 * 别的工具（jq / 日志采集器 / Excel）也能吃。
 *
 * 用法：node test/audit-check.mjs（无需 Office；需要 pwsh）
 * 退出码：0 通过；1 断言失败；2 环境问题。
 */
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { spawnSync } from 'node:child_process'
import { apply } from '../lib/index.js'
import { resolveDefineTool } from '../lib/define-tool.js'

const root = mkdtempSync(join(tmpdir(), 'dsh-office-audit-check-'))
const problems = []
console.log('=== 审计日志独立验证 ===')
console.log(`工作区: ${root}`)

try {
  // 1) 产生日志：注册插件（开启审计）→ 调三个工具（一成功、一失败、一写文件）
  writeFileSync(join(root, 'demo.txt'), 'hello')
  const registry = new Map()
  const { defineTool } = await resolveDefineTool({})
  await apply(
    { tools: { register: (definition) => registry.set(definition.name, definition) }, logger: { info() {} } },
    { workspaceRoot: root, auditLogEnabled: true, defineTool }
  )
  const call = (name, args) => registry.get(name).execute(args, { signal: undefined, agent: null })
  await call('office_list_files', { directory: '.' })
  await call('office_read_workbook', { path: 'missing.xlsx' })
  await call('office_create_pdf', { path: 'report.pdf', lines: ['审计独立验证'] })

  const file = join(root, '.dsh-exp-office', 'audit', 'office-tools.jsonl')
  if (!existsSync(file)) throw new Error('未生成审计文件')
  const raw = readFileSync(file, 'utf8')
  const lineCount = raw.split('\n').filter((line) => line.trim() !== '').length
  console.log(`日志文件: ${file.replace(root, '<workspace>')}（${lineCount} 行，${raw.length} 字节）`)

  // 2) 第三方解析：PowerShell 的 ConvertFrom-Json 逐行解析
  const script = [
    '$ErrorActionPreference = "Stop"',
    `$lines = Get-Content -LiteralPath '${file.replace(/'/g, "''")}' -Encoding UTF8 | Where-Object { $_.Trim() -ne '' }`,
    '$parsed = @()',
    'foreach ($line in $lines) { $parsed += ($line | ConvertFrom-Json) }',
    'Write-Output "PARSED_ROWS: $($parsed.Count)"',
    'Write-Output "TOOL_ROWS: $(($parsed | Where-Object { $_.source -eq \'tool\' }).Count)"',
    'Write-Output "REGISTERED: $(($parsed | Where-Object { $_.event -eq \'registered\' }).Count)"',
    'Write-Output "FAILED_ROWS: $(($parsed | Where-Object { $_.ok -eq $false }).Count)"',
    'Write-Output "HAS_DURATION: $(($parsed | Where-Object { $_.source -eq \'tool\' -and $null -ne $_.duration_ms }).Count)"',
    'Write-Output "ABS_PATH_HITS: $((($parsed | ForEach-Object { $_.target }) | Where-Object { $_ -match \'^[A-Za-z]:\' }).Count)"'
  ].join('; ')
  const result = spawnSync('pwsh', ['-NoProfile', '-NonInteractive', '-Command', script], { encoding: 'utf8' })
  if (result.status !== 0) throw new Error(`PowerShell 解析失败：${(result.stderr ?? '').trim().split('\n').pop()}`)
  const stats = {}
  for (const line of result.stdout.split('\n')) {
    const match = /^([A-Z_]+):\s*(.*)$/.exec(line.trim())
    if (match) stats[match[1]] = match[2]
  }
  console.log(`第三方解析（.NET ConvertFrom-Json）：${JSON.stringify(stats)}`)

  // 3) 断言
  const expect = (label, actual, wanted) => {
    const ok = String(actual) === String(wanted)
    console.log(`  ${ok ? '✅' : '❌'} ${label}：${actual}${ok ? '' : `（期望 ${wanted}）`}`)
    if (!ok) problems.push(label)
  }
  expect('总行数 = 1 条装载事件 + 3 条调用', stats.PARSED_ROWS, 4)
  expect('工具调用行数', stats.TOOL_ROWS, 3)
  expect('装载事件行数', stats.REGISTERED, 1)
  expect('失败行数（应恰好 1 条 FILE_NOT_FOUND）', stats.FAILED_ROWS, 1)
  expect('每条调用都有耗时', stats.HAS_DURATION, 3)
  expect('target 里没有盘符绝对路径', stats.ABS_PATH_HITS, 0)

  // 4) 隐私硬断言（不依赖解析器）
  const hasAbsolute = /[A-Za-z]:[\\/]/.test(raw)
  console.log(`  ${hasAbsolute ? '❌' : '✅'} 原始日志不含任何盘符路径`)
  if (hasAbsolute) problems.push('原始日志含绝对路径')
  console.log(`  ${raw.includes(root) ? '❌' : '✅'} 原始日志不含工作区绝对路径`)
  if (raw.includes(root)) problems.push('原始日志含工作区路径')
} catch (err) {
  console.log(`❌ ${err.message}`)
  problems.push(err.message)
} finally {
  rmSync(root, { recursive: true, force: true })
}

console.log(`\n结论: ${problems.length === 0 ? '✅ 审计日志是标准 JSONL，字段齐、无绝对路径' : `❌ ${problems.length} 项未通过`}`)
process.exit(problems.length === 0 ? 0 : 1)
