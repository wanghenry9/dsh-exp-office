/**
 * 统一操作审计（lib/audit.js）的测试。
 *
 * 关注三件事：
 *   1. **格式与字段**：一行一个合法 JSON，工具名 / 耗时 / 成败 / 错误码 / 目标文件都在；
 *   2. **隐私**：默认只记文件名，日志里**不能出现绝对路径**；
 *   3. **不添麻烦**：审计关闭时不产生文件；写盘失败也绝不让工具失败。
 */
import assert from 'node:assert/strict'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { ToolAuditLog, withToolAudit, withToolAuditAll } from '../lib/audit.js'
import { apply } from '../lib/index.js'
import { createOfficeTools } from '../lib/tools.js'
import { resolveDefineTool } from '../lib/define-tool.js'

let passed = 0
let failed = 0

/**
 * 跑一个用例。
 * @param {string} title - 用例名。
 * @param {Function} fn - 用例体。
 * @returns {Promise<void>} 完成。
 */
async function test(title, fn) {
  try {
    await fn()
    passed += 1
    console.log(`  ✅ ${title}`)
  } catch (err) {
    failed += 1
    console.log(`  ❌ ${title}\n       ${err.message}`)
  }
}

/**
 * 建一个临时工作区。
 * @returns {string} 目录。
 */
function tempDir() {
  return mkdtempSync(join(tmpdir(), 'dsh-exp-office-audit-'))
}

/**
 * 读 JSONL 并逐行解析。
 * @param {string} file - 文件。
 * @returns {object[]} 记录列表。
 */
function readEntries(file) {
  return readFileSync(file, 'utf8')
    .split('\n')
    .filter((line) => line.trim() !== '')
    .map((line) => JSON.parse(line))
}

console.log('\n=== 1. 路径记录方式（隐私开关）===')

await test('basename 模式：绝对路径只留文件名，相对路径原样', () => {
  const log = new ToolAuditLog({ file: 'x', workspaceRoot: 'E:\\ws' })
  assert.equal(log.describePath('E:\\ws\\sub\\报告.xlsx'), '报告.xlsx')
  assert.equal(log.describePath('out/report.pdf'), 'out/report.pdf')
  assert.equal(log.describePath(''), null)
  assert.equal(log.describePath(undefined), null)
})

await test('relative 模式：工作区内给相对路径，工作区外退回文件名', () => {
  const dir = tempDir()
  const log = new ToolAuditLog({ file: 'x', workspaceRoot: dir, pathMode: 'relative' })
  assert.equal(log.describePath(join(dir, 'a', 'b.xlsx')), join('a', 'b.xlsx'))
  assert.equal(log.describePath(join(tmpdir(), 'outside', 'c.xlsx')), 'c.xlsx')
  rmSync(dir, { recursive: true, force: true })
})

await test('非法的 pathMode 退回 basename（不静默变成别的行为）', () => {
  const log = new ToolAuditLog({ file: 'x', workspaceRoot: 'E:\\ws', pathMode: 'full' })
  assert.equal(log.describePath('E:\\ws\\a\\b.xlsx'), 'b.xlsx')
})

console.log('\n=== 2. 写入与轮转 ===')

await test('record 写出单行合法 JSON，可逐行解析', () => {
  const dir = tempDir()
  const file = join(dir, 'audit.jsonl')
  const log = new ToolAuditLog({ file, workspaceRoot: dir })
  assert.equal(log.record({ time: 'T', tool: 'a', ok: true }), true)
  assert.equal(log.record({ time: 'T', tool: 'b', ok: false }), true)
  const entries = readEntries(file)
  assert.equal(entries.length, 2)
  assert.deepEqual(entries.map((e) => e.tool), ['a', 'b'])
  rmSync(dir, { recursive: true, force: true })
})

await test('超过上限自动轮转成 .1，原文件重新开始', () => {
  const dir = tempDir()
  const file = join(dir, 'audit.jsonl')
  const log = new ToolAuditLog({ file, workspaceRoot: dir, maxBytes: 200 })
  for (let i = 0; i < 40; i += 1) log.record({ i, padding: 'x'.repeat(40) })
  assert.ok(existsSync(`${file}.1`), '应产生轮转备份')
  const backup = readEntries(`${file}.1`)
  const current = readEntries(file)
  assert.ok(backup.length > 0, '备份里应有内容')
  assert.ok(current.length > 0, '当前文件应继续写入')
  assert.ok(backup[0].i < current[0].i, '备份应是更早的记录')
  rmSync(dir, { recursive: true, force: true })
})

await test('写盘失败不抛异常，只累加 writeErrors', () => {
  const dir = tempDir()
  // 把「文件路径」指向一个目录：appendFileSync 必然失败
  const log = new ToolAuditLog({ file: dir, workspaceRoot: dir })
  assert.equal(log.record({ a: 1 }), false)
  assert.equal(log.writeErrors, 1)
  assert.ok(typeof log.lastError === 'string' && log.lastError.length > 0)
  rmSync(dir, { recursive: true, force: true })
})

console.log('\n=== 3. 工具中间件 ===')

await test('返回值原样透传，成功/失败/异常都记一行', async () => {
  const dir = tempDir()
  const file = join(dir, 'audit.jsonl')
  const log = new ToolAuditLog({ file, workspaceRoot: dir })
  const okTool = withToolAudit(
    { name: 'office_ok', execute: async () => ({ success: true, request_id: 'r1', changes: [1, 2], output_file: { size: 7 } }) },
    log
  )
  const failTool = withToolAudit({ name: 'office_fail', execute: async () => ({ success: false, request_id: 'r2', error: { code: 'FILE_NOT_FOUND' } }) }, log)
  const throwTool = withToolAudit(
    { name: 'office_throw', execute: async () => { throw Object.assign(new Error('boom'), { code: 'INTERNAL_ERROR' }) } },
    log
  )
  const value = await okTool.execute({ path: join(dir, 'a.xlsx') }, {})
  assert.deepEqual(value, { success: true, request_id: 'r1', changes: [1, 2], output_file: { size: 7 } })
  await failTool.execute({ path: 'b.xlsx' }, {})
  await assert.rejects(() => throwTool.execute({}, {}), /boom/)
  const entries = readEntries(file)
  assert.equal(entries.length, 3)
  assert.deepEqual(entries[0].tool, 'office_ok')
  assert.equal(entries[0].ok, true)
  assert.equal(entries[0].target, 'a.xlsx')
  assert.equal(entries[0].changes_count, 2)
  assert.equal(entries[0].output_bytes, 7)
  assert.equal(entries[0].request_id, 'r1')
  assert.equal(typeof entries[0].duration_ms, 'number')
  assert.equal(entries[1].ok, false)
  assert.equal(entries[1].error_code, 'FILE_NOT_FOUND')
  assert.equal(entries[2].ok, false)
  assert.equal(entries[2].error_code, 'INTERNAL_ERROR')
  assert.equal(entries[2].error_message, 'boom')
  rmSync(dir, { recursive: true, force: true })
})

await test('withToolAuditAll 保持工具数量与顺序', () => {
  const dir = tempDir()
  const log = new ToolAuditLog({ file: join(dir, 'a.jsonl'), workspaceRoot: dir })
  const definitions = [{ name: 'a', execute: async () => ({ success: true }) }, { name: 'b', execute: async () => ({ success: true }) }]
  const wrapped = withToolAuditAll(definitions, log)
  assert.deepEqual(wrapped.map((d) => d.name), ['a', 'b'])
  assert.notEqual(wrapped[0].execute, definitions[0].execute, '应换成包装后的 execute')
  rmSync(dir, { recursive: true, force: true })
})

console.log('\n=== 4. 插件集成（真注册 90 个工具）===')

await test('★ 开启审计：注册事件 + 每次调用各一行，且日志里没有绝对路径', async () => {
  const root = tempDir()
  writeFileSync(join(root, 'demo.txt'), 'hello')
  const registry = new Map()
  const { defineTool } = await resolveDefineTool({})
  await apply(
    { tools: { register: (definition) => registry.set(definition.name, definition) }, logger: { info() {} } },
    { workspaceRoot: root, auditLogEnabled: true, defineTool }
  )
  assert.equal(registry.size, 90)
  const call = (name, args) => registry.get(name).execute(args, { signal: undefined, agent: null })
  await call('office_list_files', { directory: '.' })
  await call('office_read_workbook', { path: 'missing.xlsx' })
  await call('office_create_pdf', { path: 'report.pdf', lines: ['审计'] })
  const file = join(root, '.dsh-exp-office', 'audit', 'office-tools.jsonl')
  assert.ok(existsSync(file), '应生成审计文件')
  const entries = readEntries(file)
  assert.equal(entries.length, 4, '1 条装载事件 + 3 条调用')
  assert.equal(entries[0].source, 'plugin')
  assert.equal(entries[0].event, 'registered')
  assert.equal(entries[0].tools, 90)
  assert.deepEqual(entries.slice(1).map((e) => e.source), ['tool', 'tool', 'tool'])
  assert.deepEqual(entries.slice(1).map((e) => e.tool), ['office_list_files', 'office_read_workbook', 'office_create_pdf'])
  assert.deepEqual(entries.slice(1).map((e) => e.ok), [true, false, true])
  assert.equal(entries[2].error_code, 'FILE_NOT_FOUND')
  // 隐私硬断言：整个文件里不能出现盘符路径
  const raw = readFileSync(file, 'utf8')
  assert.ok(!/[A-Za-z]:[\\/]/.test(raw), `审计日志不应包含绝对路径：${raw.slice(0, 200)}`)
  assert.ok(!raw.includes(root), '不应包含工作区绝对路径')
  rmSync(root, { recursive: true, force: true })
})

await test('★ 默认关闭：不传配置就不产生任何审计文件', async () => {
  const root = tempDir()
  const registry = new Map()
  const { defineTool } = await resolveDefineTool({})
  await apply(
    { tools: { register: (definition) => registry.set(definition.name, definition) }, logger: { info() {} } },
    { workspaceRoot: root, defineTool }
  )
  await registry.get('office_list_files').execute({ directory: '.' }, { signal: undefined, agent: null })
  assert.equal(existsSync(join(root, '.dsh-exp-office')), false, '默认不应创建审计目录')
  rmSync(root, { recursive: true, force: true })
})

await test('审计配置不合法时退回安全默认（basename）', async () => {
  const root = tempDir()
  const registry = new Map()
  const { defineTool } = await resolveDefineTool({})
  await apply(
    { tools: { register: (definition) => registry.set(definition.name, definition) }, logger: { info() {} } },
    { workspaceRoot: root, auditLogEnabled: true, auditLogPathMode: '不存在的模式', auditLogMaxBytes: -5, defineTool }
  )
  await registry.get('office_list_files').execute({ path: join(root, 'x.txt') }, { signal: undefined, agent: null })
  const entries = readEntries(join(root, '.dsh-exp-office', 'audit', 'office-tools.jsonl'))
  assert.equal(entries[1].target, 'x.txt')
  rmSync(root, { recursive: true, force: true })
})

await test('createOfficeTools 不受审计影响（仍是 90 个，且没有包装）', async () => {
  const dir = tempDir()
  const { defineTool } = await resolveDefineTool({})
  const definitions = createOfficeTools({ defineTool, config: { workspaceRoot: dir } })
  assert.equal(definitions.length, 90)
  mkdirSync(join(dir, 'x'), { recursive: true })
  rmSync(dir, { recursive: true, force: true })
})

console.log(`\n结果：${passed} 通过，${failed} 失败\n`)
process.exit(failed === 0 ? 0 : 1)
