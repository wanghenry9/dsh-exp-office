/**
 * 并发与文件锁测试（阶段 7「大文件和并发测试」的一半）。
 *
 * 验证的契约：
 *   1. 同一文件的并发写**串行化**，两边都成功、结果都落盘、文件仍然合法；
 *   2. 文件被别的任务持锁时，写入必须**报 FILE_LOCKED 且不改动文件**（绝不写坏）；
 *   3. 进程异常退出留下的**残留锁**超过存活时间会被回收，不会永久卡死；
 *   4. 成功路径不留锁文件。
 *
 * 运行：node test/concurrency.test.mjs
 */
import assert from 'node:assert/strict'
import { mkdtempSync, rmSync, writeFileSync, readFileSync, existsSync, utimesSync } from 'node:fs'
import { tmpdir } from 'node:os'
import os from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const DSH_TOOLS = `${os.homedir()}/AppData/Roaming/npm/node_modules/@deepseek-ai/dsh/node_modules/@deepseek-ai/dsh-tools/lib/index.js`
const { defineTool } = await import(pathToFileURL(DSH_TOOLS).href)
const { apply } = await import('../lib/index.js')
const { FileLock } = await import('../lib/workspace.js')

let passed = 0
let failed = 0

/**
 * 运行一个用例并记录结果。
 * @param {string} name - 用例名。
 * @param {() => unknown} fn - 用例体。
 * @returns {Promise<void>} 完成信号。
 */
async function test(name, fn) {
  try {
    await fn()
    passed += 1
    console.log(`  ✅ ${name}`)
  } catch (err) {
    failed += 1
    console.log(`  ❌ ${name}\n     ${err.message}`)
  }
}

const workspace = mkdtempSync(join(tmpdir(), 'dsh-office-concurrency-'))
const registry = new Map()
const ctx = {
  tools: {
    register(definition) {
      if (registry.has(definition.name)) throw new Error(`工具名重复：${definition.name}`)
      registry.set(definition.name, definition)
    }
  },
  logger: { info: () => {} }
}
await apply(ctx, { workspaceRoot: workspace, defineTool })

/**
 * 调用一个已注册工具。
 * @param {string} name - 工具名。
 * @param {object} args - 参数。
 * @returns {Promise<object>} 工具返回值。
 */
const call = (name, args) => {
  const tool = registry.get(name)
  assert.ok(tool, `工具 ${name} 未注册`)
  return tool.execute(args, { signal: undefined, agent: null })
}

console.log('\n=== 并发与文件锁 ===')

await test('准备一个工作簿', async () => {
  const created = await call('office_create_document', { path: 'concurrent.xlsx', sheet_name: '数据' })
  assert.equal(created.success, true)
  assert.ok(existsSync(join(workspace, 'concurrent.xlsx')))
})

await test('★ 同一文件并发写：串行化、两边的写入都落盘、文件仍然合法', async () => {
  const [a, b] = await Promise.all([
    call('office_write_cells', { path: 'concurrent.xlsx', sheet: '数据', cells: [{ ref: 'A1', value: '甲' }, { ref: 'A2', value: 1 }] }),
    call('office_write_cells', { path: 'concurrent.xlsx', sheet: '数据', cells: [{ ref: 'B1', value: '乙' }, { ref: 'B2', value: 2 }] })
  ])
  assert.equal(a.success, true, `第一个写入应成功：${JSON.stringify(a.error ?? {})}`)
  assert.equal(b.success, true, `第二个写入应成功（锁会等待而不是直接失败）：${JSON.stringify(b.error ?? {})}`)

  const read = await call('office_read_range', { path: 'concurrent.xlsx', sheet: '数据', range: 'A1:B2' })
  const values = read.data.rows.flat().map((cell) => cell?.value ?? null)
  assert.deepEqual(values, ['甲', '乙', 1, 2], `两次写入都必须保留：${JSON.stringify(values)}`)
  const valid = await call('office_validate_workbook', { path: 'concurrent.xlsx' })
  assert.equal(valid.data.valid, true)
})

await test('★ 文件被持锁时写入报 FILE_LOCKED，且绝不改动文件', async () => {
  const target = join(workspace, 'concurrent.xlsx')
  const before = readFileSync(target)
  const lock = await FileLock.acquire(target, { timeoutMs: 1000 })
  try {
    const result = await call('office_write_cells', { path: 'concurrent.xlsx', sheet: '数据', cells: [{ ref: 'C1', value: '不该写进去' }] })
    assert.equal(result.success, false, '被锁时必须失败而不是排队无限等')
    assert.equal(result.error.code, 'FILE_LOCKED')
    assert.ok(readFileSync(target).equals(before), '被锁拒绝时文件必须逐字节不变')
  } finally {
    await lock.release()
  }
  // 解锁后同样的写入应当成功
  const after = await call('office_write_cells', { path: 'concurrent.xlsx', sheet: '数据', cells: [{ ref: 'C1', value: '解锁后写入' }] })
  assert.equal(after.success, true)
})

await test('★ 残留锁超过存活时间会被回收（不会永久卡死）', async () => {
  const target = join(workspace, 'concurrent.xlsx')
  const lockPath = `${target}.dsh-office.lock`
  writeFileSync(lockPath, JSON.stringify({ pid: 999999, at: '2000-01-01T00:00:00Z' }))
  const old = new Date(Date.now() - 11 * 60 * 1000)
  utimesSync(lockPath, old, old)
  const result = await call('office_write_cells', { path: 'concurrent.xlsx', sheet: '数据', cells: [{ ref: 'D1', value: '回收后写入' }] })
  assert.equal(result.success, true, `残留锁应被回收：${JSON.stringify(result.error ?? {})}`)
  assert.equal(existsSync(lockPath), false, '回收后锁文件不应残留')
})

await test('成功路径不留锁文件', () => {
  const lockPath = join(workspace, 'concurrent.xlsx.dsh-office.lock')
  assert.equal(existsSync(lockPath), false)
})

await test('并发读不互相阻塞（读不加排他锁）', async () => {
  const started = Date.now()
  const results = await Promise.all([
    call('office_read_range', { path: 'concurrent.xlsx', sheet: '数据', range: 'A1:D2' }),
    call('office_read_range', { path: 'concurrent.xlsx', sheet: '数据', range: 'A1:D2' }),
    call('office_read_sheet', { path: 'concurrent.xlsx', sheet: '数据', limit: 10 })
  ])
  assert.ok(results.every((r) => r.success), '并发读应全部成功')
  assert.ok(Date.now() - started < 5000, '并发读不应被锁阻塞')
})

rmSync(workspace, { recursive: true, force: true })

console.log(`\n结果：${passed} 通过，${failed} 失败\n`)
process.exit(failed === 0 ? 0 : 1)
