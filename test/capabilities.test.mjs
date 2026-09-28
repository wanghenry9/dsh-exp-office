/**
 * 能力矩阵与清单一致性：`dsh.plugin.json` ↔ 运行时定义 ↔ `docs/能力矩阵.md`。
 *
 * 为什么把它放进 `npm test`：这三份东西最容易「改了一处忘了另一处」——
 * 加了工具却没更新清单，或文档里的工具表与代码不符，都会在交付时变成假信息。
 * 这里用**真实的运行时定义**对账，并通过 `--check` 校验文档不漂移。
 */
import assert from 'node:assert/strict'
import { spawnSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { createOfficeTools } from '../lib/tools.js'
import { resolveDefineTool } from '../lib/define-tool.js'
import { toCapabilityManifest } from '../lib/capabilities.js'

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

const here = dirname(fileURLToPath(import.meta.url))
const root = join(here, '..')
const manifest = JSON.parse(readFileSync(join(root, 'dsh.plugin.json'), 'utf8'))
const workspace = mkdtempSync(join(tmpdir(), 'dsh-office-capabilities-'))
const { defineTool } = await resolveDefineTool({})
const definitions = createOfficeTools({ defineTool, config: { workspaceRoot: workspace } })
rmSync(workspace, { recursive: true, force: true })

console.log('\n=== 能力矩阵一致性 ===')

await test('★ dsh.plugin.json 声明的工具与运行时定义逐一一致', () => {
  const declared = [...manifest.contributes.tools].sort()
  const runtime = definitions.map((d) => d.name).sort()
  assert.equal(runtime.length, declared.length, `数量不一致：manifest ${declared.length} / 运行时 ${runtime.length}`)
  assert.deepEqual(runtime, declared, '清单与运行时定义必须完全相同')
})

await test('每个工具都有非空描述与参数 schema（能力矩阵要展示用途）', () => {
  for (const definition of definitions) {
    assert.ok(typeof definition.description === 'string' && definition.description.length > 5, `${definition.name} 缺少描述`)
    assert.ok(definition.parameters !== undefined, `${definition.name} 缺少参数 schema`)
    assert.equal(typeof definition.output?.render, 'function', `${definition.name} 缺少 output.render（0.1.7 起必需）`)
  }
})

await test('能力清单与未实现清单没有交集、也不重复', () => {
  const capabilityManifest = toCapabilityManifest()
  const capabilities = capabilityManifest.capabilities
  const notImplemented = capabilityManifest.not_implemented
  assert.equal(new Set(capabilities).size, capabilities.length, '能力清单里有重复项')
  assert.equal(new Set(notImplemented).size, notImplemented.length, '未实现清单里有重复项')
  const overlap = capabilities.filter((item) => notImplemented.includes(item))
  assert.deepEqual(overlap, [], `同一项能力不能既声明支持又声明不做：${overlap.join('、')}`)
})

await test('★ docs/能力矩阵.md 与代码不漂移（--check 模式）', () => {
  const result = spawnSync(process.execPath, [join(here, 'capability-matrix.mjs'), '--check'], { encoding: 'utf8' })
  assert.equal(result.status, 0, `文档漂移：${(result.stdout ?? '').trim().split('\n').slice(-2).join(' | ')}`)
})

await test('能力矩阵文档带有生成标记（否则 gen 会写坏文档）', () => {
  const doc = readFileSync(join(root, 'docs', '能力矩阵.md'), 'utf8')
  for (const key of ['FACTS', 'TOOLS', 'SUITES', 'NOT_IMPLEMENTED']) {
    assert.ok(doc.includes(`<!-- GENERATED:${key}:BEGIN -->`), `缺少标记 ${key}:BEGIN`)
    assert.ok(doc.includes(`<!-- GENERATED:${key}:END -->`), `缺少标记 ${key}:END`)
  }
})

console.log(`\n结果：${passed} 通过，${failed} 失败\n`)
process.exit(failed === 0 ? 0 : 1)
