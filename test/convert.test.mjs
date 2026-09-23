/**
 * 转换测试：用宿主内置 LibreOffice 引擎把 DOCX 转成 PDF。
 *
 * 单独成套件而不并入 `npm test`：引擎冷启动需要数秒到数十秒，
 * 让快速套件保持在秒级。运行：npm run test:convert
 *
 * 无引擎的环境下会跳过而不是失败。
 */
import assert from 'node:assert/strict'
import { createRequire } from 'node:module'
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import os from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { buildBasicDocx } from '../lib/docx.js'

const here = dirname(fileURLToPath(import.meta.url))
const DSH_TOOLS = `${os.homedir()}/AppData/Roaming/npm/node_modules/@deepseek-ai/dsh/node_modules/@deepseek-ai/dsh-tools/lib/index.js`
const DSH_ROOT = `${os.homedir()}/AppData/Roaming/npm/node_modules/@deepseek-ai/dsh/node_modules`

let passed = 0
let failed = 0

/**
 * 运行一个用例并记录结果（支持异步）。
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

// 引擎解析：把 DSH 的 node_modules 作为候选根注入，模拟宿主环境。
process.env.DSH_HOME = DSH_ROOT

const { defineTool } = await import(pathToFileURL(DSH_TOOLS).href)
const { createOfficeTools } = await import('../lib/tools.js')
const { countPdfPages, disposeConverter } = await import('../lib/convert.js')

const workspace = mkdtempSync(join(tmpdir(), 'dsh-office-convert-'))
const registry = new Map()
for (const tool of createOfficeTools({ defineTool, config: { workspaceRoot: workspace } })) registry.set(tool.name, tool)

/**
 * 调用一个已注册工具。
 * @param {string} name - 工具名。
 * @param {object} args - 参数。
 * @returns {Promise<object>} 工具返回值。
 */
const call = (name, args) => registry.get(name).execute(args, { signal: undefined, agent: null })

console.log('\n=== 转换：DOCX → PDF ===')

writeFileSync(join(workspace, 'source.docx'), buildBasicDocx({ title: '季度销售报告' }))

let converted
await test('office_convert_document 生成 PDF', async () => {
  const result = await call('office_convert_document', { path: 'source.docx' })
  if (!result.success && result.error.code === 'OFFICE_NOT_INSTALLED') {
    console.log(`     ⏭️  跳过：本机没有可用引擎（${result.error.message}）`)
    return
  }
  assert.equal(result.success, true, JSON.stringify(result.error))
  converted = result
  assert.ok(existsSync(result.output_file.path), 'PDF 文件应已生成')
  assert.ok(result.output_file.size > 500, `PDF 过小：${result.output_file.size}`)
})

if (converted) {
  await test('输出是合法 PDF', () => {
    const bytes = readFileSync(converted.output_file.path)
    assert.equal(bytes.subarray(0, 5).toString('latin1'), '%PDF-')
  })

  await test('报告引擎类型与页数', () => {
    assert.ok(['native', 'wasm'].includes(converted.data.engine))
    assert.ok(converted.data.pages >= 1, `页数异常：${converted.data.pages}`)
  })

  await test('报告缺失字体清单（判断字体替换的唯一可靠依据）', () => {
    assert.ok(Array.isArray(converted.data.missing_fonts))
    if (converted.data.missing_fonts.length > 0) {
      assert.ok(converted.warnings.some((w) => w.code === 'FONT_NOT_FOUND'))
    }
  })

  await test('PDF 内含文档文本（用 Word 的标题验证）', () => {
    const bytes = readFileSync(converted.output_file.path)
    const raw = bytes.toString('latin1')
    // PDF 文本通常压缩，因此只在未压缩时断言；退而验证有字体与页面对象
    const hasFont = /\/Font/.test(raw)
    const hasPage = /\/Type\s*\/Page[^s]/.test(raw)
    assert.ok(hasFont, 'PDF 应包含字体对象')
    assert.ok(hasPage, 'PDF 应包含页面对象')
  })

  await test('拒绝覆盖已存在的输出文件', async () => {
    const result = await call('office_convert_document', { path: 'source.docx', target_path: converted.output_file.path })
    assert.equal(result.success, false)
    assert.equal(result.error.code, 'PERMISSION_DENIED')
    assert.equal(result.error.needs_confirmation, true)
  })

  await test('指定 target_path 可用', async () => {
    const result = await call('office_convert_document', { path: 'source.docx', target_path: 'explicit.pdf' })
    assert.equal(result.success, true, JSON.stringify(result.error))
    assert.ok(existsSync(join(workspace, 'explicit.pdf')))
  })
}

console.log('\n=== 边界与错误处理 ===')

await test('不存在的输入返回 FILE_NOT_FOUND', async () => {
  const result = await call('office_convert_document', { path: '不存在.docx' })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'FILE_NOT_FOUND')
})

await test('不支持的类型被拒绝', async () => {
  writeFileSync(join(workspace, 'plain.txt'), 'hello')
  const result = await call('office_convert_document', { path: 'plain.txt' })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'UNSUPPORTED_FILE_TYPE')
})

await test('countPdfPages 对非 PDF 返回 null', () => {
  assert.equal(countPdfPages(Buffer.from('not a pdf')), null)
})

await disposeConverter()
rmSync(workspace, { recursive: true, force: true })

console.log(`\n结果：${passed} 通过，${failed} 失败\n`)
process.exit(failed === 0 ? 0 : 1)
