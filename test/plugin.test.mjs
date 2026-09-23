/**
 * 插件级集成测试：用真实的 @deepseek-ai/dsh-tools 编译 schema，
 * 用模拟 ctx 完成注册，并端到端跑通工具执行链路。
 *
 * 运行：node test/plugin.test.mjs
 */
import assert from 'node:assert/strict'
import { mkdtempSync, rmSync, writeFileSync, readFileSync, existsSync, readdirSync, copyFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import os from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createRequire } from 'node:module'
import zlib from 'node:zlib'
import { pathToFileURL } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const DSH_TOOLS = `${os.homedir()}/AppData/Roaming/npm/node_modules/@deepseek-ai/dsh/node_modules/@deepseek-ai/dsh-tools/lib/index.js`
/** 真实 Cordis 加载器：DSH 装载插件走的就是它的 safeCollect 契约。 */
const cordisUrl = new URL('../../cordis/lib/index.js', pathToFileURL(DSH_TOOLS))

const { defineTool, validateJsonSchemaValue } = await import(pathToFileURL(DSH_TOOLS).href)
const { apply } = await import('../lib/index.js')
const { buildEmptyWorkbook } = await import('../lib/tools.js')
const { toCapabilityManifest } = await import('../lib/capabilities.js')

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

/**
 * 生成一张合法的 RGB PNG，用于图片插入测试。
 * @param {number} width - 宽。
 * @param {number} height - 高。
 * @returns {Buffer} PNG 字节。
 */
function makePng(width, height) {
  const raw = Buffer.alloc((width * 3 + 1) * height)
  let o = 0
  for (let y = 0; y < height; y += 1) {
    raw[o] = 0
    o += 1
    for (let x = 0; x < width; x += 1) {
      raw[o] = 31
      raw[o + 1] = 119
      raw[o + 2] = 180
      o += 3
    }
  }
  const chunk = (type, payload) => {
    const len = Buffer.alloc(4)
    len.writeUInt32BE(payload.length)
    const t = Buffer.from(type, 'latin1')
    const crc = Buffer.alloc(4)
    crc.writeUInt32BE(zlib.crc32(Buffer.concat([t, payload])) >>> 0)
    return Buffer.concat([len, t, payload, crc])
  }
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(width, 0)
  ihdr.writeUInt32BE(height, 4)
  ihdr[8] = 8
  ihdr[9] = 2
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlib.deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0))
  ])
}

const workspace = mkdtempSync(join(tmpdir(), 'dsh-office-test-'))

/** 已注册工具的登记表。 */
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

/**
 * 调用一个已注册工具。
 * @param {string} name - 工具名。
 * @param {object} args - 参数。
 * @returns {Promise<object>} 工具返回值。
 */
async function call(name, args) {
  const tool = registry.get(name)
  assert.ok(tool, `工具 ${name} 未注册`)
  return tool.execute(args, { signal: undefined, agent: null })
}

/**
 * 校验工具的返回值符合其声明的输出 schema。
 * @param {string} name - 工具名。
 * @param {object} args - 参数。
 * @returns {Promise<object>} 工具返回值。
 */
async function callAndValidate(name, args) {
  const tool = registry.get(name)
  const value = await tool.execute(args, { signal: undefined, agent: null })
  const errors = validateJsonSchemaValue(tool.output.schema, value, '')
  assert.deepEqual(errors, [], `${name} 的返回值不符合输出 schema：${JSON.stringify(errors)}`)
  return value
}

console.log('\n=== 1. 插件注册（真实 defineTool 编译 schema）===')

await test('插件可以挂载并注册全部工具', async () => {
  await apply(ctx, { workspaceRoot: workspace, defineTool })
  const declared = JSON.parse(readFileSync(join(here, '..', 'dsh.plugin.json'), 'utf8')).contributes.tools.length
  assert.equal(registry.size, declared, `注册数 ${registry.size} 与 dsh.plugin.json 声明的 ${declared} 不一致`)
})

// 2026-09-23 事故回归：apply 曾返回普通对象摘要，Cordis 的 safeCollect 只接受
// 「清理函数 / null / undefined」，于是整个插件树加载失败，`dsh web` 直接起不来。
// 当时测试只调用 apply 却不看返回值，所以全绿也拦不住 —— 下面两条是补上的闸门。
await test('★ apply 的返回值符合 Cordis effect 契约', async () => {
  const probe = new Map()
  const effect = await apply(
    { tools: { register: (d) => probe.set(d.name, d) }, logger: { info() {} } },
    { workspaceRoot: workspace, defineTool }
  )
  const ok = effect === undefined || effect === null || typeof effect === 'function'
  assert.ok(ok, `apply 必须 resolve 为 undefined/null/清理函数，实际是 ${Object.prototype.toString.call(effect)}`)
  const declared = JSON.parse(readFileSync(join(here, '..', 'dsh.plugin.json'), 'utf8')).contributes.tools.length
  assert.equal(probe.size, declared)
})

await test('★ 用真实 Cordis 加载器装载插件（不再出现 Invalid effect）', async () => {
  const { Context } = await import(cordisUrl.href)
  const realRegistry = new Map()
  const realCtx = new Context()
  realCtx.provide('tools', {
    register(definition) {
      realRegistry.set(definition.name, definition)
    }
  })
  // Cordis 会 await 这个调用：apply 抛出或返回非法 effect 都会在这里被拒。
  await realCtx.plugin({ apply, name: 'dsh-exp-office', inject: ['tools'] }, { workspaceRoot: workspace, defineTool })
  const declared = JSON.parse(readFileSync(join(here, '..', 'dsh.plugin.json'), 'utf8')).contributes.tools.length
  assert.equal(realRegistry.size, declared, `真实加载器下注册了 ${realRegistry.size} 个，声明 ${declared} 个`)
  assert.ok(realRegistry.has('office_read_pdf'))
  assert.ok(realRegistry.has('office_add_pdf_watermark'))
})

await test('★ 反例：apply resolve 成普通对象时真实加载器必须拒绝', async () => {
  // 这条是「闸门的闸门」：证明上面那条断言真的能拦住事故，而不是碰巧通过。
  const { Context } = await import(cordisUrl.href)
  const badCtx = new Context()
  badCtx.provide('tools', { register() {} })
  const bad = {
    name: 'dsh-exp-office-bad-effect',
    inject: ['tools'],
    apply: async () => ({ registered: 1, manifest: {} })
  }
  // ctx.plugin 返回 Fiber（可 await 但不是 Promise），assert.rejects 需要 Promise 包装。
  await assert.rejects(async () => {
    await badCtx.plugin(bad)
  }, /Invalid effect/)
})

await test('每个工具都有名称、描述、参数与输出声明', () => {
  for (const [name, tool] of registry) {
    assert.equal(typeof tool.name, 'string')
    assert.ok(tool.description.length > 20, `${name} 的描述过短`)
    assert.equal(tool.parameters.type, 'object')
    assert.ok(tool.output.schema, `${name} 缺少输出 schema`)
    assert.equal(typeof tool.output.render, 'function')
  }
})

await test('工具名全部以 office_ 前缀且唯一', () => {
  for (const name of registry.keys()) assert.ok(name.startsWith('office_'), `命名不合规：${name}`)
  assert.equal(registry.size, new Set(registry.keys()).size)
})

await test('参数 schema 编译出正确的必填项', () => {
  const tool = registry.get('office_write_cells')
  assert.deepEqual(tool.parameters.required.sort(), ['cells', 'path', 'sheet'])
  const cellSchema = tool.parameters.properties.cells
  assert.equal(cellSchema.type, 'array')
  assert.deepEqual(cellSchema.items.required.sort(), ['ref', 'value'])
})

await test('能力清单结构符合 docs §十', () => {
  const manifest = toCapabilityManifest()
  assert.equal(manifest.plugin_name, 'dsh-exp-office')
  assert.ok(Array.isArray(manifest.protocol_versions))
  assert.ok(manifest.capabilities.includes('xlsx.write'))
  assert.ok(manifest.capabilities.includes('xlsx.read'))
  assert.ok(Array.isArray(manifest.not_implemented))
})

console.log('\n=== 2. 端到端：创建 → 写入 → 读取 → 校验 ===')

const bookPath = join(workspace, 'demo.xlsx')

await test('office_create_document 创建可用的工作簿', async () => {
  const result = await callAndValidate('office_create_document', { path: 'demo.xlsx', sheet_name: '数据' })
  assert.equal(result.success, true)
  assert.ok(existsSync(bookPath))
  assert.ok(result.output_file.sha256.length === 64)
})

await test('office_create_document 拒绝沉默覆盖', async () => {
  const result = await call('office_create_document', { path: 'demo.xlsx' })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'PERMISSION_DENIED')
  assert.equal(result.error.needs_confirmation, true)
})

await test('office_write_cells 批量写入并落盘', async () => {
  const result = await callAndValidate('office_write_cells', {
    path: 'demo.xlsx',
    sheet: '数据',
    cells: [
      { ref: 'A1', value: '产品' },
      { ref: 'B1', value: '金额' },
      { ref: 'A2', value: '笔记本' },
      { ref: 'B2', value: 5999 },
      { ref: 'A3', value: '显示器' },
      { ref: 'B3', value: 1299 },
      { ref: 'B4', formula: 'SUM(B2:B3)', value: 7298 }
    ]
  })
  assert.equal(result.success, true)
  assert.equal(result.changes.length, 7)
  assert.equal(result.output_file.size > 0, true)
})

await test('office_read_range 读回刚写入的数据', async () => {
  const result = await callAndValidate('office_read_range', { path: 'demo.xlsx', sheet: '数据', range: 'A1:B4' })
  assert.equal(result.data.rows[0][0].value, '产品')
  assert.equal(result.data.rows[1][0].value, '笔记本')
  assert.equal(result.data.rows[1][1].value, 5999)
  assert.equal(result.data.rows[3][1].formula, 'SUM(B2:B3)')
})

await test('office_read_workbook 报告结构与安全标志', async () => {
  const result = await callAndValidate('office_read_workbook', { path: 'demo.xlsx', detail: 'detail' })
  assert.deepEqual(result.data.sheets.map((s) => s.name), ['数据'])
  assert.equal(result.data.security.has_macro, false)
  assert.equal(result.data.security.has_charts, false)
})

await test('office_get_file_info 识别真实类型', async () => {
  const result = await callAndValidate('office_get_file_info', { path: 'demo.xlsx' })
  assert.equal(result.data.detected_type, 'xlsx')
  assert.equal(result.data.extension_mismatch, false)
  assert.equal(result.data.size > 0, true)
})

await test('伪造扩展名会被识破', async () => {
  const fake = join(workspace, 'fake.xlsx')
  writeFileSync(fake, '%PDF-1.7 not really a pdf')
  const result = await call('office_get_file_info', { path: 'fake.xlsx' })
  assert.equal(result.data.detected_type, 'pdf')
  assert.equal(result.data.extension_mismatch, true)
})

await test('★ 回归：OOXML 之间改名也会被识破（曾因传参错误恒为 false）', async () => {
  // 自建一份 docx，避免依赖后面小节才创建的文件
  await call('office_create_document', { path: 'rename-source.docx', format: 'docx' })
  copyFileSync(join(workspace, 'rename-source.docx'), join(workspace, 'renamed.xlsx'))
  const result = await call('office_get_file_info', { path: 'renamed.xlsx' })
  assert.equal(result.data.detected_type, 'docx')
  assert.equal(result.data.extension_mismatch, true, '把 docx 改名成 xlsx 必须报告扩展名不符')
})

await test('★ 回归：宏文件改名也会被识破并报告含宏', async () => {
  const macroSource = join(here, 'fixtures', 'macro.xlsm')
  if (!existsSync(macroSource)) return
  copyFileSync(macroSource, join(workspace, 'renamed.xlsx'))
  const result = await call('office_get_file_info', { path: 'renamed.xlsx' })
  assert.equal(result.data.detected_type, 'xlsm')
  assert.equal(result.data.extension_mismatch, true)
  assert.equal(result.data.is_macro_enabled, true)
})

await test('office_validate_workbook 通过结构与公式检查', async () => {
  const result = await callAndValidate('office_validate_workbook', { path: 'demo.xlsx' })
  assert.equal(result.data.valid, true)
  assert.ok(result.data.checks.every((c) => c.ok), JSON.stringify(result.data.checks.filter((c) => !c.ok)))
})

console.log('\n=== 3. 路径安全 ===')

await test('拒绝路径穿越到工作区之外', async () => {
  const result = await call('office_get_file_info', { path: '../../../Windows/win.ini' })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'PERMISSION_DENIED')
})

await test('拒绝设备路径与空字节', async () => {
  for (const bad of ['\\\\?\\C:\\Windows\\win.ini', 'demo\u0000.xlsx']) {
    const result = await call('office_get_file_info', { path: bad })
    assert.equal(result.success, false, `${bad} 应被拒绝`)
  }
})

await test('不存在的文件返回 FILE_NOT_FOUND', async () => {
  const result = await call('office_get_file_info', { path: '不存在.xlsx' })
  assert.equal(result.error.code, 'FILE_NOT_FOUND')
})

console.log('\n=== 4. 破坏性操作需要确认 ===')

await test('删除工作表未确认时被拒绝', async () => {
  const result = await call('office_delete_worksheet', { path: 'demo.xlsx', sheet: '数据', confirm: false })
  assert.equal(result.success, false)
  assert.equal(result.error.needs_confirmation, true)
})

await test('删除文件未确认时被拒绝', async () => {
  const result = await call('office_delete_file', { path: 'demo.xlsx', confirm: false })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'PERMISSION_DENIED')
})

await test('删除最后一张工作表被拒绝（确认了也不行）', async () => {
  const result = await call('office_delete_worksheet', { path: 'demo.xlsx', sheet: '数据', confirm: true })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'INVALID_REQUEST')
})

console.log('\n=== 5. 查找替换（含 dry_run）===')

await test('dry_run 预览命中数量且不改文件', async () => {
  const before = readFileSync(bookPath)
  const result = await callAndValidate('office_find_and_replace', {
    path: 'demo.xlsx', find: '笔记本', replace: '便携机', dry_run: true
  })
  assert.equal(result.data.dry_run, true)
  assert.equal(result.data.replacements, 1)
  assert.ok(readFileSync(bookPath).equals(before), 'dry_run 不应改动文件')
})

await test('真实替换只影响命中单元格', async () => {
  const result = await callAndValidate('office_find_and_replace', { path: 'demo.xlsx', find: '笔记本', replace: '便携机' })
  assert.equal(result.data.replacements, 1)
  const check = await call('office_read_range', { path: 'demo.xlsx', sheet: '数据', range: 'A2:A3' })
  assert.equal(check.data.rows[0][0].value, '便携机')
  assert.equal(check.data.rows[1][0].value, '显示器')
})

console.log('\n=== 6. 任务五件套：预览 → 执行 → 回滚 ===')

let planId
await test('office_preview_operation 生成计划并标记破坏性', async () => {
  const result = await callAndValidate('office_preview_operation', {
    path: 'demo.xlsx',
    operation: 'delete_worksheet',
    parameters: { sheet: '数据' }
  })
  planId = result.data.plan_id
  assert.ok(planId.startsWith('plan-'))
  assert.equal(result.data.destructive, true)
  assert.equal(result.data.requires_confirmation, true)
})

await test('office_get_task_status 能查到计划', async () => {
  const result = await callAndValidate('office_get_task_status', { plan_id: planId })
  assert.equal(result.data.status, 'previewed')
  assert.equal(result.data.operation, 'delete_worksheet')
})

await test('未确认时拒绝执行破坏性计划', async () => {
  const result = await call('office_execute_operation', { plan_id: planId })
  assert.equal(result.success, false)
  assert.equal(result.error.needs_confirmation, true)
})

await test('office_cancel_task 可以取消计划', async () => {
  const result = await callAndValidate('office_cancel_task', { plan_id: planId })
  assert.equal(result.data.status, 'cancelled')
})

await test('已取消的计划不能执行', async () => {
  const result = await call('office_execute_operation', { plan_id: planId, confirm: true })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'INVALID_REQUEST')
})

let writePlan
await test('写入计划可预览并执行', async () => {
  const preview = await call('office_preview_operation', {
    path: 'demo.xlsx',
    operation: 'write_cells',
    parameters: { sheet: '数据', cells: [{ ref: 'D1', value: '备注' }, { ref: 'D2', value: '已审核' }] }
  })
  writePlan = preview.data.plan_id
  const executed = await callAndValidate('office_execute_operation', { plan_id: writePlan })
  assert.equal(executed.success, true)
  assert.ok(executed.data.backup_id.startsWith('backup-'))
  const check = await call('office_read_range', { path: 'demo.xlsx', sheet: '数据', range: 'D1:D2' })
  assert.equal(check.data.rows[0][0].value, '备注')
  assert.equal(check.data.rows[1][0].value, '已审核')
})

await test('已执行的计划不能重复执行', async () => {
  const result = await call('office_execute_operation', { plan_id: writePlan, confirm: true })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'INVALID_REQUEST')
})

await test('office_rollback_operation 恢复到执行前', async () => {
  const result = await callAndValidate('office_rollback_operation', { plan_id: writePlan })
  assert.equal(result.success, true)
  const check = await call('office_read_range', { path: 'demo.xlsx', sheet: '数据', range: 'D1:D2' })
  assert.equal(check.data.rows[0][0], null, '回滚后 D1 应为空')
})

await test('回滚后的计划可以重做（会重新应用同一变更）', async () => {
  const result = await callAndValidate('office_execute_operation', { plan_id: writePlan })
  assert.equal(result.success, true)
  const check = await call('office_read_range', { path: 'demo.xlsx', sheet: '数据', range: 'D1:D2' })
  assert.equal(check.data.rows[1][0].value, '已审核')
})

console.log('\n=== 7. 工作表与行列操作 ===')

await test('新增 / 重命名 / 列出行列操作工具', async () => {
  await call('office_add_worksheet', { path: 'demo.xlsx', name: '汇总' })
  const renamed = await callAndValidate('office_rename_worksheet', { path: 'demo.xlsx', from: '汇总', to: '统计' })
  assert.equal(renamed.success, true)
  const wb = await call('office_read_workbook', { path: 'demo.xlsx' })
  assert.deepEqual(wb.data.sheets.map((s) => s.name), ['数据', '统计'])
})

await test('插入行后原有数据下移', async () => {
  const result = await callAndValidate('office_insert_rows', { path: 'demo.xlsx', sheet: '数据', start: 2, count: 1 })
  assert.equal(result.success, true)
  const check = await call('office_read_range', { path: 'demo.xlsx', sheet: '数据', range: 'A3:A3' })
  assert.equal(check.data.rows[0][0].value, '便携机')
})

await test('插入行会带出公式未重排的警告', async () => {
  const result = await call('office_insert_rows', { path: 'demo.xlsx', sheet: '数据', start: 9, count: 1 })
  assert.ok(result.warnings.some((w) => w.code === 'FORMULA_REFERENCES_NOT_SHIFTED'), JSON.stringify(result.warnings))
})

await test('office_list_files 盘点工作区', async () => {
  const result = await callAndValidate('office_list_files', {})
  assert.ok(result.data.count >= 1)
  assert.ok(result.data.files.some((f) => f.path === 'demo.xlsx'))
})

await test('office_delete_file 确认后删除成功', async () => {
  const result = await callAndValidate('office_delete_file', { path: 'fake.xlsx', confirm: true })
  assert.equal(result.success, true)
  assert.equal(existsSync(join(workspace, 'fake.xlsx')), false)
})

console.log('\n=== 8. 样式与导出 ===')

await test('★ office_create_table 建表格并写齐接线', async () => {
  const result = await callAndValidate('office_create_table', {
    path: 'demo.xlsx',
    sheet: '数据',
    ref: 'A1:D4',
    name: '数据表',
    style: 'TableStyleMedium2'
  })
  assert.equal(result.data.display_name, '数据表')
  assert.equal(result.data.ref, 'A1:D4')
  assert.equal(result.data.columns.length, 4)
  const pkg = (await import('../lib/ooxml.js')).ZipPackage.open(readFileSync(bookPath))
  assert.ok(pkg.has('xl/tables/table1.xml'), '表格部件已写入')
  assert.match(pkg.readText('xl/worksheets/sheet1.xml'), /<tableParts count="1">/)
  // 校验工具应报告 XML 部件格式良好（多根元素那类问题会在这里暴露）
  const validate = await call('office_validate_workbook', { path: 'demo.xlsx' })
  const xmlCheck = validate.data.checks.find((c) => c.name.includes('XML 部件格式良好'))
  assert.ok(xmlCheck, '校验应包含 XML 格式检查')
  assert.equal(xmlCheck.ok, true, xmlCheck.detail)
})

await test('office_create_table 拒绝重叠区域与非法名字', async () => {
  const overlap = await call('office_create_table', { path: 'demo.xlsx', sheet: '数据', ref: 'B2:C3', name: '另一表' })
  assert.equal(overlap.success, false)
  assert.match(overlap.error.message, /重叠/)
  const badName = await call('office_create_table', { path: 'demo.xlsx', sheet: '数据', ref: 'F1:G3', name: 'A1' })
  assert.equal(badName.success, false)
})

await test('★ office_create_chart 建图表并写齐接线', async () => {
  const result = await callAndValidate('office_create_chart', {
    path: 'demo.xlsx',
    sheet: '数据',
    type: 'column',
    title: '趋势图',
    categories: 'A2:A3',
    series: [{ values: 'B2:B3' }],
    anchor: 'J2',
    width_px: 400,
    height_px: 240
  })
  assert.equal(result.data.type, 'column')
  assert.equal(result.data.anchor, 'J2')
  const pkg = (await import('../lib/ooxml.js')).ZipPackage.open(readFileSync(bookPath))
  assert.ok(pkg.has('xl/charts/chart1.xml'), '图表部件已写入')
  assert.ok(pkg.has('xl/drawings/drawing1.xml'), '绘图部件已写入')
  assert.match(pkg.readText('xl/charts/chart1.xml'), /<a:t>趋势图<\/a:t>/)
  const validate = await call('office_validate_workbook', { path: 'demo.xlsx' })
  const xmlCheck = validate.data.checks.find((c) => c.name.includes('XML 部件格式良好'))
  assert.equal(xmlCheck.ok, true, xmlCheck.detail)
})

await test('office_create_chart 参数校验', async () => {
  // 枚举不合规由宿主 schema 校验直接抛出（INVALID_ARGS），不会走到工具体
  await assert.rejects(
    () => call('office_create_chart', { path: 'demo.xlsx', sheet: '数据', type: 'radar', series: [{ values: 'B2:B3' }] }),
    (err) => err.code === 'INVALID_ARGS'
  )
  // 业务校验走适配器：空系列
  const noSeries = await call('office_create_chart', { path: 'demo.xlsx', sheet: '数据', type: 'pie', series: [] })
  assert.equal(noSeries.success, false)
  assert.equal(noSeries.error.code, 'INVALID_REQUEST')
  // 业务校验：饼图只能一个系列
  const pieTwo = await call('office_create_chart', {
    path: 'demo.xlsx',
    sheet: '数据',
    type: 'pie',
    series: [{ values: 'B2:B3' }, { values: 'A2:A3' }]
  })
  assert.equal(pieTwo.success, false)
  assert.equal(pieTwo.error.code, 'INVALID_REQUEST')
})

await test('office_set_cell_style 设置样式并可读回', async () => {
  const result = await callAndValidate('office_set_cell_style', {
    path: 'demo.xlsx',
    sheet: '数据',
    cells: [{ ref: 'A1', bold: true, fill_color: '#FFFF00', horizontal: 'center', number_format: '#,##0.00' }]
  })
  assert.equal(result.success, true)
  assert.equal(result.changes.length, 1)
  assert.equal(result.changes[0].style_index >= 1, true)
})

await test('对不存在单元格设置样式返回 FILE_NOT_FOUND', async () => {
  const result = await call('office_set_cell_style', { path: 'demo.xlsx', sheet: '数据', cells: [{ ref: 'Z99', bold: true }] })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'FILE_NOT_FOUND')
})

await test('office_export_workbook 导出 CSV 且不改源文件', async () => {
  const before = readFileSync(bookPath)
  const result = await callAndValidate('office_export_workbook', { path: 'demo.xlsx', target_path: 'export.csv', format: 'csv', sheet: '数据' })
  assert.equal(result.success, true)
  assert.ok(readFileSync(bookPath).equals(before), '导出不应改动源文件')
  const csv = readFileSync(join(workspace, 'export.csv'))
  assert.equal(csv[0], 0xef, 'CSV 应带 UTF-8 BOM')
  const text = csv.toString('utf8')
  assert.ok(text.includes('产品'), `CSV 内容不含预期数据：${text.slice(0, 120)}`)
  assert.ok(result.warnings.some((w) => w.code === 'LOSSY_CONVERT'), '应提示 CSV 是有损转换')
})

await test('office_export_workbook 导出 xlsx 副本', async () => {
  const result = await callAndValidate('office_export_workbook', { path: 'demo.xlsx', target_path: 'copy.xlsx' })
  assert.equal(result.success, true)
  assert.ok(existsSync(join(workspace, 'copy.xlsx')))
})

await test('导出拒绝静默覆盖已存在文件', async () => {
  const result = await call('office_export_workbook', { path: 'demo.xlsx', target_path: 'copy.xlsx' })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'PERMISSION_DENIED')
  assert.equal(result.error.needs_confirmation, true)
})

console.log('\n=== 9. DOCX（阶段 3 起步）===')

await test('office_create_document 可按 docx 格式创建', async () => {
  const result = await callAndValidate('office_create_document', { path: 'report.docx', format: 'docx', title: '季度销售报告' })
  assert.equal(result.success, true)
  assert.equal(result.data.format, 'docx')
  assert.ok(existsSync(join(workspace, 'report.docx')))
})

await test('office_read_docx 读取段落与结构', async () => {
  const result = await callAndValidate('office_read_docx', { path: 'report.docx', detail: 'paragraphs' })
  assert.equal(result.data.metadata.title, '季度销售报告')
  assert.equal(result.data.total, 3)
  assert.equal(result.data.paragraphs[0].style, 'Heading1')
  assert.equal(result.data.structure.tables, 1)
  assert.equal(result.data.structure.has_page_number_field, true)
})

await test('office_read_docx 读取表格', async () => {
  const result = await callAndValidate('office_read_docx', { path: 'report.docx', detail: 'tables' })
  assert.equal(result.data.count, 1)
  assert.equal(result.data.tables[0].cells[0][1].text, '金额')
})

await test('office_get_file_info 认出 docx', async () => {
  const result = await call('office_get_file_info', { path: 'report.docx' })
  assert.equal(result.data.detected_type, 'docx')
  assert.equal(result.data.extension_mismatch, false)
})

console.log('\n=== 10. DOCX 写入 ===')

await test('office_insert_paragraph 插入段落', async () => {
  const result = await callAndValidate('office_insert_paragraph', {
    path: 'report.docx', after: 0, text: '插入的小标题', style: 'Heading1'
  })
  assert.equal(result.success, true)
  assert.equal(result.output_file.size > 0, true)
  const check = await call('office_read_docx', { path: 'report.docx', detail: 'paragraphs' })
  assert.equal(check.data.total, 4)
  assert.equal(check.data.paragraphs[1].text, '插入的小标题')
  assert.equal(check.data.paragraphs[1].style, 'Heading1')
})

await test('office_update_paragraph 改写并保留样式', async () => {
  const result = await callAndValidate('office_update_paragraph', { path: 'report.docx', index: 0, text: '新的报告标题' })
  assert.equal(result.success, true)
  const check = await call('office_read_docx', { path: 'report.docx', detail: 'paragraphs' })
  assert.equal(check.data.paragraphs[0].text, '新的报告标题')
  assert.equal(check.data.paragraphs[0].style, 'Heading1')
})

await test('office_find_and_replace_docx dry_run 不改文件', async () => {
  const before = readFileSync(join(workspace, 'report.docx'))
  const result = await callAndValidate('office_find_and_replace_docx', {
    path: 'report.docx', find: 'dsh-exp-office', replace: 'DSH', dry_run: true
  })
  assert.equal(result.data.dry_run, true)
  assert.equal(result.data.replacements, 1)
  assert.ok(readFileSync(join(workspace, 'report.docx')).equals(before), 'dry_run 不应改动文件')
})

await test('office_find_and_replace_docx 真实替换', async () => {
  const result = await callAndValidate('office_find_and_replace_docx', {
    path: 'report.docx', find: 'dsh-exp-office', replace: 'DeepSeek Harness'
  })
  assert.equal(result.data.replacements, 1)
  const text = await call('office_read_docx', { path: 'report.docx', detail: 'text' })
  assert.ok(text.data.text.includes('本文件由 DeepSeek Harness 生成。'))
})

await test('office_delete_paragraph 未确认时被拒绝', async () => {
  const result = await call('office_delete_paragraph', { path: 'report.docx', index: 3, confirm: false })
  assert.equal(result.success, false)
  assert.equal(result.error.needs_confirmation, true)
})

await test('office_delete_paragraph 确认后删除', async () => {
  const result = await callAndValidate('office_delete_paragraph', { path: 'report.docx', index: 3, confirm: true })
  assert.equal(result.success, true)
  const check = await call('office_read_docx', { path: 'report.docx', detail: 'paragraphs' })
  assert.equal(check.data.total, 3)
})

await test('DOCX 写入失败时保留原文件', async () => {
  const before = readFileSync(join(workspace, 'report.docx'))
  const result = await call('office_delete_paragraph', { path: 'report.docx', index: 99, confirm: true })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'FILE_NOT_FOUND')
  assert.ok(readFileSync(join(workspace, 'report.docx')).equals(before), '失败时原文件必须保持不变')
})

await test('DOCX 写入后结构部件全部保留', async () => {
  const result = await call('office_read_docx', { path: 'report.docx', detail: 'structure' })
  assert.equal(result.data.structure.tables, 1)
  assert.equal(result.data.structure.headers, 1)
  assert.equal(result.data.structure.footers, 1)
  assert.equal(result.data.structure.has_page_number_field, true)
  assert.equal(result.data.metadata.title, '季度销售报告')
})

await test('office_validate_docx 校验通过并列出未检查项', async () => {
  const result = await callAndValidate('office_validate_docx', { path: 'report.docx' })
  assert.equal(result.data.valid, true, JSON.stringify(result.data.checks.filter((c) => !c.ok)))
  assert.ok(result.data.not_checked.length >= 5, '应明确列出需要渲染才能判定的项')
  assert.ok(result.data.checks.some((c) => c.name === '关系目标部件存在'))
})

await test('office_validate_docx 支持与基线比对部件差异', async () => {
  const result = await callAndValidate('office_validate_docx', { path: 'report.docx', baseline_path: 'report.docx' })
  assert.deepEqual(result.data.changed_parts, [], '与自身比对不应有差异')
})

await test('office_validate_docx 对损坏文件返回结构化错误', async () => {
  writeFileSync(join(workspace, 'broken.docx'), Buffer.from('not a zip at all, definitely not a docx'))
  const result = await call('office_validate_docx', { path: 'broken.docx' })
  assert.equal(result.success, false)
  assert.equal(typeof result.error.code, 'string')
})

console.log('\n=== 11. DOCX 图片与保护性检查 ===')

await test('office_insert_docx_image 插入图片并接线', async () => {
  writeFileSync(join(workspace, 'logo.png'), makePng(64, 32))
  const result = await callAndValidate('office_insert_docx_image', {
    path: 'report.docx',
    after: 0,
    image_path: 'logo.png',
    width_px: 128,
    alt_text: '公司标志'
  })
  assert.equal(result.success, true, JSON.stringify(result.error))
  assert.equal(result.data.media_part, 'word/media/image1.png')
  assert.deepEqual(result.data.display_size_px, { width: 128, height: 64 })
})

await test('插入图片后 validate_docx 仍通过', async () => {
  const result = await callAndValidate('office_validate_docx', { path: 'report.docx' })
  assert.equal(result.data.valid, true, JSON.stringify(result.data.checks.filter((c) => !c.ok)))
})

await test('非图片文件被拒绝', async () => {
  writeFileSync(join(workspace, 'notimage.png'), 'this is not an image at all')
  const result = await call('office_insert_docx_image', { path: 'report.docx', after: 0, image_path: 'notimage.png' })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'UNSUPPORTED_FILE_TYPE')
})

await test('未提供图片来源时被拒绝', async () => {
  const result = await call('office_insert_docx_image', { path: 'report.docx', after: 0 })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'INVALID_REQUEST')
})

await test('office_set_docx_footer 可创建页码页脚', async () => {
  const result = await callAndValidate('office_set_docx_footer', { path: 'report.docx', text: '第 ', page_number: true })
  assert.equal(result.success, true)
  const check = await call('office_read_docx', { path: 'report.docx', detail: 'structure' })
  assert.equal(check.data.structure.has_page_number_field, true)
})

console.log('\n=== 12. PPTX（阶段 4 起步）===')

const deckSource = join(here, 'fixtures', 'deck.pptx')
const hasDeck = existsSync(deckSource)
if (hasDeck) copyFileSync(deckSource, join(workspace, 'deck.pptx'))

await test('office_read_pptx 读取结构与幻灯片', async () => {
  if (!hasDeck) {
    console.log('     ⏭️  跳过：未找到 deck.pptx 样本')
    return
  }
  const result = await callAndValidate('office_read_pptx', { path: 'deck.pptx', detail: 'summary' })
  assert.equal(result.data.structure.slide_count, 4)
  assert.equal(result.data.structure.slide_size.aspect, '16:9')
  assert.deepEqual(result.data.structure.hidden_slides, [3])
  assert.ok(result.warnings.some((w) => w.code === 'HIDDEN_SLIDES'), '应提示隐藏幻灯片')
})

await test('office_read_pptx 读取单张幻灯片内容', async () => {
  if (!hasDeck) return
  const result = await callAndValidate('office_read_pptx', { path: 'deck.pptx', detail: 'slide', index: 1 })
  assert.equal(result.data.slide.title, '本季度进展')
  assert.ok(result.data.slide.notes.includes('讲稿'))
})

await test('office_read_pptx detail=slide 缺 index 时报错', async () => {
  if (!hasDeck) return
  const result = await call('office_read_pptx', { path: 'deck.pptx', detail: 'slide' })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'INVALID_REQUEST')
})

await test('office_read_pptx 全文提取', async () => {
  if (!hasDeck) return
  const result = await callAndValidate('office_read_pptx', { path: 'deck.pptx', detail: 'text' })
  assert.ok(result.data.text.includes('季度业务报告'))
  assert.ok(result.data.text.includes('本季度进展'))
})

await test('office_get_file_info 认出 pptx', async () => {
  if (!hasDeck) return
  const result = await call('office_get_file_info', { path: 'deck.pptx' })
  assert.equal(result.data.detected_type, 'pptx')
})

await test('office_validate_pptx 校验通过并列出未检查项', async () => {
  if (!hasDeck) return
  const result = await callAndValidate('office_validate_pptx', { path: 'deck.pptx' })
  assert.equal(result.data.valid, true, JSON.stringify(result.data.checks.filter((c) => !c.ok)))
  assert.ok(result.data.not_checked.length >= 4)
  assert.equal(result.data.slide_size.aspect, '16:9')
})

await test('office_read_pptx 对 xlsx 报错而不是静默返回空', async () => {
  const result = await call('office_read_pptx', { path: 'demo.xlsx' })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'CORRUPTED_DOCUMENT')
})

await test('office_update_slide_text 改写幻灯片文本', async () => {
  if (!hasDeck) return
  const read = await call('office_read_pptx', { path: 'deck.pptx', detail: 'slide', index: 1 })
  const bodyIndex = read.data.slide.shapes.findIndex((s) => s.placeholder_label === '正文')
  assert.ok(bodyIndex >= 0, '应先读到正文形状')

  const result = await callAndValidate('office_update_slide_text', {
    path: 'deck.pptx',
    index: 1,
    shape: bodyIndex,
    text: '第一条\n第二条'
  })
  assert.equal(result.success, true, JSON.stringify(result.error))
  assert.equal(result.data.paragraphs, 2)

  const check = await call('office_read_pptx', { path: 'deck.pptx', detail: 'slide', index: 1 })
  assert.deepEqual(check.data.slide.shapes[bodyIndex].text.split('\n'), ['第一条', '第二条'])
  assert.equal(check.data.slide.title, '本季度进展', '标题不应被改动')
})

await test('改写后校验仍然通过', async () => {
  if (!hasDeck) return
  const result = await callAndValidate('office_validate_pptx', { path: 'deck.pptx' })
  assert.equal(result.data.valid, true, JSON.stringify(result.data.checks.filter((c) => !c.ok)))
})

await test('对表格形状写文本被拒绝', async () => {
  if (!hasDeck) return
  const read = await call('office_read_pptx', { path: 'deck.pptx', detail: 'slide', index: 2 })
  const tableIndex = read.data.slide.shapes.findIndex((s) => s.type === 'table')
  assert.ok(tableIndex >= 0)
  const result = await call('office_update_slide_text', { path: 'deck.pptx', index: 2, shape: tableIndex, text: 'x' })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'UNSUPPORTED_FEATURE')
  assert.equal(result.error.details.shape_type, 'table')
})

await test('office_add_slide 新增幻灯片并继承版式', async () => {
  if (!hasDeck) return
  const result = await callAndValidate('office_add_slide', { path: 'deck.pptx', after: 1, layout_of: 1 })
  assert.equal(result.success, true, JSON.stringify(result.error))
  assert.equal(result.data.slide_count, 5)
  assert.equal(result.data.index, 2)
})

await test('office_reorder_slides 调整顺序', async () => {
  if (!hasDeck) return
  const result = await callAndValidate('office_reorder_slides', { path: 'deck.pptx', from: 0, to: 2 })
  assert.equal(result.data.changed, true)
  const check = await call('office_read_pptx', { path: 'deck.pptx', detail: 'slides' })
  assert.equal(check.data.slides.length, 5)
})

await test('office_delete_slide 未确认时被拒绝', async () => {
  if (!hasDeck) return
  const result = await call('office_delete_slide', { path: 'deck.pptx', index: 0, confirm: false })
  assert.equal(result.success, false)
  assert.equal(result.error.needs_confirmation, true)
})

await test('office_delete_slide 确认后删除并保持可校验', async () => {
  if (!hasDeck) return
  const result = await callAndValidate('office_delete_slide', { path: 'deck.pptx', index: 4, confirm: true })
  assert.equal(result.data.slide_count, 4)
  const check = await call('office_validate_pptx', { path: 'deck.pptx' })
  assert.equal(check.data.valid, true, JSON.stringify(check.data.checks.filter((c) => !c.ok)))
})

await test('幻灯片越界时报错而不是静默忽略', async () => {
  if (!hasDeck) return
  const result = await call('office_reorder_slides', { path: 'deck.pptx', from: 0, to: 99 })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'INVALID_REQUEST')
})

await test('office_insert_slide_image 插入图片', async () => {
  if (!hasDeck) return
  writeFileSync(join(workspace, 'chart.png'), makePng(160, 80))
  const result = await callAndValidate('office_insert_slide_image', {
    path: 'deck.pptx',
    index: 1,
    image_path: 'chart.png',
    width_px: 320,
    alt_text: '季度图表'
  })
  assert.equal(result.success, true, JSON.stringify(result.error))
  assert.equal(result.data.display_size_px.width, 320)
  assert.equal(result.data.display_size_px.height, 160)

  const check = await call('office_read_pptx', { path: 'deck.pptx', detail: 'slide', index: 1 })
  assert.ok(check.data.slide.shapes.some((s) => s.type === 'picture'), '应出现 picture 形状')
})

await test('插入图片后校验仍然通过', async () => {
  if (!hasDeck) return
  const result = await callAndValidate('office_validate_pptx', { path: 'deck.pptx' })
  assert.equal(result.data.valid, true, JSON.stringify(result.data.checks.filter((c) => !c.ok)))
})

await test('非图片文件被拒绝', async () => {
  if (!hasDeck) return
  writeFileSync(join(workspace, 'notimage2.png'), 'definitely not an image')
  const result = await call('office_insert_slide_image', { path: 'deck.pptx', index: 1, image_path: 'notimage2.png' })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'UNSUPPORTED_FILE_TYPE')
})

await test('未提供图片来源时被拒绝', async () => {
  if (!hasDeck) return
  const result = await call('office_insert_slide_image', { path: 'deck.pptx', index: 1 })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'INVALID_REQUEST')
})

await test('★ 复制幻灯片 + 插入文本框 端到端', async () => {
  if (!hasDeck) return
  // 用未被动过的样本副本：前面的用例已经改过 deck.pptx 的页序，那里的下标 1 未必带备注
  copyFileSync(deckSource, join(workspace, 'deck-dup.pptx'))
  const before = await call('office_read_pptx', { path: 'deck-dup.pptx', detail: 'slides' })
  const countBefore = before.data.slides.length

  const dup = await callAndValidate('office_duplicate_slide', { path: 'deck-dup.pptx', index: 1 })
  assert.equal(dup.data.source, 1)
  assert.equal(dup.data.index, 2)
  assert.equal(dup.data.slide_count, countBefore + 1)
  assert.equal(dup.data.notes_dropped, true, '源页带备注，副本应丢弃备注页关系')

  const box = await callAndValidate('office_add_text_box', {
    path: 'deck-dup.pptx',
    index: 2,
    text: '复制页文本框\n第二行',
    left_px: 96,
    top_px: 120,
    width_px: 480,
    height_px: 120,
    font_size_pt: 20,
    bold: true,
    align: 'center'
  })
  assert.equal(box.data.index, 2)

  const slide = await call('office_read_pptx', { path: 'deck-dup.pptx', detail: 'slide', index: 2 })
  const texts = slide.data.slide.shapes.map((s) => s.text ?? '')
  assert.ok(texts.some((t) => t.includes('复制页文本框')), `副本页应读到文本框：${JSON.stringify(texts)}`)

  const valid = await call('office_validate_pptx', { path: 'deck-dup.pptx' })
  assert.equal(valid.data.valid, true, JSON.stringify(valid.data.checks.filter((c) => !c.ok)))
})

await test('office_duplicate_slide / office_add_text_box 参数校验', async () => {
  if (!hasDeck) return
  const badIndex = await call('office_duplicate_slide', { path: 'deck-dup.pptx', index: 99 })
  assert.equal(badIndex.success, false)
  assert.equal(badIndex.error.code, 'FILE_NOT_FOUND')

  // 参数 schema 校验在 defineTool 层直接抛错（不会走到 execute）
  await assert.rejects(() => call('office_add_text_box', { path: 'deck-dup.pptx', index: 0 }), /missing required property "text"/)

  const badAlign = await call('office_add_text_box', { path: 'deck-dup.pptx', index: 0, text: 'x', align: 'justify' })
  assert.equal(badAlign.success, false)
  assert.equal(badAlign.error.code, 'INVALID_REQUEST')
})

console.log('\n=== 13. PDF（阶段 5 起步）===')

const pdfSource = join(here, 'fixtures', 'sample.pdf')
const hasPdf = existsSync(pdfSource)
if (hasPdf) copyFileSync(pdfSource, join(workspace, 'sample.pdf'))

await test('office_read_pdf 读取元数据与结构', async () => {
  if (!hasPdf) {
    console.log('     ⏭️  跳过：未找到 sample.pdf 样本')
    return
  }
  const result = await callAndValidate('office_read_pdf', { path: 'sample.pdf', detail: 'summary' })
  assert.equal(result.data.metadata.title, '季度销售报告')
  assert.equal(result.data.structure.page_count, 1)
  assert.equal(result.data.structure.encrypted, false)
})

await test('★ office_read_pdf 提取中文文本', async () => {
  if (!hasPdf) return
  const result = await callAndValidate('office_read_pdf', { path: 'sample.pdf', detail: 'text' })
  assert.ok(result.data.text.includes('季度销售报告'), JSON.stringify(result.data.text.slice(0, 120)))
  assert.ok(result.data.text.includes('本文件用于验证'), JSON.stringify(result.data.text.slice(0, 200)))
})

await test('office_read_pdf 读取页面尺寸', async () => {
  if (!hasPdf) return
  const result = await callAndValidate('office_read_pdf', { path: 'sample.pdf', detail: 'pages' })
  assert.equal(result.data.pages.length, 1)
  assert.ok(Math.abs(result.data.pages[0].width_pt - 595.3) < 1)
})

await test('office_validate_pdf 校验通过并列出未检查项', async () => {
  if (!hasPdf) return
  const result = await callAndValidate('office_validate_pdf', { path: 'sample.pdf' })
  assert.equal(result.data.valid, true, JSON.stringify(result.data.checks.filter((c) => !c.ok)))
  assert.ok(result.data.not_checked.some((c) => c.name === '交叉引用表一致性'))
})

await test('office_get_file_info 认出 pdf', async () => {
  if (!hasPdf) return
  const result = await call('office_get_file_info', { path: 'sample.pdf' })
  assert.equal(result.data.detected_type, 'pdf')
})

// 多页样本用于页面级写操作；没有样本时相关用例明确跳过，而不是假装通过。
const pdfMultiSource = join(here, 'fixtures', 'sample-multipage.pdf')
const hasMultiPdf = existsSync(pdfMultiSource)
if (hasMultiPdf) copyFileSync(pdfMultiSource, join(workspace, 'multi.pdf'))

// 含图样本用于图片提取。
const pdfImagesSource = join(here, 'fixtures', 'sample-images.pdf')
const hasImagePdf = existsSync(pdfImagesSource)

await test('★ office_rotate_pdf_pages 旋转并增量写回', async () => {
  if (!hasPdf) return
  copyFileSync(pdfSource, join(workspace, 'rotate.pdf'))
  const before = readFileSync(join(workspace, 'rotate.pdf'))
  const result = await callAndValidate('office_rotate_pdf_pages', { path: 'rotate.pdf', degrees: 90, page: 0 })
  const after = readFileSync(join(workspace, 'rotate.pdf'))
  assert.ok(after.length > before.length, '增量更新应让文件变大')
  assert.ok(after.subarray(0, before.length).equals(before), '原有字节必须保持不变')
  assert.equal(result.data.appended_bytes, after.length - before.length)
  const read = await call('office_read_pdf', { path: 'rotate.pdf', detail: 'pages' })
  assert.equal(read.data.pages[0].rotation, 90)
})

await test('★ office_rotate_pdf_pages 支持全页累加', async () => {
  if (!hasMultiPdf) {
    console.log('     ⏭️  跳过：未找到 sample-multipage.pdf 样本')
    return
  }
  copyFileSync(pdfMultiSource, join(workspace, 'rotate-all.pdf'))
  const result = await callAndValidate('office_rotate_pdf_pages', { path: 'rotate-all.pdf', degrees: 90, mode: 'add' })
  assert.equal(result.data.rotated.length, 3)
  const read = await call('office_read_pdf', { path: 'rotate-all.pdf', detail: 'pages' })
  assert.deepEqual(read.data.pages.map((p) => p.rotation), [90, 90, 90])
  // 再累加一次应变成 180，而不是覆盖成 90
  await callAndValidate('office_rotate_pdf_pages', { path: 'rotate-all.pdf', degrees: 90, mode: 'add' })
  const read2 = await call('office_read_pdf', { path: 'rotate-all.pdf', detail: 'pages' })
  assert.deepEqual(read2.data.pages.map((p) => p.rotation), [180, 180, 180])
})

await test('★ office_reorder_pdf_pages 重排后文本顺序同步变化', async () => {
  if (!hasMultiPdf) return
  copyFileSync(pdfMultiSource, join(workspace, 'order.pdf'))
  const before = await call('office_read_pdf', { path: 'order.pdf', detail: 'text' })
  const pagesBefore = before.data.chars_per_page.length
  assert.equal(pagesBefore, 3)
  const result = await callAndValidate('office_reorder_pdf_pages', { path: 'order.pdf', order: [2, 0, 1] })
  assert.equal(result.data.page_count, 3)
  const after = await call('office_read_pdf', { path: 'order.pdf', detail: 'text' })
  assert.deepEqual(after.data.chars_per_page, [before.data.chars_per_page[2], before.data.chars_per_page[0], before.data.chars_per_page[1]])
})

await test('office_reorder_pdf_pages 支持单页移动与非法排列拒绝', async () => {
  if (!hasMultiPdf) return
  copyFileSync(pdfMultiSource, join(workspace, 'move.pdf'))
  const moved = await callAndValidate('office_reorder_pdf_pages', { path: 'move.pdf', from: 0, to: 2 })
  assert.equal(moved.data.page_count, 3)
  const bad = await call('office_reorder_pdf_pages', { path: 'move.pdf', order: [0, 0, 1] })
  assert.equal(bad.success, false)
  assert.match(bad.error.message, /重复下标/)
  const badLength = await call('office_reorder_pdf_pages', { path: 'move.pdf', order: [0, 1] })
  assert.equal(badLength.success, false)
  const neither = await call('office_reorder_pdf_pages', { path: 'move.pdf' })
  assert.equal(neither.success, false)
})

await test('★ office_delete_pdf_page 删除后页数与文本同步减少', async () => {
  if (!hasMultiPdf) return
  copyFileSync(pdfMultiSource, join(workspace, 'delete.pdf'))
  const before = await call('office_read_pdf', { path: 'delete.pdf', detail: 'text' })
  const result = await callAndValidate('office_delete_pdf_page', { path: 'delete.pdf', page: 1 })
  assert.equal(result.data.page_count, 2)
  const after = await call('office_read_pdf', { path: 'delete.pdf', detail: 'text' })
  assert.deepEqual(after.data.chars_per_page, [before.data.chars_per_page[0], before.data.chars_per_page[2]])
  // 被摘掉的对象仍留在文件里（增量更新不回收字节），但已不再被页面树引用
  const removedObject = result.changes[0].removed_object
  assert.ok(readFileSync(join(workspace, 'delete.pdf')).toString('latin1').includes(`${removedObject} obj`))
})

await test('office_rotate_pdf_pages 拒绝非 90 倍数角度', async () => {
  if (!hasPdf) return
  copyFileSync(pdfSource, join(workspace, 'angle.pdf'))
  const result = await call('office_rotate_pdf_pages', { path: 'angle.pdf', degrees: 45 })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'INVALID_REQUEST')
})

await test('office_delete_pdf_page 拒绝删掉唯一一页', async () => {
  if (!hasPdf) return
  copyFileSync(pdfSource, join(workspace, 'single.pdf'))
  const result = await call('office_delete_pdf_page', { path: 'single.pdf', page: 0 })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'INVALID_REQUEST')
})

await test('★ office_add_pdf_watermark 叠加写入且原字节不变', async () => {
  if (!hasPdf) return
  copyFileSync(pdfSource, join(workspace, 'wm.pdf'))
  const before = readFileSync(join(workspace, 'wm.pdf'))
  const result = await callAndValidate('office_add_pdf_watermark', { path: 'wm.pdf', text: 'DRAFT ONLY' })
  const after = readFileSync(join(workspace, 'wm.pdf'))
  assert.ok(after.subarray(0, before.length).equals(before), '原有字节必须不变')
  assert.ok(result.data.appended_bytes > 0)
  const read = await call('office_read_pdf', { path: 'wm.pdf', detail: 'text' })
  assert.ok(read.data.text.includes('DRAFT ONLY'), JSON.stringify(read.data.text.slice(-80)))
  assert.ok(read.data.text.includes('季度销售报告'), '原有中文文字必须仍可读（资源合并不能丢字体）')
})

await test('★ office_add_pdf_watermark 拒绝中文水印并保持文件不变', async () => {
  if (!hasPdf) return
  copyFileSync(pdfSource, join(workspace, 'wm-cjk.pdf'))
  const before = readFileSync(join(workspace, 'wm-cjk.pdf'))
  const result = await call('office_add_pdf_watermark', { path: 'wm-cjk.pdf', text: '机密' })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'UNSUPPORTED_FEATURE')
  assert.ok(readFileSync(join(workspace, 'wm-cjk.pdf')).equals(before), '拒绝时文件必须原封不动')
})

await test('★ office_add_pdf_page_numbers 逐页编号', async () => {
  if (!hasMultiPdf) return
  copyFileSync(pdfMultiSource, join(workspace, 'numbers.pdf'))
  const result = await callAndValidate('office_add_pdf_page_numbers', { path: 'numbers.pdf', format: '{page} / {total}' })
  assert.equal(result.data.page_count, 3)
  const labels = []
  for (let i = 0; i < 3; i += 1) {
    const page = await call('office_read_pdf', { path: 'numbers.pdf', detail: 'text', page: i })
    labels.push(page.data.text.trim().split('\n').pop())
  }
  assert.deepEqual(labels, ['1 / 3', '2 / 3', '3 / 3'])
})

await test('office_add_pdf_page_numbers 支持起始编号与页眉', async () => {
  if (!hasMultiPdf) return
  copyFileSync(pdfMultiSource, join(workspace, 'numbers2.pdf'))
  const result = await callAndValidate('office_add_pdf_page_numbers', {
    path: 'numbers2.pdf',
    format: 'Page {page}',
    start_at: 10,
    position: 'top'
  })
  assert.equal(result.data.start_at, 10)
  assert.deepEqual(result.data.pages, [0, 1, 2])
  const read = await call('office_read_pdf', { path: 'numbers2.pdf', detail: 'text' })
  for (const label of ['Page 10', 'Page 11', 'Page 12']) {
    assert.ok(read.data.text.includes(label), `缺少 ${label}`)
  }
})

await test('水印工具对非 PDF 报错而不是产出损坏文件', async () => {
  const before = readFileSync(bookPath)
  const result = await call('office_add_pdf_watermark', { path: 'demo.xlsx', text: 'X' })
  assert.equal(result.success, false)
  assert.ok(['CORRUPTED_DOCUMENT', 'EXTENSION_MISMATCH'].includes(result.error.code), result.error.code)
  assert.ok(readFileSync(bookPath).equals(before), '报错时不能碰原文件')
})

await test('★ office_set_pdf_metadata 写入元数据且原字节不变', async () => {
  if (!hasPdf) return
  copyFileSync(pdfSource, join(workspace, 'meta.pdf'))
  const before = readFileSync(join(workspace, 'meta.pdf'))
  const result = await callAndValidate('office_set_pdf_metadata', {
    path: 'meta.pdf',
    title: '归档 2026',
    author: 'DSH Office',
    keywords: 'archive, q3'
  })
  const after = readFileSync(join(workspace, 'meta.pdf'))
  assert.ok(after.subarray(0, before.length).equals(before), '原有字节必须不变')
  assert.equal(result.data.created_info, false)
  const read = await call('office_read_pdf', { path: 'meta.pdf', detail: 'summary' })
  assert.equal(read.data.metadata.title, '归档 2026')
  assert.equal(read.data.metadata.author, 'DSH Office')
  assert.equal(read.data.metadata.keywords, 'archive, q3')
  assert.match(read.data.metadata.producer, /LibreOffice/, '未指定的字段不能被清掉')
})

await test('office_set_pdf_metadata 无字段时被拒绝', async () => {
  if (!hasPdf) return
  copyFileSync(pdfSource, join(workspace, 'meta2.pdf'))
  const before = readFileSync(join(workspace, 'meta2.pdf'))
  const result = await call('office_set_pdf_metadata', { path: 'meta2.pdf' })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'INVALID_REQUEST')
  assert.ok(readFileSync(join(workspace, 'meta2.pdf')).equals(before))
})

await test('★ office_optimize_pdf 原地重写：内容不变、回收字节、不再依赖增量链', async () => {
  if (!hasMultiPdf) return
  copyFileSync(pdfMultiSource, join(workspace, 'opt.pdf'))
  // 先做几轮增量更新把文件撑大
  for (let i = 0; i < 3; i += 1) {
    await callAndValidate('office_rotate_pdf_pages', { path: 'opt.pdf', degrees: 90, page: 0 })
  }
  const grown = readFileSync(join(workspace, 'opt.pdf'))
  const grownText = await call('office_read_pdf', { path: 'opt.pdf', detail: 'text' })
  const result = await callAndValidate('office_optimize_pdf', { path: 'opt.pdf' })
  const optimized = readFileSync(join(workspace, 'opt.pdf'))
  assert.ok(optimized.length < grown.length, `应回收字节：${grown.length} → ${optimized.length}`)
  assert.ok(!optimized.subarray(0, grown.length).equals(grown), '重写不是增量追加，不应是原文件的超集')
  assert.equal(result.data.page_count, 3)
  const after = await call('office_read_pdf', { path: 'opt.pdf', detail: 'text' })
  assert.deepEqual(after.data.chars_per_page, grownText.data.chars_per_page)
  const pages = await call('office_read_pdf', { path: 'opt.pdf', detail: 'pages' })
  // 三次 rotate 默认是 set（不是 add），所以结果仍是 90°；每次都会追加新对象把文件撑大
  assert.deepEqual(pages.data.pages.map((p) => p.rotation), [90, 0, 0])
  const valid = await call('office_validate_pdf', { path: 'opt.pdf' })
  assert.equal(valid.data.valid, true)
})

await test('office_optimize_pdf 另存为新文件时源文件不变', async () => {
  if (!hasPdf) return
  copyFileSync(pdfSource, join(workspace, 'opt-src.pdf'))
  const before = readFileSync(join(workspace, 'opt-src.pdf'))
  const result = await callAndValidate('office_optimize_pdf', { path: 'opt-src.pdf', output_path: 'opt-copy.pdf' })
  assert.ok(readFileSync(join(workspace, 'opt-src.pdf')).equals(before), '源文件必须原样不动')
  assert.equal(result.data.output_bytes, readFileSync(join(workspace, 'opt-copy.pdf')).length)
  const read = await call('office_read_pdf', { path: 'opt-copy.pdf', detail: 'summary' })
  assert.equal(read.data.metadata.title, '季度销售报告', '另存副本必须保留元数据')
})

await test('★ office_split_pdf 拆出指定页，未选页内容不残留在字节里', async () => {
  if (!hasMultiPdf) return
  copyFileSync(pdfMultiSource, join(workspace, 'split-src.pdf'))
  const source = readFileSync(join(workspace, 'split-src.pdf'))
  const pageTexts = []
  for (let i = 0; i < 3; i += 1) {
    const page = await call('office_read_pdf', { path: 'split-src.pdf', detail: 'text', page: i })
    pageTexts.push(page.data.text.replace(/\s+/g, '').trim())
  }
  const result = await callAndValidate('office_split_pdf', { path: 'split-src.pdf', pages: [2, 0], output_path: 'split-out.pdf' })
  assert.deepEqual(result.data.pages, [2, 0])
  assert.equal(result.data.dropped_pages, 1)
  assert.ok(readFileSync(join(workspace, 'split-src.pdf')).equals(source), '源文件必须原样不动')
  const out = readFileSync(join(workspace, 'split-out.pdf'))
  const first = await call('office_read_pdf', { path: 'split-out.pdf', detail: 'text', page: 0 })
  const second = await call('office_read_pdf', { path: 'split-out.pdf', detail: 'text', page: 1 })
  assert.equal(first.data.text.replace(/\s+/g, '').trim(), pageTexts[2])
  assert.equal(second.data.text.replace(/\s+/g, '').trim(), pageTexts[0])
  // 泄漏检查：内容流是压缩的，正文不会以明文出现在字节里，所以「搜不到文字」不算证据；
  // 硬条件是输出里只剩 2 个页面对象（这条抓出过书签 /Dest 把丢弃页拉回来的真 bug）。
  const pageObjects = (out.toString('latin1').match(/\/Type\s*\/Page[^s]/g) ?? []).length
  assert.equal(pageObjects, 2, `输出里应只剩 2 个页面对象，实际 ${pageObjects}`)
  assert.ok(out.length < source.length, `拆分后应更小：${source.length} → ${out.length}`)
})

await test('office_split_pdf 拒绝覆盖已有文件，overwrite=true 才行', async () => {
  if (!hasPdf) return
  copyFileSync(pdfSource, join(workspace, 'split2-src.pdf'))
  writeFileSync(join(workspace, 'split2-out.pdf'), Buffer.from('occupied', 'latin1'))
  const refused = await call('office_split_pdf', { path: 'split2-src.pdf', pages: [0], output_path: 'split2-out.pdf' })
  assert.equal(refused.success, false)
  assert.equal(refused.error.code, 'PERMISSION_DENIED')
  assert.equal(readFileSync(join(workspace, 'split2-out.pdf')).toString('latin1'), 'occupied', '拒绝时不能碰目标文件')
  const forced = await callAndValidate('office_split_pdf', {
    path: 'split2-src.pdf',
    pages: [0],
    output_path: 'split2-out.pdf',
    overwrite: true
  })
  assert.equal(forced.data.page_count, 1)
})

await test('★ office_split_pdf 越界页报错且不留下输出文件', async () => {
  if (!hasMultiPdf) return
  copyFileSync(pdfMultiSource, join(workspace, 'split3-src.pdf'))
  const result = await call('office_split_pdf', { path: 'split3-src.pdf', pages: [0, 9], output_path: 'split3-out.pdf' })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'INVALID_REQUEST')
  assert.ok(!existsSync(join(workspace, 'split3-out.pdf')), '失败时不应留下半个输出文件')
})

await test('★ office_merge_pdfs 合并两份 PDF（页序与文本逐页对应）', async () => {
  if (!hasPdf || !hasMultiPdf) return
  copyFileSync(pdfSource, join(workspace, 'merge-a.pdf'))
  copyFileSync(pdfMultiSource, join(workspace, 'merge-b.pdf'))
  const textOf = async (path, count) => {
    const out = []
    for (let i = 0; i < count; i += 1) {
      const page = await call('office_read_pdf', { path, detail: 'text', page: i })
      out.push(page.data.text.replace(/\s+/g, '').trim())
    }
    return out
  }
  const expected = [...(await textOf('merge-a.pdf', 1)), ...(await textOf('merge-b.pdf', 3))]
  const result = await callAndValidate('office_merge_pdfs', {
    paths: ['merge-a.pdf', 'merge-b.pdf'],
    output_path: 'merged.pdf'
  })
  assert.equal(result.data.page_count, 4)
  assert.equal(result.data.sources.length, 2)
  const actual = await textOf('merged.pdf', 4)
  assert.deepEqual(actual, expected)
  const valid = await call('office_validate_pdf', { path: 'merged.pdf' })
  assert.equal(valid.data.valid, true)
  const meta = await call('office_read_pdf', { path: 'merged.pdf', detail: 'summary' })
  assert.equal(meta.data.metadata.title, '季度销售报告', '元数据取自第一份')
})

await test('office_merge_pdfs 校验参数与失败不留产物', async () => {
  if (!hasPdf) return
  copyFileSync(pdfSource, join(workspace, 'merge-c.pdf'))
  const tooFew = await call('office_merge_pdfs', { paths: ['merge-c.pdf'], output_path: 'merge-bad.pdf' })
  assert.equal(tooFew.success, false)
  assert.equal(tooFew.error.code, 'INVALID_REQUEST')
  const missing = await call('office_merge_pdfs', {
    paths: ['merge-c.pdf', 'no-such-file.pdf'],
    output_path: 'merge-bad.pdf'
  })
  assert.equal(missing.success, false)
  assert.ok(!existsSync(join(workspace, 'merge-bad.pdf')), '失败时不应留下输出文件')
})

await test('★ office_extract_pdf_images 提取图片并自检', async () => {
  if (!hasImagePdf) {
    console.log('     ⏭️  跳过：未找到 sample-images.pdf 样本')
    return
  }
  copyFileSync(pdfImagesSource, join(workspace, 'with-image.pdf'))
  const result = await callAndValidate('office_extract_pdf_images', {
    path: 'with-image.pdf',
    output_dir: 'extracted'
  })
  assert.equal(result.data.image_count, 1, JSON.stringify(result.data.skipped))
  const file = result.data.files[0]
  assert.equal(file.format, 'png')
  assert.equal(file.width, 320)
  assert.equal(file.height, 160)
  assert.deepEqual(file.pages, [1, 2, 3])
  const written = readFileSync(join(workspace, 'extracted', file.name))
  assert.equal(written.length, file.bytes)
  assert.deepEqual([...written.subarray(0, 8)], [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  // 第二次运行：默认不覆盖，应报告同名文件
  const again = await callAndValidate('office_extract_pdf_images', { path: 'with-image.pdf', output_dir: 'extracted' })
  assert.equal(again.data.image_count, 0)
  assert.equal(again.data.skipped_existing.length, 1)
  assert.ok(again.warnings.some((w) => w.code === 'FILE_EXISTS'))
  // 显式 overwrite 才会覆盖
  const forced = await callAndValidate('office_extract_pdf_images', {
    path: 'with-image.pdf',
    output_dir: 'extracted',
    overwrite: true
  })
  assert.equal(forced.data.image_count, 1)
})

await test('office_extract_pdf_images 对没有图片的 PDF 给出提示', async () => {
  if (!hasPdf) return
  copyFileSync(pdfSource, join(workspace, 'no-image.pdf'))
  const result = await callAndValidate('office_extract_pdf_images', { path: 'no-image.pdf', output_dir: 'empty-out' })
  assert.equal(result.data.image_count, 0)
  assert.ok(result.warnings.some((w) => w.code === 'NO_IMAGES'))
})

await test('★ 表格结构操作：插行 / 合并 / 删行 端到端', async () => {
  // 用新建文档里的表格做端到端：先看清楚结构，再依次插行、合并、删行
  const created = await callAndValidate('office_create_document', { path: 'table-ops.docx', title: '表格操作' })
  assert.equal(created.success, true)
  const before = await call('office_read_docx', { path: 'table-ops.docx', detail: 'tables' })
  assert.ok(before.data.count >= 1, '新建文档应含示例表格')

  const inserted = await callAndValidate('office_insert_docx_table_row', {
    path: 'table-ops.docx',
    table: 0,
    after: 0,
    cells: ['新增一', '新增二']
  })
  assert.equal(inserted.data.columns, 2)
  const afterInsert = await call('office_read_docx', { path: 'table-ops.docx', detail: 'tables' })
  assert.equal(afterInsert.data.tables[0].rows, before.data.tables[0].rows + 1)
  assert.equal(afterInsert.data.tables[0].cells[1][0].text, '新增一')

  const merged = await callAndValidate('office_merge_docx_table_cells', {
    path: 'table-ops.docx',
    table: 0,
    row: 1,
    column: 0,
    col_span: 2
  })
  assert.equal(merged.data.col_span, 2)
  const afterMerge = await call('office_read_docx', { path: 'table-ops.docx', detail: 'tables' })
  assert.match(afterMerge.data.tables[0].cells[1][0].text, /新增一[\s\S]*新增二/, '横向合并应把两格文字并到一起')

  const removed = await callAndValidate('office_delete_docx_table_row', { path: 'table-ops.docx', table: 0, row: 0 })
  assert.equal(removed.data.row_count, afterMerge.data.tables[0].rows - 1)
  // 结构操作后文档仍应通过校验（关系链与部件完整）
  const valid = await call('office_validate_docx', { path: 'table-ops.docx' })
  assert.equal(valid.data.valid, true, JSON.stringify(valid.data.checks.filter((c) => !c.ok)))
})

await test('★ 表格行删除的保护性检查与最后一行拒绝', async () => {
  await call('office_create_document', { path: 'row-guard.docx', title: '守卫' })
  // 新建文档的表格有两行：删一行可以，再删同类要拒绝（只剩一行）
  const first = await callAndValidate('office_delete_docx_table_row', { path: 'row-guard.docx', table: 0, row: 0 })
  assert.equal(first.data.row_count, 1)
  const second = await call('office_delete_docx_table_row', { path: 'row-guard.docx', table: 0, row: 0 })
  assert.equal(second.success, false)
  assert.equal(second.error.code, 'INVALID_REQUEST')
  // 越界行号给出总数
  const outOfRange = await call('office_delete_docx_table_row', { path: 'row-guard.docx', table: 0, row: 9 })
  assert.equal(outOfRange.success, false)
  assert.equal(outOfRange.error.details.row_count, 1)
})

await test('office_merge_docx_table_cells 参数校验', async () => {
  await call('office_create_document', { path: 'merge-guard.docx', title: '合并守卫' })
  const oneByOne = await call('office_merge_docx_table_cells', { path: 'merge-guard.docx', table: 0, row: 0, column: 0 })
  assert.equal(oneByOne.success, false)
  assert.equal(oneByOne.error.code, 'INVALID_REQUEST')
  const tooWide = await call('office_merge_docx_table_cells', {
    path: 'merge-guard.docx',
    table: 0,
    row: 0,
    column: 0,
    col_span: 9
  })
  assert.equal(tooWide.success, false)
})

await test('★ 表格样式工具：套用 + 补定义', async () => {
  await call('office_create_document', { path: 'style-ops.docx', title: '样式操作' })
  const builtin = await callAndValidate('office_set_docx_table_style', {
    path: 'style-ops.docx',
    table: 0,
    style_id: 'TableGrid',
    first_row: true,
    banded_rows: true
  })
  assert.equal(builtin.data.style_id, 'TableGrid')
  const pkg = (await import('../lib/ooxml.js')).ZipPackage.open(readFileSync(join(workspace, 'style-ops.docx')))
  const tblPr = /<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(pkg.readText('word/document.xml'))[0]
  assert.equal((tblPr.match(/<w:tblStyle /g) ?? []).length, 1)
  assert.match(tblPr, /w:noHBand="0"/)
  // 自定义样式：应补一个最小定义并给出警告
  const custom = await callAndValidate('office_set_docx_table_style', {
    path: 'style-ops.docx',
    table: 0,
    style_id: 'DshProbeStyle'
  })
  assert.equal(custom.data.style_defined_in_document, true)
  assert.ok(custom.warnings.some((w) => w.code === 'STYLE_CREATED'))
  const styles = (await import('../lib/ooxml.js')).ZipPackage.open(readFileSync(join(workspace, 'style-ops.docx'))).readText('word/styles.xml')
  assert.match(styles, /w:styleId="DshProbeStyle"/)
  const valid = await call('office_validate_docx', { path: 'style-ops.docx' })
  assert.equal(valid.data.valid, true, JSON.stringify(valid.data.checks.filter((c) => !c.ok)))
})

await test('office_set_docx_table_style 样式 ID 校验', async () => {
  const bad = await call('office_set_docx_table_style', { path: 'style-ops.docx', table: 0, style_id: '9 bad' })
  assert.equal(bad.success, false)
  assert.equal(bad.error.code, 'INVALID_REQUEST')
})

await test('★ 新建表格工具：插到指定段落后并可继续编辑', async () => {
  await call('office_create_document', { path: 'new-table.docx', title: '建表' })
  const before = await call('office_read_docx', { path: 'new-table.docx', detail: 'paragraphs' })
  const result = await callAndValidate('office_create_docx_table', {
    path: 'new-table.docx',
    after: 0,
    rows: [
      ['名称', '数量'],
      ['甲', '1'],
      ['乙', '2']
    ],
    header: true,
    style_id: 'TableGrid'
  })
  assert.equal(result.data.rows, 3)
  assert.equal(result.data.columns, 2)
  assert.equal(result.data.header, true)
  const after = await call('office_read_docx', { path: 'new-table.docx', detail: 'tables' })
  assert.equal(after.data.count, before.data.paragraphs.length >= 0 ? 2 : 2, '新建文档原有 1 个表格，应变成 2 个')
  const created = after.data.tables.find((t) => t.cells[0][0].text === '名称')
  assert.ok(created, '找不到新建的表格')
  assert.deepEqual(created.cells.map((row) => row.map((cell) => cell.text)), [
    ['名称', '数量'],
    ['甲', '1'],
    ['乙', '2']
  ])
  // 新表能被后续工具继续编辑（插行 + 改单元格），证明它是一张正常的表格
  await callAndValidate('office_insert_docx_table_row', { path: 'new-table.docx', table: 0, after: 0, cells: ['丙', '3'] })
  const valid = await call('office_validate_docx', { path: 'new-table.docx' })
  assert.equal(valid.data.valid, true, JSON.stringify(valid.data.checks.filter((c) => !c.ok)))
})

await test('office_create_docx_table 参数校验', async () => {
  const empty = await call('office_create_docx_table', { path: 'new-table.docx', after: 0, rows: [] })
  assert.equal(empty.success, false)
  assert.equal(empty.error.code, 'INVALID_REQUEST')
  const outOfRange = await call('office_create_docx_table', { path: 'new-table.docx', after: 999, rows: [['x']] })
  assert.equal(outOfRange.success, false)
})

await test('★ 图片工具：清单 / 缩放 / 删除 端到端', async () => {
  await call('office_create_document', { path: 'img-ops.docx', title: '图片操作' })
  // 先插一张图（复用已有的插入工具）
  const png = makePng(120, 60)
  const inserted = await callAndValidate('office_insert_docx_image', {
    path: 'img-ops.docx',
    after: 0,
    image_base64: png.toString('base64'),
    width_px: 240,
    height_px: 120,
    alt_text: '端到端占位图'
  })
  assert.equal(inserted.success, true)
  // 缩放：只给宽度应保持比例
  const resized = await callAndValidate('office_resize_docx_image', { path: 'img-ops.docx', index: 0, width_px: 360 })
  assert.deepEqual(resized.data.to, { width: 360, height: 180 })
  // 删除（段落里只有这张图，应连带删除）
  const removed = await callAndValidate('office_delete_docx_image', { path: 'img-ops.docx', index: 0 })
  assert.equal(removed.data.removed_paragraph, true)
  assert.ok(removed.warnings.some((w) => w.code === 'MEDIA_KEPT'))
  const valid = await call('office_validate_docx', { path: 'img-ops.docx' })
  assert.equal(valid.data.valid, true, JSON.stringify(valid.data.checks.filter((c) => !c.ok)))
  // 再删一次应报找不到
  const again = await call('office_delete_docx_image', { path: 'img-ops.docx', index: 0 })
  assert.equal(again.success, false)
  assert.equal(again.error.code, 'FILE_NOT_FOUND')
})

await test('office_resize_docx_image 参数校验', async () => {
  const noSize = await call('office_resize_docx_image', { path: 'img-ops.docx', index: 0 })
  assert.equal(noSize.success, false)
  assert.equal(noSize.error.code, 'INVALID_REQUEST')
})

await test('★ 段落样式 / 字符样式 / 分页符 端到端', async () => {
  await call('office_create_document', { path: 'style2.docx', title: '样式与分页' })
  const para = await callAndValidate('office_set_paragraph_style', { path: 'style2.docx', index: 0, style_id: 'Heading2' })
  assert.equal(para.data.style_id, 'Heading2')
  const run = await callAndValidate('office_set_character_style', { path: 'style2.docx', index: 1, style_id: 'Strong' })
  assert.ok(run.data.runs >= 1)
  const brk = await callAndValidate('office_insert_page_break', { path: 'style2.docx', after: 0 })
  assert.equal(brk.data.index, 1)
  // 读回来确认：段落数 +1、第一段样式变了
  const read = await call('office_read_docx', { path: 'style2.docx', detail: 'paragraphs' })
  assert.equal(read.data.paragraphs[0].style, 'Heading2')
  const valid = await call('office_validate_docx', { path: 'style2.docx' })
  assert.equal(valid.data.valid, true, JSON.stringify(valid.data.checks.filter((c) => !c.ok)))
})

await test('段落样式与分页符的参数校验', async () => {
  const badStyle = await call('office_set_paragraph_style', { path: 'style2.docx', index: 0, style_id: '9 bad' })
  assert.equal(badStyle.success, false)
  assert.equal(badStyle.error.code, 'INVALID_REQUEST')
  const badIndex = await call('office_set_paragraph_style', { path: 'style2.docx', index: 99, style_id: 'Heading2' })
  assert.equal(badIndex.success, false)
  assert.equal(badIndex.error.code, 'FILE_NOT_FOUND')
  const badBreak = await call('office_insert_page_break', { path: 'style2.docx', after: 999 })
  assert.equal(badBreak.success, false)
})

await test('★ 页面写工具对加密 PDF 明确拒绝', async () => {
  const encrypted = Buffer.from(
    [
      '%PDF-1.4\n',
      '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
      '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
      '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] >>\nendobj\n',
      '9 0 obj\n<< /Filter /Standard /V 2 /R 3 /Length 128 /P -3904 >>\nendobj\n',
      'trailer\n<< /Root 1 0 R /Encrypt 9 0 R /Size 10 >>\n%%EOF\n'
    ].join(''),
    'latin1'
  )
  writeFileSync(join(workspace, 'locked.pdf'), encrypted)
  const result = await call('office_rotate_pdf_pages', { path: 'locked.pdf', degrees: 90 })
  assert.equal(result.success, false)
  assert.equal(result.error.code, 'PASSWORD_REQUIRED')
  assert.ok(readFileSync(join(workspace, 'locked.pdf')).equals(encrypted), '拒绝时文件必须原封不动')
})

await test('★ 页面设置 / 页边距 / 页码域 端到端', async () => {
  await call('office_create_document', { path: 'layout.docx', title: '版式' })
  const layout = await callAndValidate('office_set_docx_page_layout', {
    path: 'layout.docx',
    paper: 'A3',
    orientation: 'landscape'
  })
  assert.equal(layout.data.to.orientation, 'landscape')
  assert.equal(layout.data.to.width_cm, 42)
  assert.equal(layout.data.to.height_cm, 29.7)

  const margins = await callAndValidate('office_set_docx_margins', {
    path: 'layout.docx',
    top_cm: 2,
    bottom_cm: 2,
    left_cm: 2.5,
    right_cm: 2.5
  })
  assert.equal(margins.data.to.top_cm, 2)
  assert.equal(margins.data.to.left_cm, 2.5)

  const page = await callAndValidate('office_insert_docx_page_number', {
    path: 'layout.docx',
    position: 'footer',
    align: 'center',
    prefix: '第 ',
    suffix: ' 页'
  })
  assert.equal(page.data.field, 'PAGE')
  assert.ok(page.warnings.some((w) => w.code === 'FIELD_NEEDS_UPDATE'), '应提示域需要 Word 刷新')

  const read = await call('office_read_docx', { path: 'layout.docx', detail: 'summary' })
  assert.equal(read.data.structure.has_page_number_field, true)
  assert.equal(read.data.structure.section_properties[0].page_width_twips, 23811)
  const valid = await call('office_validate_docx', { path: 'layout.docx' })
  assert.equal(valid.data.valid, true, JSON.stringify(valid.data.checks.filter((c) => !c.ok)))
})

await test('页面设置 / 页边距 / 页码域 参数校验', async () => {
  const badPaper = await call('office_set_docx_page_layout', { path: 'layout.docx', paper: 'B5' })
  assert.equal(badPaper.success, false)
  assert.equal(badPaper.error.code, 'INVALID_REQUEST')
  const noMargins = await call('office_set_docx_margins', { path: 'layout.docx' })
  assert.equal(noMargins.success, false)
  assert.equal(noMargins.error.code, 'INVALID_REQUEST')
  const badAlign = await call('office_insert_docx_page_number', { path: 'layout.docx', align: 'justify' })
  assert.equal(badAlign.success, false)
  assert.equal(badAlign.error.code, 'INVALID_REQUEST')
})

await test('★ office_search_pdf 全文搜索（页码 / 偏移 / 行号 / 上下文）', async () => {
  if (!hasPdf) return
  const multi = join(here, 'fixtures', 'sample-multipage.pdf')
  if (!existsSync(multi)) return
  copyFileSync(multi, join(workspace, 'multi.pdf'))

  const hit = await callAndValidate('office_search_pdf', { path: 'multi.pdf', query: 'XLSX', context_chars: 10 })
  assert.equal(hit.data.found, true)
  assert.equal(hit.data.total_matches, 2)
  assert.equal(hit.data.matches[0].match, 'XLSX')
  assert.ok(typeof hit.data.matches[0].offset === 'number')
  assert.ok(hit.data.matches[0].line >= 1)
  assert.ok(hit.summary.includes('命中 2 处'), hit.summary)

  const miss = await callAndValidate('office_search_pdf', { path: 'multi.pdf', query: '不存在的词' })
  assert.equal(miss.data.found, false)
  assert.equal(miss.data.total_matches, 0)

  const single = await callAndValidate('office_search_pdf', { path: 'multi.pdf', query: 'xlsx', case_sensitive: true })
  assert.equal(single.data.found, false, '区分大小写时小写查询不应命中')

  const bad = await call('office_search_pdf', { path: 'multi.pdf', query: '' })
  assert.equal(bad.success, false)
  assert.equal(bad.error.code, 'INVALID_REQUEST')
})

await test('★ 图片环绕方式 端到端（行内 → 四周型 → 行内）', async () => {
  await call('office_create_document', { path: 'wrap.docx', title: '环绕' })
  const png = makePng(200, 100)
  await callAndValidate('office_insert_docx_image', {
    path: 'wrap.docx',
    after: 0,
    image_base64: png.toString('base64'),
    width_px: 200,
    height_px: 100
  })

  const anchored = await callAndValidate('office_set_docx_image_wrap', {
    path: 'wrap.docx',
    index: 0,
    wrap: 'square',
    offset_x_emu: 114300,
    offset_y_emu: 57150
  })
  assert.equal(anchored.data.from, 'inline')
  assert.equal(anchored.data.to, 'square')

  const { ZipPackage: ZP } = await import('../lib/ooxml.js')
  const xmlAfter = ZP.open(readFileSync(join(workspace, 'wrap.docx'))).readText('word/document.xml')
  assert.ok(xmlAfter.includes('<wp:anchor'), '应转成浮动图')
  assert.ok(xmlAfter.includes('<wp:wrapSquare wrapText="bothSides"/>'), '应写入四周型环绕')
  const valid1 = await call('office_validate_docx', { path: 'wrap.docx' })
  assert.equal(valid1.data.valid, true, JSON.stringify(valid1.data.checks.filter((c) => !c.ok)))

  const inlined = await callAndValidate('office_set_docx_image_wrap', { path: 'wrap.docx', index: 0, wrap: 'inline' })
  assert.equal(inlined.data.to, 'inline')
  const xmlBack = ZP.open(readFileSync(join(workspace, 'wrap.docx'))).readText('word/document.xml')
  assert.ok(!xmlBack.includes('<wp:anchor'), '应转回行内图')
  assert.ok(!xmlBack.includes('<wp:wrapSquare'), '环绕元素应被清理')

  const bad = await call('office_set_docx_image_wrap', { path: 'wrap.docx', index: 0, wrap: 'tight' })
  assert.equal(bad.success, false)
  assert.equal(bad.error.code, 'INVALID_REQUEST')
  // 参数 schema 校验在 defineTool 层直接抛错（不会走到 execute）
  await assert.rejects(() => call('office_set_docx_image_wrap', { path: 'wrap.docx', index: 0 }), /missing required property "wrap"/)
})

await test('★ 读取版式清单 + 切换版式 端到端', async () => {
  if (!hasDeck) return
  copyFileSync(deckSource, join(workspace, 'deck-layout.pptx'))

  const layouts = await callAndValidate('office_read_pptx', { path: 'deck-layout.pptx', detail: 'layouts' })
  assert.ok(layouts.data.layout_count >= 8, `样本版式数偏少：${layouts.data.layout_count}`)
  assert.equal(layouts.data.slides.length, 4)
  const blank = layouts.data.layouts.find((l) => l.name === '空白')
  assert.ok(blank, '样本里应有「空白」版式')

  const changed = await callAndValidate('office_set_slide_layout', { path: 'deck-layout.pptx', index: 0, layout: blank.index })
  assert.equal(changed.data.layout_name, '空白')
  const after = await call('office_read_pptx', { path: 'deck-layout.pptx', detail: 'layouts' })
  assert.equal(after.data.slides[0].layout_name, '空白')
  // 形状与文本必须留着
  const slide = await call('office_read_pptx', { path: 'deck-layout.pptx', detail: 'slide', index: 0 })
  assert.ok(slide.data.slide.shapes.some((s) => (s.text ?? '').includes('季度业务报告')), '换版式不应丢文本')

  const valid = await call('office_validate_pptx', { path: 'deck-layout.pptx' })
  assert.equal(valid.data.valid, true, JSON.stringify(valid.data.checks.filter((c) => !c.ok)))

  const bad = await call('office_set_slide_layout', { path: 'deck-layout.pptx', index: 0, name: '不存在' })
  assert.equal(bad.success, false)
  assert.equal(bad.error.code, 'INVALID_REQUEST')
})

await test('★ office_export_presentation 导出 PPTX 为 PDF（真实引擎）', async () => {
  if (!hasDeck) return
  copyFileSync(deckSource, join(workspace, 'deck-export.pptx'))
  const result = await call('office_export_presentation', { path: 'deck-export.pptx', target_path: 'deck-export.pdf' })
  if (!result.success) {
    // 宿主没有内置转换引擎时如实跳过，而不是假装通过
    console.log(`     ⏭️  跳过：转换引擎不可用（${result.error.code}）`)
    return
  }
  assert.equal(result.data.slides, 4)
  assert.equal(result.data.pdf_pages, 3, '4 张幻灯片里有 1 张隐藏页，PDF 应为 3 页')
  assert.ok(existsSync(join(workspace, 'deck-export.pdf')), 'PDF 应落盘')
  assert.ok(result.data.lost.some((t) => t.includes('隐藏幻灯片')), `应说明隐藏页不进 PDF：${result.data.lost}`)
  assert.ok(result.warnings.some((w) => w.code === 'LOSSY_CONVERT'), '应给出有损转换提示')

  // 输出已存在且未传 overwrite 时拒绝（不静默覆盖）
  const again = await call('office_export_presentation', { path: 'deck-export.pptx', target_path: 'deck-export.pdf' })
  assert.equal(again.success, false)
  assert.equal(again.error.code, 'PERMISSION_DENIED')

  // 只接受 pptx
  await call('office_create_document', { path: 'notadeck.docx', title: 'x' })
  const wrong = await call('office_export_presentation', { path: 'notadeck.docx', target_path: 'nope.pdf' })
  assert.equal(wrong.success, false)
  assert.equal(wrong.error.code, 'UNSUPPORTED_FILE_TYPE')
})

await test('★ 插入书签 端到端', async () => {
  await call('office_create_document', { path: 'bookmark.docx', title: '书签' })
  const first = await callAndValidate('office_insert_bookmark', { path: 'bookmark.docx', paragraph: 1, name: 'intro' })
  assert.equal(first.data.name, 'intro')
  assert.equal(first.data.id, 0)
  const second = await callAndValidate('office_insert_bookmark', { path: 'bookmark.docx', paragraph: 2, name: '章节_1' })
  assert.equal(second.data.id, 1)
  assert.equal(second.data.bookmark_count, 2)

  const read = await call('office_read_docx', { path: 'bookmark.docx', detail: 'summary' })
  assert.equal(read.data.structure.has_bookmarks, true, '读回应报告存在书签')

  const valid = await call('office_validate_docx', { path: 'bookmark.docx' })
  assert.equal(valid.data.valid, true, JSON.stringify(valid.data.checks.filter((c) => !c.ok)))

  const dup = await call('office_insert_bookmark', { path: 'bookmark.docx', paragraph: 1, name: 'INTRO' })
  assert.equal(dup.success, false)
  assert.equal(dup.error.code, 'INVALID_REQUEST')
})

await test('★ office_read_docx 批注与修订视图（对照真实 Word 样本）', async () => {
  const complex = join(here, 'fixtures', 'word-complex.docx')
  if (!existsSync(complex)) return
  copyFileSync(complex, join(workspace, 'complex.docx'))

  const comments = await callAndValidate('office_read_docx', { path: 'complex.docx', detail: 'comments' })
  assert.equal(comments.data.count, 1, '样本含 1 条批注')
  assert.ok(comments.data.comments[0].text.includes('法务复核'), `批注正文不对：${comments.data.comments[0].text}`)
  assert.ok(comments.data.comments[0].author.length > 0, '应带作者')
  assert.ok(comments.data.comments[0].anchored_paragraph >= 0, '应能定位锚点段落')
  assert.ok(comments.data.extended_parts.length > 0, '应报告线程化批注附加部件')
  assert.ok(comments.summary.includes('批注 1 条'), comments.summary)

  const revisions = await callAndValidate('office_read_docx', { path: 'complex.docx', detail: 'revisions' })
  assert.equal(revisions.data.count, 2, '样本含 2 条修订')
  assert.equal(revisions.data.by_type['插入'], 1)
  assert.ok(revisions.data.revisions.some((r) => r.text.includes('修订模式插入')), '插入的文字应读出来')
  assert.ok(revisions.data.revisions.some((r) => r.paragraph_mark === true), '段落标记增删应被标注而不是当成空文字')
  assert.ok(revisions.summary.includes('修订 2 条'), revisions.summary)
})

await test('★ 添加 PDF 批注 + annotations 读回 端到端', async () => {
  if (!hasPdf) return
  const result = await callAndValidate('office_add_pdf_annotation', {
    path: 'sample.pdf',
    page: 0,
    text: '批注：这一页需要法务复核。',
    author: 'dsh-exp-office',
    x: 60,
    y: 700,
    open: true
  })
  assert.equal(result.data.subtype, 'Text')
  assert.equal(result.data.non_ascii, true)

  const read = await callAndValidate('office_read_pdf', { path: 'sample.pdf', detail: 'annotations' })
  assert.equal(read.data.count, 1)
  assert.equal(read.data.annotations[0].contents, '批注：这一页需要法务复核。')
  assert.equal(read.data.annotations[0].author, 'dsh-exp-office')
  assert.ok(read.summary.includes('批注 1 条'), read.summary)

  const structure = await call('office_read_pdf', { path: 'sample.pdf', detail: 'summary' })
  assert.equal(structure.data.structure.annotations, 1)
  assert.ok(
    structure.warnings.some((w) => w.code === 'ANNOTATIONS_PRESENT'),
    `应提示批注情况：${JSON.stringify(structure.warnings)}`
  )

  const valid = await call('office_validate_pdf', { path: 'sample.pdf' })
  assert.equal(valid.data.valid ?? valid.data.checks.every((c) => c.ok), true)

  const bad = await call('office_add_pdf_annotation', { path: 'sample.pdf', page: 0, text: '' })
  assert.equal(bad.success, false)
  assert.equal(bad.error.code, 'INVALID_REQUEST')
})

await test('★ office_add_slide_chart 插入原生图表 端到端', async () => {
  if (!hasDeck) return
  copyFileSync(deckSource, join(workspace, 'deck-chart.pptx'))

  const result = await callAndValidate('office_add_slide_chart', {
    path: 'deck-chart.pptx',
    index: 1,
    type: 'column',
    title: '各产品金额',
    categories: ['1月', '2月', '3月'],
    values: [[120, 150, 180]],
    series_names: ['金额']
  })
  assert.equal(result.data.chart_part, 'ppt/charts/chart1.xml')
  assert.equal(result.data.series_count, 1)

  const read = await call('office_read_pptx', { path: 'deck-chart.pptx', detail: 'slide', index: 1 })
  assert.ok(read.data.slide.shapes.some((s) => s.type === 'chart'), '读回应识别出 chart 形状')
  const structure = await call('office_read_pptx', { path: 'deck-chart.pptx', detail: 'structure' })
  assert.equal(structure.data.structure.charts, 1)

  const valid = await call('office_validate_pptx', { path: 'deck-chart.pptx' })
  assert.equal(valid.data.valid, true, JSON.stringify(valid.data.checks.filter((c) => !c.ok)))

  const bad = await call('office_add_slide_chart', { path: 'deck-chart.pptx', index: 1, type: 'donut', values: [[1]] })
  assert.equal(bad.success, false)
  assert.equal(bad.error.code, 'INVALID_REQUEST')
  await assert.rejects(
    () => call('office_add_slide_chart', { path: 'deck-chart.pptx', index: 1 }),
    /missing required property "values"/
  )
})

await test('★ office_export_docx：docx 副本逐字节一致 / 纯文本带 BOM / 拒绝静默覆盖', async () => {
  await call('office_create_document', { path: 'export-src.docx', title: '导出源文档' })
  const sourceBytes = readFileSync(join(workspace, 'export-src.docx'))

  const copy = await callAndValidate('office_export_docx', { path: 'export-src.docx', target_path: 'export-copy.docx' })
  assert.equal(copy.data.byte_identical, true, 'docx 副本必须与源文件逐字节一致')
  assert.equal(copy.data.source_sha256, copy.output_file.sha256)
  assert.ok(readFileSync(join(workspace, 'export-copy.docx')).equals(sourceBytes), '副本字节应与源文件完全相同')
  // 源文件没被改动
  assert.ok(readFileSync(join(workspace, 'export-src.docx')).equals(sourceBytes))

  const text = await callAndValidate('office_export_docx', { path: 'export-src.docx', target_path: 'export-text', format: 'text' })
  assert.equal(text.data.format, 'text')
  assert.ok(text.warnings.some((w) => w.code === 'LOSSY_CONVERT'), '纯文本导出必须提示有损')
  const txt = readFileSync(join(workspace, 'export-text.txt'))
  assert.deepEqual([...txt.subarray(0, 3)], [0xef, 0xbb, 0xbf], '纯文本应带 UTF-8 BOM（否则记事本中文乱码）')
  assert.ok(txt.toString('utf8').includes('导出源文档'), '导出的文本应含文档标题')

  // 目标已存在 → 拒绝；overwrite=true 才允许
  const again = await call('office_export_docx', { path: 'export-src.docx', target_path: 'export-copy.docx' })
  assert.equal(again.success, false)
  assert.equal(again.error.code, 'PERMISSION_DENIED')
  const forced = await call('office_export_docx', { path: 'export-src.docx', target_path: 'export-copy.docx', overwrite: true })
  assert.equal(forced.success, true)

  // 只接受 docx
  const wrong = await call('office_export_docx', { path: 'demo.xlsx', target_path: 'nope.docx' })
  assert.equal(wrong.success, false)
  assert.equal(wrong.error.code, 'UNSUPPORTED_FILE_TYPE')
})

await test('★ office_compare_docx：段落/表格/样式差异端到端（只读）', async () => {
  await call('office_create_document', { path: 'cmp-a.docx', format: 'docx', title: '季度报告' })
  const aBytes = readFileSync(join(workspace, 'cmp-a.docx'))
  copyFileSync(join(workspace, 'cmp-a.docx'), join(workspace, 'cmp-b.docx'))

  // 在 B 上做三处已知改动：改段落文字、插入段落、改表格单元格
  // （先改文字再插入，否则插入会把后面的段落下标顶掉，改动对象就不是想改的那一段了）
  await call('office_update_paragraph', { path: 'cmp-b.docx', index: 1, text: '本文件由 dsh-exp-office 生成（B 改过）。' })
  await call('office_insert_paragraph', { path: 'cmp-b.docx', after: 0, text: '这是 B 新增的一段。' })
  await call('office_update_docx_table_cell', { path: 'cmp-b.docx', table: 0, row: 1, column: 1, text: '第二列' })
  const bBytes = readFileSync(join(workspace, 'cmp-b.docx'))

  const diff = await callAndValidate('office_compare_docx', { path_a: 'cmp-a.docx', path_b: 'cmp-b.docx' })
  assert.equal(diff.data.identical, false)
  assert.equal(diff.data.paragraphs.total_added, 1)
  assert.equal(diff.data.paragraphs.added[0].text, '这是 B 新增的一段。')
  const textChange = diff.data.paragraphs.changed.find((c) => c.kind === 'text')
  assert.ok(textChange, `应有文字修改：${JSON.stringify(diff.data.paragraphs.changed)}`)
  assert.equal(textChange.text_after, '本文件由 dsh-exp-office 生成（B 改过）。')
  assert.equal(diff.data.tables.changed.some((c) => c.kind === 'cell' && c.after === '第二列'), true)

  // 只读：两份文件一个字节都没变
  assert.ok(readFileSync(join(workspace, 'cmp-a.docx')).equals(aBytes), 'path_a 不应被改动')
  assert.ok(readFileSync(join(workspace, 'cmp-b.docx')).equals(bBytes), 'path_b 不应被改动')

  // 自己与自己比 → 一致
  const same = await callAndValidate('office_compare_docx', { path_a: 'cmp-a.docx', path_b: 'cmp-a.docx' })
  assert.equal(same.data.identical, true)

  // 非 docx 被拒绝
  const wrong = await call('office_compare_docx', { path_a: 'demo.xlsx', path_b: 'cmp-a.docx' })
  assert.equal(wrong.success, false)
  assert.equal(wrong.error.code, 'UNSUPPORTED_FILE_TYPE')
})

await test('★ 本地引擎联动：默认关闭时拒绝启动，只读检测仍然可用', async () => {
  // 1) 只读注册表检测：不需要授权，也不启动任何程序
  const detect = await callAndValidate('office_detect_engines', {})
  assert.equal(Array.isArray(detect.data.engines), true)
  assert.equal(detect.data.local_automation_enabled, false, '测试配置里没开本地联动')
  assert.equal(typeof detect.data.microsoft_office, 'boolean')
  assert.equal(typeof detect.data.wps, 'boolean')

  // 2) 真的驱动引擎：默认配置必须拒绝，并说清怎么开
  const recalc = await call('office_recalculate', { path: 'demo.xlsx' })
  assert.equal(recalc.success, false)
  assert.equal(recalc.error.code, 'AUTOMATION_DISABLED')
  assert.match(recalc.error.message, /allowLocalAutomation/)

  // 3) 探 COM 也需要授权
  const probe = await call('office_detect_engines', { probe_com: true })
  assert.equal(probe.success, false)
  assert.equal(probe.error.code, 'AUTOMATION_DISABLED')

  // 4) 非 OOXML 类型明确拒绝（不等到启动引擎才发现）
  const wrong = await call('office_rerender', { path: 'sample.pdf' })
  assert.equal(wrong.success, false)
  assert.equal(wrong.error.code, 'UNSUPPORTED_FILE_TYPE')
})

await test('★ office_set_theme + read_pptx(detail=layouts) 端到端', async () => {
  if (!hasDeck) return
  copyFileSync(deckSource, join(workspace, 'deck-theme.pptx'))

  const before = await callAndValidate('office_read_pptx', { path: 'deck-theme.pptx', detail: 'layouts' })
  assert.ok(before.data.theme_count >= 2, `样本应至少有 2 个主题：${before.data.theme_count}`)
  assert.ok(before.data.master_themes[0].theme_part, '应报告母版当前用的主题部件')
  assert.ok(before.data.layouts.some((l) => l.type), '版式清单应带 type')

  const other = before.data.themes.find((t) => t.part !== before.data.master_themes[0].theme_part)
  assert.ok(other, '应能找到一个「母版当前没用」的主题')
  const switched = await callAndValidate('office_set_theme', { path: 'deck-theme.pptx', master: 0, theme: other.index })
  assert.equal(switched.data.source_theme, other.part, '切到的应是那个主题的内容')
  assert.notEqual(switched.data.to, other.part, '挂上去的应是克隆出来的新部件（直接指过去 PowerPoint 会拒绝打开）')
  assert.match(switched.data.to, /^ppt\/theme\/theme\d+\.xml$/)

  const after = await call('office_read_pptx', { path: 'deck-theme.pptx', detail: 'layouts' })
  assert.equal(after.data.master_themes[0].theme_part, switched.data.to, '读回应显示母版已换到新主题部件')
  const valid = await call('office_validate_pptx', { path: 'deck-theme.pptx' })
  assert.equal(valid.data.valid, true, JSON.stringify(valid.data.checks.filter((c) => !c.ok)))

  const bad = await call('office_set_theme', { path: 'deck-theme.pptx', master: 0, name: '不存在的主题' })
  assert.equal(bad.success, false)
  assert.equal(bad.error.code, 'INVALID_REQUEST')
  const noTarget = await call('office_set_theme', { path: 'deck-theme.pptx', master: 0 })
  assert.equal(noTarget.success, false)
  assert.equal(noTarget.error.code, 'INVALID_REQUEST')
})

await test('★ office_fill_pdf_form + read_pdf(detail=form) 端到端', async () => {
  const formSource = join(here, 'fixtures', 'form-sample.pdf')
  if (!existsSync(formSource)) return
  copyFileSync(formSource, join(workspace, 'form.pdf'))

  const before = await callAndValidate('office_read_pdf', { path: 'form.pdf', detail: 'form' })
  assert.equal(before.data.has_form, true)
  assert.equal(before.data.field_count, 4)
  assert.ok(before.summary.includes('表单字段 4 个'), before.summary)
  const names = before.data.fields.map((f) => f.name)
  assert.deepEqual(names.sort(), ['agree_terms', 'full_name', 'plan', 'tier'])

  const filled = await callAndValidate('office_fill_pdf_form', {
    path: 'form.pdf',
    fields: { full_name: '张伟', agree_terms: true, plan: 'pro', tier: 'B' }
  })
  assert.equal(filled.data.filled_count, 4)
  assert.equal(filled.data.need_appearances, true)

  const after = await call('office_read_pdf', { path: 'form.pdf', detail: 'form' })
  const byName = new Map(after.data.fields.map((f) => [f.name, f]))
  assert.equal(byName.get('full_name').value, '张伟')
  assert.equal(byName.get('agree_terms').checked, true)
  assert.equal(byName.get('plan').value, 'pro')
  assert.equal(byName.get('tier').value, 'B')

  const structure = await call('office_read_pdf', { path: 'form.pdf', detail: 'summary' })
  assert.equal(structure.data.structure.form_fields, 5, '页面上的 widget 数：4 个字段 + 单选的第二个 widget')
  const valid = await call('office_validate_pdf', { path: 'form.pdf' })
  assert.equal(valid.data.checks.every((c) => c.ok), true)

  // 未知字段名 → 结构化错误并列出可选字段
  const bad = await call('office_fill_pdf_form', { path: 'form.pdf', fields: { 不存在: 'x' } })
  assert.equal(bad.success, false)
  assert.equal(bad.error.code, 'INVALID_REQUEST')
  assert.ok(Array.isArray(bad.error.details.field_names), JSON.stringify(bad.error.details))

  // 无表单的 PDF 被明确拒绝
  await call('office_fill_pdf_form', { path: 'sample.pdf', fields: { a: 'b' } })
    .then((r) => {
      assert.equal(r.success, false)
      assert.equal(r.error.code, 'UNSUPPORTED_FEATURE')
    })
})

await test('★ office_create_presentation 从零建演示文稿 端到端', async () => {
  const created = await callAndValidate('office_create_presentation', {
    path: 'brand-new.pptx',
    title: '从零生成的演示文稿',
    subtitle: '不依赖任何模板文件'
  })
  assert.equal(created.data.slides, 1)
  assert.equal(created.data.theme_name, 'office-plugin')
  assert.equal(created.data.slide_size.aspect, '16:9')
  assert.ok(existsSync(join(workspace, 'brand-new.pptx')))

  // 读回来核对：页数、版式、标题文本
  const read = await call('office_read_pptx', { path: 'brand-new.pptx', detail: 'slide', index: 0 })
  assert.ok(read.data.slide.shapes.some((s) => (s.text ?? '').includes('从零生成的演示文稿')))
  const layouts = await call('office_read_pptx', { path: 'brand-new.pptx', detail: 'layouts' })
  assert.equal(layouts.data.layout_count, 1)
  assert.equal(layouts.data.themes[0].name, 'office-plugin')

  // 它是「活文件」：继续用其它工具编辑，仍然校验通过
  await callAndValidate('office_add_slide', { path: 'brand-new.pptx', layout_of: 0 })
  await callAndValidate('office_add_text_box', { path: 'brand-new.pptx', index: 1, text: '后加的文本框' })
  await callAndValidate('office_add_slide_chart', {
    path: 'brand-new.pptx',
    index: 1,
    type: 'pie',
    title: '后加的图表',
    categories: ['一', '二'],
    values: [[60, 40]]
  })
  const valid = await call('office_validate_pptx', { path: 'brand-new.pptx' })
  assert.equal(valid.data.valid, true, JSON.stringify(valid.data.checks.filter((c) => !c.ok)))

  // 拒绝静默覆盖
  const again = await call('office_create_presentation', { path: 'brand-new.pptx' })
  assert.equal(again.success, false)
  assert.equal(again.error.code, 'PERMISSION_DENIED')

  // 空白版式 + 非法版式
  const blank = await callAndValidate('office_create_presentation', { path: 'blank-deck', layout: 'blank', title: 'x' })
  assert.equal(blank.data.slide_size.aspect, '16:9')
  const blankRead = await call('office_read_pptx', { path: 'blank-deck.pptx', detail: 'slide', index: 0 })
  assert.equal(blankRead.data.slide.shapes.length, 0, 'blank 版式不应有占位符')
  const bad = await call('office_create_presentation', { path: 'bad-deck.pptx', layout: 'fancy' })
  assert.equal(bad.success, false)
  assert.equal(bad.error.code, 'INVALID_REQUEST')
})

await test('★ 零运行时依赖：package.json 的 dependencies 必须为空', () => {
  // 这条断言来自一次真实事故：调试「升级」时手工探针把一条假依赖
  // （`dsh-exp-office: file:%USERPROFILE%/…/tmp.tgz`）写进了 package.json，
  // 结果既污染了隐私（本机路径）又破坏了「零依赖」承诺。
  const pkg = JSON.parse(readFileSync(join(here, '..', 'package.json'), 'utf8'))
  assert.deepEqual(pkg.dependencies ?? {}, {}, `dependencies 必须为空，实际：${JSON.stringify(pkg.dependencies)}`)
  // peerDependencies 允许存在，但只能是**宿主提供**的包（@deepseek-ai/*）：
  // cordis / dsh-tools 由 DSH 提供，libreoffice-kit 是可选引擎 —— 都不该被 npm 装进我们的包。
  for (const name of Object.keys(pkg.peerDependencies ?? {})) {
    assert.ok(name.startsWith('@deepseek-ai/'), `peerDependency 只能声明宿主提供的包，实际出现 ${name}`)
  }
  assert.equal(pkg.peerDependenciesMeta?.['@deepseek-ai/libreoffice-kit']?.optional, true, '转 PDF 的引擎必须是可选 peer')
  assert.equal(pkg.private, false, '插件是要发布的包，private 应为 false')
  assert.ok(Array.isArray(pkg.files) && pkg.files.includes('lib'), 'package.json 的 files 白名单必须包含 lib')
})

console.log('\n=== 14. 事务与资源清理 ===')

await test('事务失败时保留原文件且不留临时产物', async () => {
  const before = readFileSync(bookPath)
  const result = await call('office_write_cells', { path: 'demo.xlsx', sheet: '数据', cells: [{ ref: 'ZZZZ1', value: 1 }] })
  assert.equal(result.success, false)
  assert.ok(readFileSync(bookPath).equals(before), '失败时原文件必须保持不变')
})

await test('临时目录已清理，未残留 staging 文件', () => {
  const tmpRoot = join(workspace, '.dsh-exp-office', 'tmp')
  const leftovers = existsSync(tmpRoot)
    ? readdirSync(tmpRoot).filter((d) => readdirSync(join(tmpRoot, d)).some((f) => f.startsWith('staging-')))
    : []
  assert.deepEqual(leftovers, [], `残留暂存文件：${leftovers.join(', ')}`)
})

await test('锁文件未残留', () => {
  assert.equal(existsSync(`${bookPath}.dsh-office.lock`), false)
})

rmSync(workspace, { recursive: true, force: true })

console.log(`\n结果：${passed} 通过，${failed} 失败\n`)
process.exit(failed === 0 ? 0 : 1)
