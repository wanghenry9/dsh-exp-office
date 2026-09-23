/**
 * 工具定义层：把统一文档能力暴露为 DSH 工具。
 *
 * 每个工具都遵循开发要求 §4 的统一响应格式，并在文件修改路径上强制走
 * 「备份 → 暂存 → 校验 → 原子替换 → 失败回滚」的事务流程。
 *
 * `defineTool` 由调用方注入，因此本模块既能被 DSH 加载，也能在独立测试中
 * 用真实或桩实现的 defineTool 驱动。
 *
 * @module dsh-exp-office/tools
 */

import { readdir, stat, writeFile, mkdir, rm, readFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { basename, dirname, join, relative } from 'node:path'
import { MAX_EDITABLE_CELLS, Workbook } from './xlsx.js'
import { DocxDocument, buildBasicDocx, detectImageType, readImageSize } from './docx.js'
import { PptxPresentation } from './pptx.js'
import { buildBlankPptx } from './pptx-template.js'
import { PdfDocument } from './pdf.js'
import { PDF_PAGE_SIZES, buildPdf } from './pdf-writer.js'
import { convertToPdf, CONVERTIBLE_EXTENSIONS } from './convert.js'
import { AUTOMATION_KINDS, DEFAULT_AUTOMATION_TIMEOUT_MS, OfficeAutomation } from './automation.js'
import { ZipPackage, XmlDoc, find, findAll, attr } from './ooxml.js'
import {
  ManagedStore,
  Transaction,
  FileLock,
  detectFileType,
  readDocument,
  resolveSafePath,
  sha256,
  MAX_FILE_BYTES
} from './workspace.js'
import { OfficeError, ok, fail, newRequestId, projectLevel } from './errors.js'

/** 可空字段的 schema：docs §4 的响应格式在成功/失败时保留 null 占位。 */
const nullable = (spec) => ({ oneOf: [spec, { type: 'null' }] })

/** 所有工具共用的输出 schema：统一响应信封（docs §4.2 / §4.3）。 */
const ENVELOPE_SCHEMA = {
  type: 'object',
  additionalProperties: true,
  properties: {
    success: { type: 'boolean', required: true },
    request_id: { type: 'string', required: true },
    summary: { type: 'string', required: true },
    document_id: nullable({ type: 'string' }),
    output_file: nullable({ type: 'object', additionalProperties: true }),
    warnings: { type: 'array', items: { type: 'json' } },
    changes: { type: 'array', items: { type: 'json' } },
    performance: nullable({ type: 'object', additionalProperties: true }),
    error: nullable({ type: 'object', additionalProperties: true }),
    data: nullable({ type: 'object', additionalProperties: true })
  }
}

/**
 * 构造统一的输出声明。
 * @param {string} cardTitle - 卡片标题。
 * @returns {object} `output` 字段。
 */
function envelopeOutput(cardTitle) {
  return {
    schema: ENVELOPE_SCHEMA,
    render: (_args, value) => [{ type: 'text', text: value.summary ?? cardTitle }]
  }
}

/**
 * 把异常收敛为失败响应，保证工具永不抛裸异常。
 * @param {string} requestId - 关联标识。
 * @param {unknown} err - 异常。
 * @returns {object} 失败信封。
 */
function failure(requestId, err) {
  const envelope = fail(requestId, err)
  const office = envelope.error
  return {
    ...envelope,
    summary: `❌ [${office.code}] ${office.message}｜可重试：${office.retryable ? '是' : '否'}｜建议：${office.solution}`
  }
}

/**
 * 创建全部 Office 工具定义。
 *
 * @param {object} deps - 依赖注入。
 * @param {Function} deps.defineTool - DSH 的工具定义函数。
 * @param {object} [deps.config] - 插件配置。
 * @returns {object[]} 工具定义数组。
 */
export function createOfficeTools({ defineTool, config = {} }) {
  const workspaceRoot = config.workspaceRoot ?? process.cwd()
  const allowOverwriteOriginals = config.allowOverwriteOriginals === true
  const maxBytes = config.maxFileBytes ?? MAX_FILE_BYTES
  const maxEditableCells = config.maxEditableCells ?? MAX_EDITABLE_CELLS
  // 本地 Office/WPS 联动**默认关闭**：只有显式配置才允许启动本机 Office 进程。
  const automation = new OfficeAutomation({
    enabled: config.allowLocalAutomation === true,
    powershell: config.powershellPath ?? undefined,
    scriptPath: config.automationScriptPath ?? undefined,
    scratchDir: join(workspaceRoot, '.dsh-exp-office', 'tmp'),
    auditPath: join(workspaceRoot, '.dsh-exp-office', 'audit', 'office-automation.jsonl'),
    timeoutMs: config.automationTimeoutMs ?? DEFAULT_AUTOMATION_TIMEOUT_MS
  })

  /** 惰性初始化的受管目录。 */
  let storePromise = null
  /**
   * 取得受管目录。
   * @returns {Promise<ManagedStore>} 受管目录。
   */
  const store = () => {
    storePromise ??= ManagedStore.open(workspaceRoot)
    return storePromise
  }

  /** 任务登记表：支撑 preview/execute/status/cancel 五件套。 */
  const tasks = new Map()

  /**
   * 解析输入文件路径。
   * @param {string} path - 用户路径。
   * @param {boolean} [mustExist] - 是否要求存在。
   * @returns {string} 绝对路径。
   */
  const resolvePath = (path, mustExist = true) => resolveSafePath(path, { workspaceRoot, mustExist })

  /**
   * 打开一个工作簿并附带字节与类型信息。
   * @param {string} path - 用户路径。
   * @returns {Promise<object>} `{wb, buffer, type, absolute}`。
   */
  async function openWorkbook(path) {
    const absolute = resolvePath(path, true)
    const { buffer, type } = await readDocument(absolute, { maxBytes })
    if (type.kind !== 'zip' || !['xlsx', 'xlsm', 'xltx'].includes(type.ext)) {
      throw new OfficeError('UNSUPPORTED_FILE_TYPE', `文件 ${basename(absolute)} 不是 Excel 工作簿（识别为 ${type.ext}）。`, { detected: type.ext })
    }
    const wb = Workbook.open(buffer)
    // 写入路径要把整张表解析成 DOM，规模上限在这里统一注入（见 xlsx.js 的 #assertEditable）。
    wb.maxEditableCells = maxEditableCells
    return { wb, buffer, type, absolute }
  }

  /**
   * 汇总工作簿的安全标志，供写入前提示（docs §五.4）。
   * @param {Workbook} wb - 工作簿。
   * @returns {object[]} 警告列表。
   */
  function securityWarnings(wb) {
    const warnings = []
    if (wb.info.hasMacro) {
      warnings.push({ code: 'MACRO_DETECTED', message: '文件含 VBA 宏。插件不会执行宏，宏内容按原字节保留。' })
    }
    if (wb.info.hasExternalLinks) {
      warnings.push({ code: 'EXTERNAL_LINKS_PRESENT', message: '文件含外部链接，插件不会更新外部链接，也不会访问外部数据源。' })
    }
    if (wb.info.hasSignatures) {
      warnings.push({ code: 'DIGITAL_SIGNATURE_PRESENT', message: '文件含数字签名，任何修改都会使签名失效。' })
    }
    if (wb.info.hasPivot) {
      warnings.push({ code: 'PIVOT_PRESENT', message: '文件含数据透视表，插件不修改透视表定义。' })
    }
    if (wb.info.hasCharts) {
      warnings.push({ code: 'CHARTS_PRESENT', message: '文件含图表，图表部件按原字节保留。' })
    }
    return warnings
  }

  /**
   * 事务式保存工作簿：暂存 → 重新打开校验 → 原子替换。
   *
   * @param {object} args - 参数。
   * @returns {Promise<object>} `{outputFile, warnings}`。
   */
  async function saveWorkbook({ wb, sourcePath, changes, warnings, requestId }) {
    const temp = await (await store()).temp('xlsx')
    const transaction = await Transaction.begin(sourcePath, temp.path)
    for (const w of warnings) transaction.warn(w)
    try {
      const bytes = wb.save()
      const result = await transaction.commit(bytes, {
        overwrite: true,
        // docs §十一：保存后必须重新打开验证，验证失败不能返回成功
        verify: (out) => {
          const reopened = Workbook.open(out)
          const before = wb.sheetNames()
          const after = reopened.sheetNames()
          if (before.length !== after.length || before.some((n, i) => n !== after[i])) {
            throw new OfficeError('OUTPUT_VALIDATION_FAILED', `重新打开后工作表清单不一致：${before.join(',')} → ${after.join(',')}`)
          }
          for (const c of changes) {
            if (c.type !== 'cell_update' && c.type !== 'cell_clear') continue
            const value = reopened.readRange(c.sheet, c.range).rows[0]?.[0]
            if (c.type === 'cell_clear' && value !== null) {
              throw new OfficeError('OUTPUT_VALIDATION_FAILED', `清空的单元格 ${c.sheet}!${c.range} 重新读取后仍有内容。`)
            }
          }
        }
      })
      return {
        outputFile: {
          name: basename(sourcePath),
          path: sourcePath,
          size: result.size,
          mime_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          sha256: result.sha256
        },
        warnings: result.warnings
      }
    } finally {
      await transaction.rollback()
      await temp.dispose()
    }
  }

  /**
   * 带锁与事务地执行一次工作簿修改。
   * @param {object} args - 参数。
   * @returns {Promise<object>} 响应信封。
   */
  async function mutateWorkbook({ path, requestId, mutate, summaryOf, options = {} }) {
    const absolute = resolvePath(path, true)
    const lock = await FileLock.acquire(absolute)
    const started = Date.now()
    try {
      const { wb, type } = await openWorkbook(path)
      const warnings = securityWarnings(wb)
      const changes = []
      const extra = await mutate({ wb, changes, warnings, type })
      const saved = await saveWorkbook({ wb, sourcePath: absolute, changes, warnings, requestId })
      const envelope = ok({
        requestId,
        documentId: sha256(absolute),
        outputFile: saved.outputFile,
        warnings: saved.warnings,
        changes,
        performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
        data: { ...extra }
      })
      return { ...envelope, summary: summaryOf(saved, changes, extra) }
    } finally {
      await lock.release()
    }
  }

  const tools = []

  // ───────────────────────── 文件信息 ─────────────────────────

  tools.push(
    defineTool({
      name: 'office_get_file_info',
      description:
        '检视一个办公文件的真实类型、大小、哈希与结构。不依赖扩展名，而是读取文件真实字节判断类型；对 Excel/Word/PowerPoint 还会报告宏、外部链接、数字签名、图表、透视表等工作簿级安全标志。修改文件前先用它做安全检查。',
      parameters: {
        path: { type: 'string', required: true, description: '文件路径（相对会话工作区或绝对路径）。' }
      },
      output: envelopeOutput('文件信息'),
      async execute(args, exec) {
        const requestId = newRequestId()
        try {
          const absolute = resolvePath(args.path, true)
          const { buffer, type } = await readDocument(absolute, { maxBytes })
          const info = await stat(absolute)
          const data = {
            path: absolute,
            name: basename(absolute),
            size: info.size,
            detected_type: type.ext,
            mime_type: type.mime,
            sha256: sha256(buffer),
            extension_mismatch: type.mismatch,
            is_macro_enabled: type.macro
          }
          let structure = null
          if (['xlsx', 'xlsm', 'xltx'].includes(type.ext)) {
            const wb = Workbook.open(buffer)
            structure = {
              sheets: wb.info.sheets.map((s) => ({ name: s.name, state: s.state, sheet_id: s.sheetId })),
              date_system: wb.info.date1904 ? '1904' : '1900',
              has_macro: wb.info.hasMacro,
              has_external_links: wb.info.hasExternalLinks,
              has_charts: wb.info.hasCharts,
              has_pivot_tables: wb.info.hasPivot,
              has_digital_signature: wb.info.hasSignatures
            }
          } else if (type.kind === 'zip') {
            const pkg = ZipPackage.open(buffer)
            structure = { part_count: pkg.names().length, has_macro: pkg.names().some((n) => n.includes('vbaProject')) }
          } else if (type.kind === 'pdf') {
            structure = { note: 'PDF 结构解析在后续阶段提供。' }
          } else {
            structure = { note: `类型 ${type.ext} 暂无结构解析。` }
          }
          return {
            ...ok({ requestId, documentId: data.sha256, performance: null, data: { ...data, structure } }),
            summary: `📄 ${data.name}｜类型 ${type.ext}（${type.mime}）｜${info.size} 字节｜sha256 ${data.sha256.slice(0, 16)}…${type.mismatch ? '｜⚠️ 扩展名与实际类型不符' : ''}`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '检视办公文件', kind: 'read', rawInput: args.path })
    })
  )

  tools.push(
    defineTool({
      name: 'office_list_files',
      description: '列出会话工作区（或指定子目录）中的办公文件，附带真实类型与大小。用于在批量处理前盘点可用文件。',
      parameters: {
        directory: { type: 'string', description: '相对工作区的子目录，默认工作区根目录。' },
        extensions: { type: 'array', items: { type: 'string' }, description: '仅列出这些扩展名，默认 xlsx/xlsm/docx/pptx/pdf。' },
        recursive: { type: 'boolean', description: '是否递归子目录，默认 false。' }
      },
      output: envelopeOutput('文件列表'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const wanted = new Set((args.extensions ?? ['xlsx', 'xlsm', 'docx', 'pptx', 'pdf']).map((e) => e.toLowerCase().replace(/^\./, '')))
          const base = args.directory ? resolvePath(args.directory, true) : resolvePath('.', true)
          const found = []
          const walk = async (dir, depth) => {
            if (found.length > 2000) return
            const entries = await readdir(dir, { withFileTypes: true })
            for (const entry of entries) {
              if (entry.name.startsWith('.')) continue
              const full = join(dir, entry.name)
              if (entry.isDirectory()) {
                if (args.recursive === true && depth < 6) await walk(full, depth + 1)
                continue
              }
              const ext = entry.name.slice(entry.name.lastIndexOf('.') + 1).toLowerCase()
              if (!wanted.has(ext)) continue
              const info = await stat(full)
              found.push({ path: relative(workspaceRoot, full) || entry.name, size: info.size, modified: info.mtime.toISOString() })
            }
          }
          await walk(base, 0)
          return {
            ...ok({ requestId, performance: null, data: { count: found.length, files: found } }),
            summary: `📁 ${base} 下找到 ${found.length} 个办公文件${args.recursive === true ? '（含子目录）' : ''}。`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '列出办公文件', kind: 'read', rawInput: args })
    })
  )

  // ───────────────────────── XLSX 读取 ─────────────────────────

  tools.push(
    defineTool({
      name: 'office_read_workbook',
      description:
        '读取工作簿结构：工作表清单、可见性、使用范围与安全标志（宏/外部链接/图表/透视表/数字签名）。这是读取 Excel 的入口，先看结构再用 office_read_sheet 或 office_read_range 取具体数据。',
      parameters: {
        path: { type: 'string', required: true, description: '工作簿路径。' },
        detail: { type: 'string', enum: ['summary', 'detail'], description: '响应级别，summary 只给结构概览（默认），detail 附带每张表的行列规模。' }
      },
      output: envelopeOutput('工作簿结构'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const { wb, buffer, type } = await openWorkbook(args.path)
          const sheets = wb.info.sheets.map((s) => {
            const entry = { name: s.name, state: s.state, sheet_id: s.sheetId }
            if (args.detail === 'detail') {
              const content = wb.readSheet(s.name)
              entry.row_count = content.row_count
              entry.column_count = content.column_count
              entry.cell_count = content.cells.length
            }
            return entry
          })
          const warnings = securityWarnings(wb)
          const envelope = ok({
            requestId,
            documentId: sha256(buffer),
            warnings,
            performance: null,
            data: {
              path: args.path,
              name: basename(args.path),
              detected_type: type.ext,
              date_system: wb.info.date1904 ? '1904' : '1900',
              sheet_count: sheets.length,
              sheets,
              security: {
                has_macro: wb.info.hasMacro,
                has_external_links: wb.info.hasExternalLinks,
                has_charts: wb.info.hasCharts,
                has_pivot_tables: wb.info.hasPivot,
                has_digital_signature: wb.info.hasSignatures
              }
            }
          })
          return { ...envelope, summary: `📊 ${basename(args.path)}｜${sheets.length} 张工作表：${sheets.map((s) => s.name).join('、')}${warnings.length ? `｜${warnings.length} 条安全提示` : ''}` }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '读取工作簿结构', kind: 'read', rawInput: args.path })
    })
  )

  tools.push(
    defineTool({
      name: 'office_read_sheet',
      description: '读取一张工作表的单元格数据（稀疏列表，含公式与类型）。大表请改用 office_read_range 只取所需区域。',
      parameters: {
        path: { type: 'string', required: true, description: '工作簿路径。' },
        sheet: { type: 'string', required: true, description: '工作表名。' },
        max_cells: { type: 'integer', description: '单元格数量上限，默认 50000，防止超大表撑爆上下文。' },
        offset: { type: 'integer', description: '从第几个单元格开始返回，用于分页。' },
        limit: { type: 'integer', description: '本次最多返回多少个单元格，默认 2000。' }
      },
      output: envelopeOutput('工作表数据'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const { wb, buffer } = await openWorkbook(args.path)
          const content = wb.readSheet(args.sheet, { maxCells: args.max_cells ?? 50000 })
          const offset = args.offset ?? 0
          const limit = args.limit ?? 2000
          const page = content.cells.slice(offset, offset + limit)
          return {
            ...ok({
              requestId,
              documentId: sha256(buffer),
              performance: null,
              data: {
                sheet: args.sheet,
                row_count: content.row_count,
                column_count: content.column_count,
                total_cells: content.cells.length,
                offset,
                returned: page.length,
                truncated: offset + page.length < content.cells.length,
                cells: page
              }
            }),
            summary: `📄 ${args.sheet}｜共 ${content.row_count} 行 × ${content.column_count} 列、${content.cells.length} 个单元格｜本次返回 ${page.length} 个${offset + page.length < content.cells.length ? '（还有更多，用 offset 继续）' : ''}`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '读取工作表', kind: 'read', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_read_range',
      description: '按 A1 区域引用读取工作表的一个矩形区域，返回二维数组。这是最省资源、最推荐的 Excel 读取方式。',
      parameters: {
        path: { type: 'string', required: true, description: '工作簿路径。' },
        sheet: { type: 'string', required: true, description: '工作表名。' },
        range: { type: 'string', required: true, description: '区域引用，如 A1:D20。' }
      },
      output: envelopeOutput('区域数据'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const { wb, buffer } = await openWorkbook(args.path)
          const result = wb.readRange(args.sheet, args.range)
          return {
            ...ok({ requestId, documentId: sha256(buffer), performance: null, data: result }),
            summary: `📄 ${args.sheet}!${args.range}｜${result.rows.length} 行 × ${result.rows[0]?.length ?? 0} 列`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '读取区域', kind: 'read', rawInput: args })
    })
  )

  // ───────────────────────── XLSX 写入 ─────────────────────────

  tools.push(
    defineTool({
      name: 'office_write_cells',
      description:
        '批量写入单元格（文本/数字/布尔/日期/公式），这是修改 Excel 的首选工具。默认采用最小修改模式：只重写目标单元格，图表、图片、宏、条件格式、数据验证等无关部件保持逐字节不变。写公式时请同时给出 value 作为缓存结果；插件不会自动重算公式。',
      parameters: {
        path: { type: 'string', required: true, description: '工作簿路径。' },
        sheet: { type: 'string', required: true, description: '工作表名。' },
        cells: {
          type: 'array',
          required: true,
          description: '要写入的单元格列表。',
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              ref: { type: 'string', required: true, description: 'A1 引用，如 B2。' },
              value: { type: 'json', required: true, description: '写入的值；null 表示清空该单元格。' },
              formula: { type: 'string', description: '公式正文（不含前导 =），如 SUM(A1:A9)。' },
              type: { type: 'string', enum: ['auto', 'number', 'string', 'boolean', 'date', 'datetime'], description: '值类型，默认 auto。' },
              style: { type: 'integer', description: '要套用的 cellXfs 样式索引。' }
            }
          }
        },
        dry_run: { type: 'boolean', description: '仅预览，不写入任何文件。' }
      },
      output: envelopeOutput('写入单元格'),
      async execute(args) {
        const requestId = newRequestId()
        if (args.dry_run === true) {
          return {
            ...ok({
              requestId,
              performance: null,
              data: {
                dry_run: true,
                sheet: args.sheet,
                planned: args.cells.map((c) => ({ ref: c.ref, value: c.value ?? null, formula: c.formula ?? null })),
                note: '预览模式，未写入任何文件。'
              }
            }),
            summary: `🔍 预览：将在「${args.sheet}」写入 ${args.cells.length} 个单元格，未改动文件。`
          }
        }
        try {
          const envelope = await mutateWorkbook({
            path: args.path,
            requestId,
            mutate: async ({ wb, changes }) => {
              const applied = wb.writeCells(args.sheet, args.cells.map((c) => ({ ...c, date1904: wb.info.date1904 })))
              changes.push(...applied)
              return { written: applied.length }
            },
            summaryOf: (_saved, changes, extra) => `✍️ 已写入「${args.sheet}」${extra.written} 个单元格：${changes.slice(0, 8).map((c) => c.range).join('、')}${changes.length > 8 ? ` 等 ${changes.length} 处` : ''}`
          })
          return envelope
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `写入 ${args.cells?.length ?? 0} 个单元格`, kind: 'write', rawInput: args })
    })
  )

  /**
   * 声明一个行列移位工具。
   * @param {string} toolName - 工具名。
   * @param {'row'|'column'} axis - 轴。
   * @param {'insert'|'delete'} action - 动作。
   * @param {string} label - 中文标签。
   * @returns {object} 工具定义。
   */
  const shiftTool = (toolName, axis, action, label) =>
    defineTool({
      name: toolName,
      description: `${label}。会重排受影响的行号与单元格引用；公式文本、合并区域与条件格式中的引用不会被自动改写，此时会返回警告，请在 Excel/WPS 中复核。`,
      parameters: {
        path: { type: 'string', required: true, description: '工作簿路径。' },
        sheet: { type: 'string', required: true, description: '工作表名。' },
        start: { type: 'integer', required: true, description: `起始${axis === 'row' ? '行号' : '列号'}（1 基）。` },
        count: { type: 'integer', description: '数量，默认 1。' }
      },
      output: envelopeOutput(label),
      async execute(args) {
        const requestId = newRequestId()
        try {
          let warnings = []
          const envelope = await mutateWorkbook({
            path: args.path,
            requestId,
            mutate: async ({ wb, changes }) => {
              const result = wb.shiftRowsOrColumns(args.sheet, { axis, action, start: args.start, count: args.count ?? 1 })
              changes.push(...result.changes)
              warnings = result.warnings
              return { shifted: result.changes[0].shifted, removed: result.changes[0].removed }
            },
            summaryOf: () => `${label}完成（起始 ${args.start}，数量 ${args.count ?? 1}）。`
          })
          envelope.warnings = [...envelope.warnings, ...warnings]
          return envelope
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: label, kind: 'write', rawInput: args })
    })

  tools.push(shiftTool('office_insert_rows', 'row', 'insert', '插入行'))
  tools.push(shiftTool('office_delete_rows', 'row', 'delete', '删除行'))
  tools.push(shiftTool('office_insert_columns', 'column', 'insert', '插入列'))
  tools.push(shiftTool('office_delete_columns', 'column', 'delete', '删除列'))

  tools.push(
    defineTool({
      name: 'office_find_and_replace',
      description: '在工作簿中查找并替换文本。默认跳过公式单元格以避免破坏公式语义；默认大小写不敏感。返回每一处替换的精确位置。大范围替换属于高风险操作，建议先用 dry_run 预览。',
      parameters: {
        path: { type: 'string', required: true, description: '工作簿路径。' },
        find: { type: 'string', required: true, description: '要查找的文本。' },
        replace: { type: 'string', required: true, description: '替换为的文本。' },
        sheet: { type: 'string', description: '限定单张工作表，默认全部工作表。' },
        match_case: { type: 'boolean', description: '是否区分大小写，默认 false。' },
        whole_cell: { type: 'boolean', description: '是否要求整格完全匹配，默认 false。' },
        include_formulas: { type: 'boolean', description: '是否也替换公式文本，默认 false。' },
        dry_run: { type: 'boolean', description: '仅预览命中数量，不写入。' }
      },
      output: envelopeOutput('查找替换'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          if (args.dry_run === true) {
            const { wb } = await openWorkbook(args.path)
            const result = wb.findAndReplace({
              find: args.find,
              replace: args.replace,
              sheet: args.sheet ?? null,
              matchCase: args.match_case === true,
              wholeCell: args.whole_cell === true,
              includeFormulas: args.include_formulas === true
            })
            return {
              ...ok({ requestId, performance: null, data: { dry_run: true, replacements: result.replacements, changes: result.changes.slice(0, 100) } }),
              summary: `🔍 预览：命中 ${result.replacements} 处，未改动文件。`
            }
          }
          const envelope = await mutateWorkbook({
            path: args.path,
            requestId,
            mutate: async ({ wb, changes }) => {
              const result = wb.findAndReplace({
                find: args.find,
                replace: args.replace,
                sheet: args.sheet ?? null,
                matchCase: args.match_case === true,
                wholeCell: args.whole_cell === true,
                includeFormulas: args.include_formulas === true
              })
              changes.push(...result.changes)
              return { replacements: result.replacements }
            },
            summaryOf: (_s, _c, extra) => `🔁 已替换 ${extra.replacements} 处「${args.find}」→「${args.replace}」。`
          })
          return envelope
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `查找替换：${args.find}`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_merge_cells',
      description: '合并一个单元格区域（写入 mergeCells 节点）。只影响合并记录，不改动其它部件。',
      parameters: {
        path: { type: 'string', required: true, description: '工作簿路径。' },
        sheet: { type: 'string', required: true, description: '工作表名。' },
        range: { type: 'string', required: true, description: '要合并的区域，如 A1:C1。' }
      },
      output: envelopeOutput('合并单元格'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateWorkbook({
            path: args.path,
            requestId,
            mutate: async ({ wb, changes }) => {
              const result = wb.mergeCells(args.sheet, args.range)
              changes.push({ type: 'merge_cells', sheet: args.sheet, range: args.range })
              return result
            },
            summaryOf: () => `🔗 已合并「${args.sheet}」的 ${args.range}。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `合并 ${args.range}`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_create_table',
      description:
        '把一片区域变成 Excel 表格对象（ListObject）：带表头、自动筛选与套用样式。会一次写齐四处（表格部件、工作表里的 tableParts、工作表关系、内容类型 Override），少任何一处 Excel 都会判定文件需要修复。列名取表头行的值；表头为空或重复时会把生成的列名（列1、列2…）写回表头单元格，保证列名非空且唯一（Excel 自己建表也是这个行为）。已在同一工作表上与该区域重叠的表格会被拒绝。',
      parameters: {
        path: { type: 'string', required: true, description: '工作簿路径。' },
        sheet: { type: 'string', required: true, description: '工作表名。' },
        ref: { type: 'string', required: true, description: '表格区域（含表头行），如 A1:D5。' },
        name: { type: 'string', description: '表格名（工作簿内唯一），默认「表1」「表2」…' },
        style: { type: 'string', description: '表格样式名，默认 TableStyleMedium2（如 TableStyleLight9、TableStyleMedium9）。' },
        show_totals: { type: 'boolean', description: '是否含汇总行；为 true 时 ref 必须已包含汇总行（最后一行）。' },
        has_header: { type: 'boolean', description: '是否含表头行，默认 true。' }
      },
      output: envelopeOutput('创建表格'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateWorkbook({
            path: args.path,
            requestId,
            mutate: async ({ wb, changes, warnings }) => {
              const result = wb.createTable(args.sheet, {
                ref: args.ref,
                name: args.name,
                style: args.style,
                showTotals: args.show_totals === true,
                hasHeader: args.has_header !== false
              })
              changes.push({ type: 'create_table', sheet: args.sheet, ...result })
              if (result.header_fixes.length > 0) {
                warnings.push({
                  code: 'HEADER_CELLS_FILLED',
                  message: `表头有 ${result.header_fixes.length} 处为空或重复，已写入生成的列名：${result.header_fixes
                    .map((f) => `${f.ref}=${f.name}`)
                    .join('、')}`
                })
              }
              return result
            },
            summaryOf: (_r, _c, extra) =>
              `🧮 已在「${args.sheet}」的 ${extra.ref} 创建表格「${extra.display_name}」（${extra.columns.length} 列，样式 ${extra.style}）。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `创建表格 ${args.ref}`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_create_chart',
      description:
        '在工作表上创建图表（柱状/条形/折线/饼图）。会一次写齐五处：图表部件、绘图部件（锚点引用图表）、工作表的 <drawing>、工作表与绘图两层关系、以及图表与绘图两个内容类型 Override —— 少任何一处 Excel 都会判定文件需要修复。分类与数值会同时写入**缓存值**（c:cat / c:val），因此打开就能看到图形，不必先重算。工作表已有绘图（图片或别的图表）时复用同一个绘图部件并追加锚点（一张工作表只能有一个 <drawing> 引用）。',
      parameters: {
        path: { type: 'string', required: true, description: '工作簿路径。' },
        sheet: { type: 'string', required: true, description: '放图表的工作表名。' },
        type: { type: 'string', required: true, enum: ['column', 'bar', 'line', 'pie'], description: '图表类型：柱状/条形/折线/饼图。' },
        series: {
          type: 'array',
          required: true,
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              values: { type: 'string', required: true, description: '数值区域，如 D2:D4（可写「表名!D2:D4」跨表引用）。' },
              name: { type: 'string', description: '系列名（字面文本）。' },
              name_ref: { type: 'string', description: '系列名取自某个单元格，如 D1（优先于 name）。' }
            }
          },
          description: '数据系列（饼图只支持一个）。'
        },
        categories: { type: 'string', description: '分类标签区域，如 A2:A4；饼图也用它做标签。' },
        title: { type: 'string', description: '图表标题。' },
        anchor: { type: 'string', description: '左上角锚点单元格，默认 F2。' },
        width_px: { type: 'integer', description: '图表宽度（像素），默认 480。' },
        height_px: { type: 'integer', description: '图表高度（像素），默认 300。' },
        data_sheet: { type: 'string', description: '数据所在工作表；默认与图表同一张。' }
      },
      output: envelopeOutput('创建图表'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateWorkbook({
            path: args.path,
            requestId,
            mutate: async ({ wb, changes }) => {
              const result = wb.createChart(args.sheet, {
                type: args.type,
                series: (args.series ?? []).map((item) => ({ values: item.values, name: item.name, nameRef: item.name_ref })),
                categories: args.categories,
                title: args.title,
                anchor: args.anchor,
                widthPx: args.width_px,
                heightPx: args.height_px,
                dataSheet: args.data_sheet
              })
              changes.push({ type: 'create_chart', ...result })
              return result
            },
            summaryOf: (_r, _c, extra) =>
              `📈 已在「${extra.sheet}」创建${extra.type}图表（${extra.series_count} 个系列，锚点 ${extra.anchor}），部件 ${extra.chart_part}。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `创建图表（${args.type}）`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_add_worksheet',
      description: '新增一张工作表，可选择插入位置。新增后即可用 office_write_cells 写入内容。',
      parameters: {
        path: { type: 'string', required: true, description: '工作簿路径。' },
        name: { type: 'string', required: true, description: '新工作表名（≤31 字符，不含 \\ / ? * [ ] :）。' },
        index: { type: 'integer', description: '插入位置（0 基），默认追加到末尾。' }
      },
      output: envelopeOutput('新增工作表'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateWorkbook({
            path: args.path,
            requestId,
            mutate: async ({ wb, changes }) => {
              const result = wb.addWorksheet(args.name, { index: args.index ?? null })
              changes.push({ type: 'add_worksheet', sheet: args.name, index: result.index })
              return result
            },
            summaryOf: (_s, _c, extra) => `➕ 已新增工作表「${args.name}」（位置 ${extra.index}）。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `新增工作表 ${args.name}`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_rename_worksheet',
      description: '重命名工作表。只改工作表名，不动其内容与其它工作表。',
      parameters: {
        path: { type: 'string', required: true, description: '工作簿路径。' },
        from: { type: 'string', required: true, description: '原工作表名。' },
        to: { type: 'string', required: true, description: '新工作表名。' }
      },
      output: envelopeOutput('重命名工作表'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateWorkbook({
            path: args.path,
            requestId,
            mutate: async ({ wb, changes }) => {
              const result = wb.renameWorksheet(args.from, args.to)
              changes.push({ type: 'rename_worksheet', sheet: args.to, from: args.from })
              return result
            },
            summaryOf: () => `✏️ 已将「${args.from}」重命名为「${args.to}」。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `重命名 ${args.from} → ${args.to}`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_delete_worksheet',
      description:
        '删除一张工作表。这是不可逆的破坏性操作：必须显式传 confirm=true 才会执行。工作簿至少保留一张可见工作表，删除最后一张会被拒绝。',
      parameters: {
        path: { type: 'string', required: true, description: '工作簿路径。' },
        sheet: { type: 'string', required: true, description: '要删除的工作表名。' },
        confirm: { type: 'boolean', required: true, description: '必须为 true 才执行删除（人工确认位）。' }
      },
      output: envelopeOutput('删除工作表'),
      async execute(args) {
        const requestId = newRequestId()
        if (args.confirm !== true) {
          return failure(requestId, new OfficeError('PERMISSION_DENIED', '删除工作表是破坏性操作，需要显式传 confirm=true。', { needsConfirmation: true }))
        }
        try {
          return await mutateWorkbook({
            path: args.path,
            requestId,
            mutate: async ({ wb, changes }) => {
              const result = wb.deleteWorksheet(args.sheet)
              changes.push({ type: 'delete_worksheet', sheet: args.sheet })
              return result
            },
            summaryOf: () => `🗑️ 已删除工作表「${args.sheet}」。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `删除工作表 ${args.sheet}`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_set_cell_style',
      description:
        '设置单元格样式：字体（加粗/斜体/下划线/颜色/字号/字体名）、填充底色、数字格式与水平对齐。实现上只在 styles.xml 追加新的样式定义并把目标单元格的 s 指过去，既有样式定义一律不动，因此不会影响其它使用相同样式的单元格。目标单元格必须已存在——请先用 office_write_cells 写入内容。',
      parameters: {
        path: { type: 'string', required: true, description: '工作簿路径。' },
        sheet: { type: 'string', required: true, description: '工作表名。' },
        cells: {
          type: 'array',
          required: true,
          description: '要设置样式的单元格列表。',
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              ref: { type: 'string', required: true, description: 'A1 引用。' },
              bold: { type: 'boolean', description: '加粗。' },
              italic: { type: 'boolean', description: '斜体。' },
              underline: { type: 'boolean', description: '下划线。' },
              font_color: { type: 'string', description: '字体颜色，#RRGGBB。' },
              font_size: { type: 'number', description: '字号（磅）。' },
              font_name: { type: 'string', description: '字体名，如 宋体 / Arial。' },
              fill_color: { type: 'string', description: '填充底色，#RRGGBB。' },
              number_format: { type: 'string', description: '数字格式代码，如 #,##0.00、0.00%、yyyy-mm-dd。' },
              horizontal: { type: 'string', enum: ['left', 'center', 'right', 'fill', 'justify'], description: '水平对齐。' }
            }
          }
        }
      },
      output: envelopeOutput('设置单元格样式'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateWorkbook({
            path: args.path,
            requestId,
            mutate: async ({ wb, changes }) => {
              const edits = args.cells.map((c) => ({
                ref: c.ref,
                bold: c.bold,
                italic: c.italic,
                underline: c.underline,
                fontColor: c.font_color,
                fontSize: c.font_size,
                fontName: c.font_name,
                fillColor: c.fill_color,
                numberFormat: c.number_format,
                horizontal: c.horizontal
              }))
              const applied = wb.applyCellStyle(args.sheet, edits)
              changes.push(...applied)
              return { styled: applied.length, details: applied.map((c) => `${c.range}: ${c.applied}`) }
            },
            summaryOf: (_s, _c, extra) => `🎨 已设置「${args.sheet}」${extra.styled} 个单元格的样式：${extra.details.slice(0, 5).join('；')}${extra.details.length > 5 ? ` 等 ${extra.details.length} 处` : ''}`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `设置 ${args.cells?.length ?? 0} 个单元格样式`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_export_workbook',
      description:
        '把工作簿导出为新文件，不修改源文件。format=xlsx 原样复制（用于留存当前状态或另存），format=csv 把指定工作表导出为带 BOM 的 UTF-8 CSV（Excel/WPS 打开中文不乱码）。注意 CSV 是纯文本，公式、样式、图表与多工作表信息都会丢失。',
      parameters: {
        path: { type: 'string', required: true, description: '源工作簿路径。' },
        target_path: { type: 'string', required: true, description: '导出目标路径。未写扩展名时按 format 自动补全。' },
        format: { type: 'string', enum: ['xlsx', 'csv'], description: '导出格式，默认 xlsx。' },
        sheet: { type: 'string', description: '导出 CSV 时指定工作表，默认第一张。' },
        overwrite: { type: 'boolean', description: '目标已存在时是否覆盖，默认 false。' }
      },
      output: envelopeOutput('导出工作簿'),
      async execute(args) {
        const requestId = newRequestId()
        const started = Date.now()
        try {
          const format = args.format ?? 'xlsx'
          const { wb, buffer } = await openWorkbook(args.path)
          const withExt = /\.[a-z0-9]+$/i.test(args.target_path) ? args.target_path : `${args.target_path}.${format}`
          const target = resolveSafePath(withExt, { workspaceRoot, mustExist: false })
          if (existsSync(target) && args.overwrite !== true) {
            throw new OfficeError('PERMISSION_DENIED', `目标文件已存在：${withExt}。如需覆盖请显式传 overwrite=true。`, { needsConfirmation: true })
          }
          const warnings = []

          let bytes
          let mime
          let summaryText
          if (format === 'csv') {
            const sheetName = args.sheet ?? wb.sheetNames()[0]
            bytes = workbookSheetToCsv(wb, sheetName)
            mime = 'text/csv'
            warnings.push({ code: 'LOSSY_CONVERT', message: 'CSV 为纯文本格式，公式、样式、图表、批注与其它工作表均不会保留。' })
            summaryText = `📤 已把「${sheetName}」导出为 CSV：${withExt}（${bytes.length} 字节，UTF-8 带 BOM）`
          } else {
            bytes = buffer
            mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            summaryText = `📤 已导出工作簿副本：${withExt}（${bytes.length} 字节）`
          }

          await mkdir(dirname(target), { recursive: true })
          await writeFile(target, bytes)
          return {
            ...ok({
              requestId,
              documentId: sha256(bytes),
              outputFile: { name: basename(target), path: target, size: bytes.length, mime_type: mime, sha256: sha256(bytes) },
              warnings,
              performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
              data: { format, source: basename(args.path), target: withExt }
            }),
            summary: summaryText
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `导出为 ${args.format ?? 'xlsx'}`, kind: 'write', rawInput: args })
    })
  )

  // ───────────────────────── DOCX 读取（阶段 3 起步）─────────────────────────

  tools.push(
    defineTool({
      name: 'office_read_docx',
      description:
        '读取 Word 文档（.docx）：元数据、段落（含样式与大纲级别）、表格、页眉页脚与页码域、以及结构识别结果（批注、修订、域代码、目录、书签、内容控件、脚注尾注、嵌入对象、宏）。当前为只读，写入能力在后续阶段提供。修订中「已删除」的文字不计入正文，「已插入」的计入。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        detail: {
          type: 'string',
          enum: ['summary', 'structure', 'paragraphs', 'tables', 'comments', 'revisions', 'text'],
          description:
            'summary=元数据+结构概览（默认）；structure=详细结构统计；paragraphs=段落列表；tables=表格内容；comments=批注（作者/日期/正文/锚点段落）；revisions=修订清单（插入/删除/段落标记增删与格式变更）；text=全文纯文本。'
        },
        offset: { type: 'integer', description: 'paragraphs 模式的起始下标，默认 0。' },
        limit: { type: 'integer', description: 'paragraphs 模式返回上限，默认 200。' }
      },
      output: envelopeOutput('读取 Word 文档'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const absolute = resolvePath(args.path, true)
          const { buffer } = await readDocument(absolute, { maxBytes })
          const doc = DocxDocument.open(buffer)
          const detail = args.detail ?? 'summary'
          const warnings = []
          if (doc.info.hasMacro) warnings.push({ code: 'MACRO_DETECTED', message: '文档含 VBA 宏。插件不会执行宏。' })
          if (doc.info.hasSignatures) warnings.push({ code: 'DIGITAL_SIGNATURE_PRESENT', message: '文档含数字签名。' })
          if (doc.info.embedded.length > 0) {
            warnings.push({ code: 'EMBEDDED_OBJECT_DETECTED', message: `文档含 ${doc.info.embedded.length} 个嵌入对象，插件不解析其内容。` })
          }

          const structure = doc.structure()
          if (structure.has_revisions) {
            warnings.push({ code: 'REVISIONS_PRESENT', message: '文档含未接受的修订；正文只包含「插入」的内容，不包含「删除」的内容。' })
          }
          if (structure.has_comments) {
            warnings.push({ code: 'COMMENTS_PRESENT', message: `文档含 ${structure.comments} 条批注，插件不修改批注。` })
          }
          if (structure.has_toc) warnings.push({ code: 'TOC_PRESENT', message: '文档含目录域，插件不会更新目录。' })
          if (structure.has_fields) warnings.push({ code: 'FIELDS_PRESENT', message: '文档含域代码，插件不会更新域。' })

          const base = {
            path: absolute,
            name: basename(absolute),
            metadata: doc.metadata(),
            structure
          }

          let data = base
          if (detail === 'paragraphs') {
            data = { ...base, ...doc.paragraphs({ offset: args.offset ?? 0, limit: args.limit ?? 200 }) }
          } else if (detail === 'tables') {
            data = { ...base, ...doc.tables() }
          } else if (detail === 'text') {
            data = { ...base, ...doc.text() }
          } else if (detail === 'structure') {
            data = { ...base, headers: doc.headersOrFooters('header'), footers: doc.headersOrFooters('footer') }
          } else if (detail === 'comments') {
            data = { ...base, ...doc.comments(), bookmarks: doc.bookmarks() }
          } else if (detail === 'revisions') {
            data = { ...base, ...doc.revisions() }
          }

          const summary =
            detail === 'paragraphs'
              ? `📄 ${basename(absolute)}｜段落 ${data.total} 个，返回第 ${data.offset} 起 ${data.returned} 个`
              : detail === 'tables'
                ? `📄 ${basename(absolute)}｜表格 ${data.count} 个`
                : detail === 'text'
                  ? `📄 ${basename(absolute)}｜全文 ${data.length} 字符${data.truncated ? '（已截断）' : ''}`
                  : detail === 'comments'
                    ? `💬 ${basename(absolute)}｜批注 ${data.count} 条${data.extended_parts.length ? `（另有 ${data.extended_parts.length} 个线程化批注附加部件未解析）` : ''}`
                    : detail === 'revisions'
                      ? `✍️ ${basename(absolute)}｜修订 ${data.count} 条${Object.keys(data.by_type).length ? `（${Object.entries(data.by_type).map(([k, v]) => `${k} ${v}`).join('、')}）` : ''}，格式变更 ${data.formatting_count} 处`
                      : `📄 ${basename(absolute)}｜段落 ${structure.paragraphs}、表格 ${structure.tables}、页眉 ${structure.headers}、页脚 ${structure.footers}${structure.has_comments ? `、批注 ${structure.comments}` : ''}${structure.has_revisions ? '、含修订' : ''}${warnings.length ? `｜${warnings.length} 条安全提示` : ''}`

          return {
            ...ok({ requestId, documentId: sha256(buffer), warnings, performance: null, data }),
            summary
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '读取 Word 文档', kind: 'read', rawInput: args.path })
    })
  )

  // ───────────────────────── PPTX 读取（阶段 4 起步）─────────────────────────

  tools.push(
    defineTool({
      name: 'office_read_pptx',
      description:
        '读取 PowerPoint 演示文稿（.pptx）：幻灯片尺寸与宽高比、母版/版式/主题、每页的标题与形状文本、演讲者备注、隐藏幻灯片、动画与切换、图表/SmartArt/音视频/嵌入对象。备注只取讲稿占位符，不混入页码。当前为只读，写入能力在后续阶段提供。',
      parameters: {
        path: { type: 'string', required: true, description: '演示文稿路径。' },
        detail: {
          type: 'string',
          enum: ['summary', 'structure', 'slides', 'slide', 'layouts', 'text'],
          description:
            'summary=概览（默认）；structure=详细结构统计；slides=全部幻灯片明细；slide=单张幻灯片（需 index）；layouts=版式清单与每页当前版式（改版式前先看这个）；text=全文纯文本。'
        },
        index: { type: 'integer', description: 'detail=slide 时的幻灯片下标（0 基）。' }
      },
      output: envelopeOutput('读取演示文稿'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const absolute = resolvePath(args.path, true)
          const { buffer } = await readDocument(absolute, { maxBytes })
          const pres = PptxPresentation.open(buffer)
          const detail = args.detail ?? 'summary'
          const report = pres.validate()
          const structure = pres.structure()

          const warnings = []
          if (structure.hidden_slides.length > 0) {
            warnings.push({ code: 'HIDDEN_SLIDES', message: `含 ${structure.hidden_slides.length} 张隐藏幻灯片（下标 ${structure.hidden_slides.join(', ')}）。` })
          }
          if (structure.animated_slides > 0) {
            warnings.push({ code: 'UNSUPPORTED_FEATURE', message: `${structure.animated_slides} 张幻灯片含动画，插件不解析也不修改动画。` })
          }
          if (structure.smartart_parts > 0) {
            warnings.push({ code: 'UNSUPPORTED_FEATURE', message: `含 ${structure.smartart_parts} 个 SmartArt 部件，插件保留但不解析其内容。` })
          }
          if (structure.media_kinds.video > 0 || structure.media_kinds.audio > 0) {
            warnings.push({
              code: 'UNSUPPORTED_FEATURE',
              message: `含 ${structure.media_kinds.video} 个视频 / ${structure.media_kinds.audio} 个音频，插件保留但不处理。`
            })
          }
          if (structure.embedded_objects > 0) {
            warnings.push({ code: 'EMBEDDED_OBJECT_DETECTED', message: `含 ${structure.embedded_objects} 个嵌入对象，插件不解析。` })
          }
          for (const w of report.warnings) warnings.push({ code: 'REVIEW_REQUIRED', message: w })

          const base = { path: absolute, name: basename(absolute), metadata: pres.metadata() }
          let data = { ...base, structure }
          if (detail === 'slides') {
            const { count, slides } = pres.slides()
            data = { ...base, structure, count, slides }
          } else if (detail === 'slide') {
            if (args.index === undefined) throw new OfficeError('INVALID_REQUEST', 'detail=slide 时必须提供 index。')
            data = { ...base, structure, slide: pres.readSlide(args.index) }
          } else if (detail === 'text') {
            data = { ...base, structure, ...pres.text() }
          } else if (detail === 'layouts') {
            data = {
              ...base,
              structure,
              layout_count: pres.layouts().length,
              layouts: pres.layouts(),
              slides: pres.slideLayouts(),
              theme_count: pres.themes().length,
              themes: pres.themes(),
              master_themes: pres.masterThemes()
            }
          }

          const summary =
            detail === 'slides'
              ? `📽️ ${basename(absolute)}｜${data.count} 张幻灯片（${structure.slide_size.aspect}）${structure.hidden_slides.length ? `，其中 ${structure.hidden_slides.length} 张隐藏` : ''}`
              : detail === 'slide'
                ? `📽️ 第 ${(args.index ?? 0) + 1} 张｜标题「${data.slide.title ?? '（无）'}」｜${data.slide.shape_count} 个形状${data.slide.notes ? '｜含备注' : ''}`
                : detail === 'text'
                  ? `📽️ ${basename(absolute)}｜全文 ${data.length} 字符${data.truncated ? '（已截断）' : ''}`
                  : `📽️ ${basename(absolute)}｜${structure.slide_count} 张幻灯片（${structure.slide_size.aspect}）｜${structure.masters} 母版/${structure.layouts} 版式｜${structure.notes_slides} 页备注${warnings.length ? `｜${warnings.length} 条提示` : ''}`

          return { ...ok({ requestId, documentId: sha256(buffer), warnings, performance: null, data }), summary }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '读取演示文稿', kind: 'read', rawInput: args.path })
    })
  )

  tools.push(
    defineTool({
      name: 'office_validate_pptx',
      description:
        '校验 PowerPoint 演示文稿结构：ZIP 完整性、必需部件、幻灯片是否可解析、关系是否有效、图片/媒体引用能否解析、母版与主题是否保留，并报告动画/SmartArt/音视频/嵌入对象。**不做**文本溢出、元素重叠、空白幻灯片等需要渲染才能判定的检查——会如实在 not_checked 中列出。',
      parameters: {
        path: { type: 'string', required: true, description: '演示文稿路径。' }
      },
      output: envelopeOutput('校验演示文稿'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const absolute = resolvePath(args.path, true)
          const { buffer } = await readDocument(absolute, { maxBytes })
          const report = PptxPresentation.open(buffer).validate()
          const warnings = report.warnings.map((message) => ({ code: 'REVIEW_REQUIRED', message }))
          const failed = report.checks.filter((c) => !c.ok)
          return {
            ...ok({
              requestId,
              documentId: sha256(buffer),
              warnings,
              performance: null,
              data: {
                valid: report.valid,
                checks: report.checks,
                not_checked: report.not_checked,
                slide_size: report.slide_size
              }
            }),
            summary: `${report.valid ? '✅' : '❌'} 校验${report.valid ? '通过' : '失败'}：${report.checks.filter((c) => c.ok).length}/${report.checks.length} 项通过${failed.length ? `（未通过：${failed.map((c) => c.name).join('、')}）` : ''}`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '校验演示文稿', kind: 'read', rawInput: args.path })
    })
  )

  /**
   * 带锁与事务地执行一次 PPTX 修改。
   *
   * 与 XLSX / DOCX 走同一条事务流程，因此「修改失败不破坏原文件」对 PPTX 同样成立。
   *
   * @param {object} args - 参数。
   * @returns {Promise<object>} 响应信封。
   */
  async function mutatePptx({ path, requestId, mutate, summaryOf }) {
    const absolute = resolvePath(path, true)
    const lock = await FileLock.acquire(absolute)
    const started = Date.now()
    try {
      const { buffer } = await readDocument(absolute, { maxBytes })
      const pres = PptxPresentation.open(buffer)
      const warnings = []
      if (pres.info.hasMacro) warnings.push({ code: 'MACRO_DETECTED', message: '演示文稿含 VBA 宏。插件不会执行宏，宏内容按原字节保留。' })
      if (pres.info.hasSignatures) warnings.push({ code: 'DIGITAL_SIGNATURE_PRESENT', message: '演示文稿含数字签名，任何修改都会使签名失效。' })
      const structure = pres.structure()
      if (structure.animated_slides > 0) {
        warnings.push({ code: 'UNSUPPORTED_FEATURE', message: `${structure.animated_slides} 张幻灯片含动画；改文本不会动动画，但动画引用的对象若被删除会失效。` })
      }

      const changes = []
      const extra = await mutate({ pres, changes, warnings })

      const temp = await (await store()).temp('pptx')
      const transaction = await Transaction.begin(absolute, temp.path)
      for (const w of warnings) transaction.warn(w)
      try {
        const result = await transaction.commit(pres.save(), {
          overwrite: true,
          verify: (out) => {
            const reopened = PptxPresentation.open(out)
            const report = reopened.validate()
            if (!report.valid) {
              throw new OfficeError('OUTPUT_VALIDATION_FAILED', `输出校验未通过：${report.checks.filter((c) => !c.ok).map((c) => c.name).join('、')}`)
            }
            if (reopened.info.slides.length === 0) {
              throw new OfficeError('OUTPUT_VALIDATION_FAILED', '输出演示文稿没有任何幻灯片。')
            }
          }
        })
        const envelope = ok({
          requestId,
          documentId: sha256(buffer),
          outputFile: {
            name: basename(absolute),
            path: absolute,
            size: result.size,
            mime_type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            sha256: result.sha256
          },
          warnings: result.warnings,
          changes,
          performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
          data: { ...extra }
        })
        return { ...envelope, summary: summaryOf(result, changes, extra) }
      } finally {
        await transaction.rollback()
        await temp.dispose()
      }
    } finally {
      await lock.release()
    }
  }

  tools.push(
    defineTool({
      name: 'office_update_slide_text',
      description:
        '改写 PowerPoint 某张幻灯片里一个形状的文本。`\\n` 会分成多个段落，并复用原第一段的段落属性（项目符号层级、缩进、字号），因此不会把项目符号列表变成无格式文本。形状下标与 office_read_pptx(detail="slide") 返回的 shapes 顺序一致。若该形状含域（页码/日期）或超链接，默认拒绝——需显式传 allow_markup_loss=true。',
      parameters: {
        path: { type: 'string', required: true, description: '演示文稿路径。' },
        index: { type: 'integer', required: true, description: '幻灯片下标（0 基）。' },
        shape: { type: 'integer', required: true, description: '形状下标（0 基，与 read 返回的 shapes 同序）。' },
        text: { type: 'string', required: true, description: '新文本；\\n 分段。' },
        allow_markup_loss: { type: 'boolean', description: '明知会丢失该形状的域或超链接仍继续，默认 false。' }
      },
      output: envelopeOutput('改写幻灯片文本'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutatePptx({
            path: args.path,
            requestId,
            mutate: async ({ pres, changes }) => {
              const result = pres.updateSlideText({
                index: args.index,
                shape: args.shape,
                text: args.text,
                allowMarkupLoss: args.allow_markup_loss === true
              })
              changes.push(result)
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `📝 已改写第 ${extra.slide + 1} 张幻灯片第 ${extra.shape} 个形状：「${extra.from}」→「${extra.to}」（${extra.paragraphs} 段${extra.reused_paragraph_properties ? '，保留原段落属性' : ''}）`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `改写第 ${args.index + 1} 张文本`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_add_slide',
      description:
        '在演示文稿中新增一张空白幻灯片。新页继承参照页的版式（slideLayout），因此标题/正文占位符样式与母版保持一致。会一并补齐幻灯片部件、内容类型声明、演示文稿关系与 sldIdLst 条目——少任何一环 Word/PowerPoint 都会忽略或报损坏。',
      parameters: {
        path: { type: 'string', required: true, description: '演示文稿路径。' },
        after: { type: 'integer', description: '插入到第几张之后（0 基）；-1 表示插到最前；省略则追加到末尾。' },
        layout_of: { type: 'integer', description: '参照哪张幻灯片取版式，默认第 0 张。' }
      },
      output: envelopeOutput('新增幻灯片'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutatePptx({
            path: args.path,
            requestId,
            mutate: async ({ pres, changes }) => {
              const result = pres.addSlide({ after: args.after ?? null, layoutOf: args.layout_of ?? 0 })
              changes.push({ type: 'add_slide', index: result.index, part: result.part })
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `➕ 已新增第 ${extra.index + 1} 张幻灯片（${extra.part}，继承版式 ${basename(extra.layout)}），共 ${extra.slide_count} 张。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '新增幻灯片', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_delete_slide',
      description:
        '删除一张幻灯片。破坏性操作：必须显式传 confirm=true。会一并清理该幻灯片部件、关系部件与内容类型声明，不留孤儿部件。演示文稿至少保留一张幻灯片。',
      parameters: {
        path: { type: 'string', required: true, description: '演示文稿路径。' },
        index: { type: 'integer', required: true, description: '幻灯片下标（0 基）。' },
        confirm: { type: 'boolean', required: true, description: '必须为 true 才执行删除（人工确认位）。' }
      },
      output: envelopeOutput('删除幻灯片'),
      async execute(args) {
        const requestId = newRequestId()
        if (args.confirm !== true) {
          return failure(requestId, new OfficeError('PERMISSION_DENIED', '删除幻灯片是破坏性操作，需要显式传 confirm=true。', { needsConfirmation: true }))
        }
        try {
          return await mutatePptx({
            path: args.path,
            requestId,
            mutate: async ({ pres, changes }) => {
              const result = pres.deleteSlide({ index: args.index })
              changes.push({ type: 'delete_slide', index: result.index, part: result.removed_part })
              return result
            },
            summaryOf: (_s, _c, extra) => `🗑️ 已删除第 ${extra.index + 1} 张幻灯片（${extra.removed_part}），剩余 ${extra.slide_count} 张。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `删除第 ${args.index + 1} 张幻灯片`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_reorder_slides',
      description: '调整幻灯片顺序：把第 from 张移动到第 to 位（均 0 基）。只重排 sldIdLst 中的条目，不改动任何幻灯片内容、版式、母版或备注。',
      parameters: {
        path: { type: 'string', required: true, description: '演示文稿路径。' },
        from: { type: 'integer', required: true, description: '原下标（0 基）。' },
        to: { type: 'integer', required: true, description: '目标下标（0 基）。' }
      },
      output: envelopeOutput('调整幻灯片顺序'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutatePptx({
            path: args.path,
            requestId,
            mutate: async ({ pres, changes }) => {
              const result = pres.moveSlide({ from: args.from, to: args.to })
              changes.push({ type: 'move_slide', from: args.from, to: args.to, changed: result.changed })
              return result
            },
            summaryOf: (_s, _c, extra) =>
              extra.changed ? `🔀 已把第 ${extra.from + 1} 张移到第 ${extra.to + 1} 位。` : `🔀 位置未变化（第 ${extra.from + 1} 张已在原位）。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `重排：${args.from} → ${args.to}`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_set_theme',
      description:
        '切换母版使用的主题（配色与字体）。主题关联只由关系决定（`slideMasters/_rels/slideMasterN.xml.rels` 里 `…/theme` 关系的 Target），所以只改关系目标，不动版式、母版与幻灯片 XML。实现上会**克隆出一个新的主题部件**再挂到母版上：直接把母版指向别处已在用的主题部件（例如备注母版的主题）会让 PowerPoint **拒绝打开整个文件**。配色变化可直接用 PowerPoint COM 读 `Theme.ThemeColorScheme.Colors(5).RGB` 客观验证。用 office_read_pptx(detail="layouts") 可查主题清单与每个母版当前用的主题。',
      parameters: {
        path: { type: 'string', required: true, description: '演示文稿路径。' },
        master: { type: 'integer', description: '母版下标（0 基），默认 0。' },
        theme: { type: 'integer', description: '目标主题下标（0 基）。' },
        name: { type: 'string', description: '目标主题名（如「Office 主题」）；与 theme 二选一。' }
      },
      output: envelopeOutput('切换主题'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutatePptx({
            path: args.path,
            requestId,
            mutate: async ({ pres, changes }) => {
              const result = pres.setTheme({ master: args.master ?? 0, theme: args.theme, name: args.name })
              changes.push({ type: 'set_theme', master: result.master, theme: result.to })
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `🎨 母版 ${extra.master} 主题：「${extra.from.split('/').pop()}」→「${extra.to.split('/').pop()}」${extra.theme_name ? `（${extra.theme_name}）` : ''}｜配色效果请在 PowerPoint 里确认`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `换主题（母版 ${args.master ?? 0}）`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_set_slide_layout',
      description:
        '切换一张幻灯片使用的版式。幻灯片的版式关联**只由关系决定**（`slideN.xml.rels` 里 `…/slideLayout` 关系的 Target），所以改的是关系目标而不是幻灯片 XML。**不挪动也不删除已有形状**：占位符文本留在原处，若新版式没有对应占位符，它会以自由形状继续显示（PowerPoint 自己的行为也是如此）。先用 office_read_pptx(detail="layouts") 看版式名与下标。',
      parameters: {
        path: { type: 'string', required: true, description: '演示文稿路径。' },
        index: { type: 'integer', required: true, description: '幻灯片下标（0 基）。' },
        layout: { type: 'integer', description: '目标版式下标（0 基，见 detail=layouts）。' },
        name: { type: 'string', description: '目标版式名（如「空白」「仅标题」）；与 layout 二选一。' }
      },
      output: envelopeOutput('切换幻灯片版式'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutatePptx({
            path: args.path,
            requestId,
            mutate: async ({ pres, changes }) => {
              const result = pres.setSlideLayout({ index: args.index, layout: args.layout, name: args.name })
              changes.push({ type: 'set_slide_layout', slide: args.index, layout: result.to })
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `📐 第 ${extra.slide + 1} 张版式：「${extra.from.split('/').pop()}」→「${extra.to.split('/').pop()}」${extra.layout_name ? `（${extra.layout_name}）` : ''}`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `第 ${args.index + 1} 张换版式`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_insert_slide_image',
      description:
        '往 PowerPoint 某张幻灯片插入一张图片。图片类型按**真实字节**判定而非扩展名；未指定尺寸时用原始像素（96 DPI），未指定位置时**按幻灯片居中**（默认落左上角会盖住标题占位符）。适配器会一并补齐媒体部件、内容类型声明与幻灯片关系。',
      parameters: {
        path: { type: 'string', required: true, description: '演示文稿路径。' },
        index: { type: 'integer', required: true, description: '幻灯片下标（0 基）。' },
        image_path: { type: 'string', description: '图片文件路径。与 image_base64 二选一。' },
        image_base64: { type: 'string', description: '图片的 base64 内容。与 image_path 二选一。' },
        left_px: { type: 'number', description: '左边距（像素）；省略则水平居中。' },
        top_px: { type: 'number', description: '上边距（像素）；省略则垂直居中。' },
        width_px: { type: 'number', description: '显示宽度（像素）。' },
        height_px: { type: 'number', description: '显示高度（像素）。' },
        alt_text: { type: 'string', description: '替代文本。' }
      },
      output: envelopeOutput('插入幻灯片图片'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          if (!args.image_path && !args.image_base64) {
            throw new OfficeError('INVALID_REQUEST', '必须提供 image_path 或 image_base64 之一。')
          }
          let data
          let sourceLabel
          if (args.image_path) {
            const imageAbsolute = resolvePath(args.image_path, true)
            const info = await stat(imageAbsolute)
            if (info.size > 64 * 1024 * 1024) throw new OfficeError('MEMORY_LIMIT', `图片 ${info.size} 字节超过 64MB 上限。`)
            const { readFile } = await import('node:fs/promises')
            data = await readFile(imageAbsolute)
            sourceLabel = basename(imageAbsolute)
          } else {
            data = Buffer.from(String(args.image_base64).replace(/^data:[^,]+,/, ''), 'base64')
            sourceLabel = 'base64'
            if (data.length === 0) throw new OfficeError('INVALID_REQUEST', 'image_base64 解码后为空。')
          }
          const extension = detectImageType(data)

          return await mutatePptx({
            path: args.path,
            requestId,
            mutate: async ({ pres, changes }) => {
              const result = pres.insertSlideImage({
                index: args.index,
                data,
                extension,
                leftPx: args.left_px,
                topPx: args.top_px,
                widthPx: args.width_px,
                heightPx: args.height_px,
                altText: args.alt_text
              })
              changes.push({ type: 'insert_slide_image', slide: args.index, media_part: result.media_part })
              return { ...result, source: sourceLabel }
            },
            summaryOf: (_s, _c, extra) =>
              `🖼️ 已在第 ${extra.slide + 1} 张幻灯片插入 ${extra.source}（${extra.display_size_px.width}×${extra.display_size_px.height} px，居中于 ${extra.offset_emu.x},${extra.offset_emu.y}）→ ${extra.media_part}`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `第 ${args.index + 1} 张插入图片`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_add_slide_table',
      description:
        '往 PowerPoint 某张幻灯片插入一个表格。默认套用 PowerPoint 内置表格样式并标记首行为表头（不指定样式时表格会渲染成无边框裸文本）。行列参差不齐时按最大列数补齐。未指定位置时按幻灯片居中。写入后可用 office_read_pptx 读回表格内容核对。',
      parameters: {
        path: { type: 'string', required: true, description: '演示文稿路径。' },
        index: { type: 'integer', required: true, description: '幻灯片下标（0 基）。' },
        rows: {
          type: 'array',
          required: true,
          description: '二维文本数组，第一维是行。',
          items: { type: 'array', items: { type: 'string' } }
        },
        left_px: { type: 'number', description: '左边距（像素）；省略则水平居中。' },
        top_px: { type: 'number', description: '上边距（像素）；省略则垂直居中。' },
        width_px: { type: 'number', description: '总宽度（像素）；省略按列数估算。' },
        height_px: { type: 'number', description: '总高度（像素）；省略按行数估算。' },
        first_row_header: { type: 'boolean', description: '是否把首行当表头，默认 true。' }
      },
      output: envelopeOutput('插入幻灯片表格'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutatePptx({
            path: args.path,
            requestId,
            mutate: async ({ pres, changes }) => {
              const result = pres.addSlideTable({
                index: args.index,
                rows: args.rows,
                leftPx: args.left_px,
                topPx: args.top_px,
                widthPx: args.width_px,
                heightPx: args.height_px,
                firstRowHeader: args.first_row_header !== false
              })
              changes.push({ type: 'add_slide_table', slide: args.index, rows: result.rows, columns: result.columns })
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `📊 已在第 ${extra.slide + 1} 张幻灯片插入 ${extra.rows}×${extra.columns} 表格（偏移 ${extra.offset_emu.x},${extra.offset_emu.y}）`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `第 ${args.index + 1} 张插入表格`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_duplicate_slide',
      description:
        '复制 PowerPoint 中的一张幻灯片，副本插入到源幻灯片**紧后面**。会一并复制该页的关系部件（图片、图表等），但**丢掉备注页关系**——共用同一个备注部件会让 PowerPoint 报「需要修复」。新页使用工作簿内唯一的部件名与形状关系 id。',
      parameters: {
        path: { type: 'string', required: true, description: '演示文稿路径。' },
        index: { type: 'integer', required: true, description: '要复制的幻灯片下标（0 基）。' }
      },
      output: envelopeOutput('复制幻灯片'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutatePptx({
            path: args.path,
            requestId,
            mutate: async ({ pres, changes }) => {
              const result = pres.duplicateSlide({ index: args.index })
              changes.push({ type: 'duplicate_slide', slide: result.index, source: result.source })
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `📑 已复制第 ${extra.source + 1} 张幻灯片 → 新第 ${extra.index + 1} 张（共 ${extra.slide_count} 张${extra.notes_dropped ? '，已丢弃备注页关系' : ''}）`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `复制第 ${args.index + 1} 张幻灯片`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_add_text_box',
      description:
        '往 PowerPoint 某张幻灯片插入一个文本框。形状 id 全页唯一（否则 PowerPoint 报修复），`\\n` 分段，位置尺寸按像素（96 DPI）换算成 EMU。文本框是自选图形（WPS 里显示为 type=17），不带占位符继承，因此不套用版式字号。',
      parameters: {
        path: { type: 'string', required: true, description: '演示文稿路径。' },
        index: { type: 'integer', required: true, description: '幻灯片下标（0 基）。' },
        text: { type: 'string', required: true, description: '文本内容；`\\n` 分段。' },
        left_px: { type: 'number', description: '左边距（像素），默认 80。' },
        top_px: { type: 'number', description: '上边距（像素），默认 80。' },
        width_px: { type: 'number', description: '宽（像素），默认 400。' },
        height_px: { type: 'number', description: '高（像素），默认 100。' },
        font_size_pt: { type: 'number', description: '字号（磅），默认 18。' },
        bold: { type: 'boolean', description: '是否加粗，默认 false。' },
        align: { type: 'string', description: '水平对齐：left | center | right，默认 left。' }
      },
      output: envelopeOutput('插入文本框'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutatePptx({
            path: args.path,
            requestId,
            mutate: async ({ pres, changes }) => {
              const result = pres.addTextBox({
                index: args.index,
                text: args.text,
                leftPx: args.left_px,
                topPx: args.top_px,
                widthPx: args.width_px,
                heightPx: args.height_px,
                fontSizePt: args.font_size_pt,
                bold: args.bold === true,
                align: args.align ?? 'left'
              })
              changes.push({ type: 'add_text_box', slide: args.index, shape_id: result.shape_id })
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `🔤 已在第 ${extra.index + 1} 张幻灯片插入文本框（形状 id ${extra.shape_id}，${extra.position_px.width}×${extra.position_px.height} px，${extra.align}${extra.bold ? '，加粗' : ''}）`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `第 ${args.index + 1} 张插入文本框`, kind: 'write', rawInput: args })
    })
  )

  // ───────────────────────── PDF 读取（阶段 5 起步）─────────────────────────

  tools.push(
    defineTool({
      name: 'office_read_pdf',
      description:
        '读取 PDF：文档元数据（标题/作者/生成器/日期）、页面数量与尺寸、每页旋转与内容流概况、文本提取、以及结构识别（是否加密、数字签名、表单字段、批注、图片、字体、嵌入文件）。文本提取通过字体的 ToUnicode CMap 解码，中文不会乱码；`max_chars` 控制返回量，防止长文档撑爆上下文。当前为只读，写入能力在后续阶段提供。',
      parameters: {
        path: { type: 'string', required: true, description: 'PDF 文件路径。' },
        detail: {
          type: 'string',
          enum: ['summary', 'structure', 'pages', 'text', 'annotations', 'form'],
          description:
            'summary=元数据+概览（默认）；structure=详细结构统计；pages=每页尺寸与旋转；text=文本提取；annotations=批注清单；form=AcroForm 表单字段（名称/类型/当前值/选项）。'
        },
        page: { type: 'integer', description: 'detail=text 时只提取某一页（0 基），省略则提取全部页。' },
        max_chars: { type: 'integer', description: '文本字符上限，默认 200000。' }
      },
      output: envelopeOutput('读取 PDF'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const absolute = resolvePath(args.path, true)
          const { buffer } = await readDocument(absolute, { maxBytes })
          const pdf = PdfDocument.open(buffer)
          const detail = args.detail ?? 'summary'
          const structure = pdf.structure()
          const pageInfos = pdf.pageInfos()

          const warnings = []
          if (structure.encrypted) {
            warnings.push({
              code: 'PASSWORD_REQUIRED',
              message: `PDF 已加密（${structure.encryption?.algorithm ?? '未知算法'}），文本提取可能失败或为空。`,
              needsConfirmation: true
            })
          }
          if (structure.has_digital_signature) {
            warnings.push({ code: 'DIGITAL_SIGNATURE_PRESENT', message: 'PDF 含数字签名；插件只检测不验证，任何修改都会使签名失效。' })
          }
          if (structure.form_fields > 0) {
            warnings.push({ code: 'UNSUPPORTED_FEATURE', message: `含 ${structure.form_fields} 个表单字段，插件当前不填写表单。` })
          }
          if (structure.annotations > 0) {
            warnings.push({
              code: 'ANNOTATIONS_PRESENT',
              message: `含 ${structure.annotations} 条批注；插件只统计与**新增**便签批注（office_add_pdf_annotation），不修改既有批注。`
            })
          }

          const base = { path: absolute, name: basename(absolute), metadata: pdf.metadata() }
          let data = { ...base, structure, pages: pageInfos }
          if (detail === 'pages') {
            data = { ...base, structure, pages: pageInfos }
          } else if (detail === 'structure') {
            data = { ...base, structure, pages: pageInfos }
          } else if (detail === 'text') {
            const extracted = pdf.extractText({ maxChars: args.max_chars ?? 200000, page: args.page ?? null })
            data = { ...base, structure, page_count: structure.page_count, ...extracted }
          } else if (detail === 'form') {
            data = { ...base, structure, ...pdf.forms() }
          } else if (detail === 'annotations') {
            data = { ...base, structure, ...pdf.annotations() }
          }

          const summary =
            detail === 'text'
              ? `📄 ${basename(absolute)}｜提取 ${data.length} 字符（${data.pages} 页）${data.truncated ? '，已截断' : ''}${warnings.length ? `｜${warnings.length} 条提示` : ''}`
              : detail === 'pages'
                ? `📄 ${basename(absolute)}｜${pageInfos.length} 页：${pageInfos.slice(0, 5).map((p) => `${p.width_pt}×${p.height_pt}pt`).join('、')}${pageInfos.length > 5 ? ' 等' : ''}`
                : detail === 'form'
                  ? `📝 ${basename(absolute)}｜表单字段 ${data.field_count} 个${data.signature_fields ? `（含 ${data.signature_fields} 个签名字段）` : ''}`
                  : detail === 'annotations'
                  ? `💬 ${basename(absolute)}｜批注 ${data.count} 条，分布在 ${data.pages_with_annotations} 页`
                  : `📄 ${basename(absolute)}｜PDF ${pdf.info.version}｜${structure.page_count} 页｜${structure.font_count} 种字体${structure.encrypted ? '｜⚠️ 已加密' : ''}${warnings.length ? `｜${warnings.length} 条提示` : ''}`

          return { ...ok({ requestId, documentId: sha256(buffer), warnings, performance: null, data }), summary }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '读取 PDF', kind: 'read', rawInput: args.path })
    })
  )

  tools.push(
    defineTool({
      name: 'office_search_pdf',
      description:
        '在 PDF 里做**全文搜索**：返回每处命中的页码、页内字符偏移、行号与上下文。默认不区分大小写，查询按**字面量**匹配（正则元字符会被转义，查 `(a)`、`1.5` 不会变成正则）。位置是**页内字符偏移而非页面坐标**——给出 x/y 需要按文本矩阵与字体宽度算版面，那属于排版引擎，插件不假装能做到。',
      parameters: {
        path: { type: 'string', required: true, description: 'PDF 文件路径。' },
        query: { type: 'string', required: true, description: '要查找的文本（字面量）。' },
        case_sensitive: { type: 'boolean', description: '是否区分大小写，默认 false。' },
        max_results: { type: 'integer', description: '最多返回多少条命中（1–1000），默认 100。' },
        context_chars: { type: 'integer', description: '命中前后各取多少字符作上下文（0–1000），默认 40。' },
        page: { type: 'integer', description: '只搜某一页（0 基）；省略则搜索全部页。' }
      },
      output: envelopeOutput('搜索 PDF 文本'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const absolute = resolvePath(args.path, true)
          const { buffer } = await readDocument(absolute, { maxBytes })
          const pdf = PdfDocument.open(buffer)
          if (pdf.structure().encrypted) throw new OfficeError('PASSWORD_REQUIRED', 'PDF 已加密，无法提取文本。')

          const result = pdf.search({
            query: args.query,
            caseSensitive: args.case_sensitive === true,
            maxResults: args.max_results ?? 100,
            contextChars: args.context_chars ?? 40,
            page: args.page ?? null
          })
          const data = { path: absolute, name: basename(absolute), ...result }
          const summary = result.found
            ? `🔍 ${basename(absolute)}｜「${result.query}」命中 ${result.total_matches} 处（${result.pages_scanned} 页）${result.truncated ? `，仅返回前 ${result.returned_matches} 条` : ''}`
            : `🔍 ${basename(absolute)}｜「${result.query}」未命中（已搜 ${result.pages_scanned} 页 / ${result.chars_scanned} 字符）`
          return { ...ok({ requestId, documentId: sha256(buffer), warnings: [], performance: null, data }), summary }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `搜索「${args.query}」`, kind: 'read', rawInput: args.path })
    })
  )

  tools.push(
    defineTool({
      name: 'office_create_presentation',
      description:
        '从零创建一份 .pptx 演示文稿（不依赖任何模板文件）：程序化生成主题、母版、版式与首页，写齐内容类型、包级与部件级关系。配色与字体是本插件自己的（主题名 office-plugin），不复制 Office 文件里的内容。创建出来的是**活文件** —— 可以继续用 office_add_slide / office_add_text_box / office_add_slide_chart / office_set_slide_layout 等工具编辑。默认 16:9，首页版式可用 layout 选「标题+副标题」或纯空白。已存在同名文件时拒绝覆盖。',
      parameters: {
        path: { type: 'string', required: true, description: '要创建的 .pptx 路径；未写扩展名时补 .pptx。' },
        title: { type: 'string', description: '首页标题，默认「新建演示文稿」。' },
        subtitle: { type: 'string', description: '首页副标题（仅 title 版式用）。' },
        author: { type: 'string', description: '文档属性作者，默认 dsh-exp-office。' },
        layout: { type: 'string', description: '首页版式：title（默认，标题+副标题）或 blank（纯空白）。' },
        overwrite: { type: 'boolean', description: '目标已存在时是否覆盖，默认 false。' }
      },
      output: envelopeOutput('创建演示文稿'),
      async execute(args) {
        const requestId = newRequestId()
        const started = Date.now()
        try {
          const withExt = /\.[a-z0-9]+$/i.test(args.path) ? args.path : `${args.path}.pptx`
          const target = resolveSafePath(withExt, { workspaceRoot, mustExist: false })
          if (existsSync(target) && args.overwrite !== true) {
            throw new OfficeError('PERMISSION_DENIED', `目标文件已存在：${withExt}。如需覆盖请显式传 overwrite=true。`, {
              needsConfirmation: true
            })
          }
          const built = buildBlankPptx({
            title: args.title ?? '新建演示文稿',
            subtitle: args.subtitle ?? '',
            author: args.author ?? 'dsh-exp-office',
            layout: args.layout ?? 'title'
          })
          await mkdir(dirname(target), { recursive: true })
          await writeFile(target, built.bytes)
          // 生成后立刻用自家读取器复核一遍：文件结构必须自洽，否则宁可报错也不留一个坏文件
          const check = PptxPresentation.open(built.bytes)
          const report = check.validate()
          if (!report.valid) {
            throw new OfficeError('OUTPUT_VALIDATION_FAILED', '生成的演示文稿未通过自检，已中止。', {
              checks: report.checks.filter((c) => !c.ok)
            })
          }
          return {
            ...ok({
              requestId,
              documentId: sha256(built.bytes),
              outputFile: {
                name: basename(target),
                path: target,
                size: built.bytes.length,
                mime_type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                sha256: sha256(built.bytes)
              },
              warnings: [],
              performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
              data: {
                path: target,
                slides: check.structure().slide_count,
                layout: args.layout ?? 'title',
                slide_size: check.structure().slide_size,
                masters: check.structure().masters,
                layouts: check.structure().layouts,
                themes: check.structure().themes,
                theme_name: check.themes()[0]?.name ?? null
              }
            }),
            summary: `📽️ 已创建演示文稿：${basename(target)}（${built.bytes.length} 字节｜${check.structure().slide_count} 页｜16:9｜主题 office-plugin）`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '创建演示文稿', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_add_slide_chart',
      description:
        '往 PowerPoint 某张幻灯片插入**原生图表**（柱状/条形/折线/饼图）。数据既写字面量缓存也写单元格引用，并**嵌入一份由本插件自己生成的 xlsx 工作簿**接上 `c:externalData`，因此在 PowerPoint 里可以直接「编辑数据」。适配器会写齐图表部件、嵌入工作簿、图表关系、幻灯片关系、`p:graphicFrame` 与内容类型五处接线。',
      parameters: {
        path: { type: 'string', required: true, description: '演示文稿路径。' },
        index: { type: 'integer', required: true, description: '幻灯片下标（0 基）。' },
        type: { type: 'string', description: 'column（默认）/ bar / line / pie。' },
        title: { type: 'string', description: '图表标题。' },
        categories: { type: 'array', items: { type: 'string' }, description: '分类标签；省略时用 1,2,3…' },
        values: {
          type: 'array',
          required: true,
          items: { type: 'array', items: { type: 'number' } },
          description:
            '二维数字数组：第一维是系列、第二维是该系列的值。例如 `[[120,150,180]]` 是一个系列三个值；饼图只支持一个系列。'
        },
        series_names: {
          type: 'array',
          items: { type: 'string' },
          description: '系列名，与 values 一一对应；省略时用「系列1」「系列2」。'
        },
        left_px: { type: 'number', description: '左边距（像素）；省略则水平居中。' },
        top_px: { type: 'number', description: '上边距（像素）；省略则垂直居中。' },
        width_px: { type: 'number', description: '宽（像素），默认 480。' },
        height_px: { type: 'number', description: '高（像素），默认 320。' }
      },
      output: envelopeOutput('插入图表'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutatePptx({
            path: args.path,
            requestId,
            mutate: async ({ pres, changes }) => {
              const names = Array.isArray(args.series_names) ? args.series_names : []
              const series = (args.values ?? []).map((values, i) => ({ name: names[i] ?? null, values }))
              const result = pres.addSlideChart({
                index: args.index,
                type: args.type ?? 'column',
                title: args.title ?? null,
                categories: Array.isArray(args.categories) ? args.categories : [],
                series,
                leftPx: args.left_px,
                topPx: args.top_px,
                widthPx: args.width_px ?? 480,
                heightPx: args.height_px ?? 320
              })
              changes.push({ type: 'add_slide_chart', slide: args.index, chart: result.chart_part })
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `📈 已在第 ${extra.slide + 1} 张插入${extra.chart_type} 图表（${extra.series_count} 个系列 × ${extra.categories} 个分类）→ ${extra.chart_part}${extra.title ? `，标题「${extra.title}」` : ''}`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `第 ${args.index + 1} 张插入图表`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_add_pdf_annotation',
      description:
        '给 PDF 的某一页加一个「便签」批注（`/Subtype /Text`）。**批注文字支持中文** —— 便签由阅读器自己排版渲染，不需要在 PDF 里嵌入字体（页面上的水印/叠加文字才受「只能 ASCII」限制）。采用增量更新：新建批注对象并把引用追加到该页 `/Annots`，页面内容流与其它对象一个字节都不动；`/Annots` 是间接数组时会解引用后追加并把数组一并写回（直接覆盖会让原有批注全部消失）。位置用 PDF 坐标（原点在**左下角**，单位 pt），省略时放在页面左上角内侧。',
      parameters: {
        path: { type: 'string', required: true, description: 'PDF 文件路径。' },
        page: { type: 'integer', required: true, description: '页面下标（0 基）。' },
        text: { type: 'string', required: true, description: '批注内容（支持中文）。' },
        author: { type: 'string', description: '作者（写入 /T）。' },
        x: { type: 'number', description: '便签图标左下角 x（pt）；省略放左上角内侧。' },
        y: { type: 'number', description: '便签图标左下角 y（pt）；省略放左上角内侧。' },
        width: { type: 'number', description: '图标宽（pt），默认 24。' },
        height: { type: 'number', description: '图标高（pt），默认 24。' },
        open: { type: 'boolean', description: '是否默认展开，默认 false。' }
      },
      output: envelopeOutput('添加 PDF 批注'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutatePdf({
            path: args.path,
            requestId,
            mutate: async ({ pdf, changes }) => {
              const result = pdf.addAnnotation({
                page: args.page,
                text: args.text,
                author: args.author ?? null,
                x: args.x ?? null,
                y: args.y ?? null,
                width: args.width ?? 24,
                height: args.height ?? 24,
                open: args.open === true
              })
              changes.push({ type: 'add_annotation', page: args.page, subtype: 'Text' })
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `💬 已在第 ${extra.page + 1} 页加便签批注（${extra.text_length} 字符${extra.non_ascii ? '，含中文' : ''}${extra.has_author ? `，作者 ${extra.author}` : ''}）位置 ${extra.rect.join(',')}`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `第 ${args.page + 1} 页加批注`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_fill_pdf_form',
      description:
        '填写 PDF 的 AcroForm 表单字段（文本 / 复选框 / 单选组 / 下拉）。只改字段字典的 `/V`（按钮另改 `/AS`），并把 `/NeedAppearances` 置为 true 让阅读器重绘外观 —— **不生成外观流**（那需要把文字按字体度量渲染成 XObject，属于排版引擎的活；做不到时宁可让阅读器重绘，也不画一个错的）。复选框的开状态名从 `/AP /N` 读出来（可能是 `Yes` / `1` / `开`，不能猜）。只读字段默认拒绝，需 `force=true`；数字签名与按钮字段明确拒绝。先用 office_read_pdf(detail="form") 看清字段名与选项。',
      parameters: {
        path: { type: 'string', required: true, description: 'PDF 文件路径。' },
        fields: {
          type: 'object',
          additionalProperties: true,
          description: '字段名 → 值：文本用字符串，复选框用 true/false 或开状态名，下拉用选项值。'
        },
        force: { type: 'boolean', description: '允许改写标记为只读的字段，默认 false。' }
      },
      output: envelopeOutput('填写 PDF 表单'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          if (Object.keys(args.fields ?? {}).length === 0) throw new OfficeError('INVALID_REQUEST', 'fields 不能为空。')
          return await mutatePdf({
            path: args.path,
            requestId,
            mutate: async ({ pdf, changes, warnings }) => {
              const result = pdf.fillForm({ fields: args.fields, force: args.force === true })
              changes.push({ type: 'fill_form', fields: result.filled.map((f) => f.name) })
              for (const w of result.warnings) warnings.push(w)
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `📝 已填写 ${extra.filled_count} 个表单字段：${extra.filled.map((f) => `${f.name}=${String(f.value).slice(0, 12)}`).join('、')}｜外观由阅读器重绘`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '填写 PDF 表单', kind: 'write', rawInput: args.path })
    })
  )

  tools.push(
    defineTool({
      name: 'office_validate_pdf',
      description:
        '校验 PDF 结构：文件头与版本、对象扫描、页面对象可解析、页面尺寸可读、内容流可解码、加密状态。**不做**交叉引用表一致性、数字签名有效性、字体嵌入完整性与版面还原——这些会如实在 not_checked 中列出并说明原因，而不是假装通过。',
      parameters: {
        path: { type: 'string', required: true, description: 'PDF 文件路径。' }
      },
      output: envelopeOutput('校验 PDF'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const absolute = resolvePath(args.path, true)
          const { buffer } = await readDocument(absolute, { maxBytes })
          const report = PdfDocument.open(buffer).validate()
          const warnings = []
          if (report.structure.encrypted) {
            warnings.push({ code: 'PASSWORD_REQUIRED', message: 'PDF 已加密，内容流可能无法解码。' })
          }
          const failed = report.checks.filter((c) => !c.ok)
          return {
            ...ok({
              requestId,
              documentId: sha256(buffer),
              warnings,
              performance: null,
              data: { valid: report.valid, checks: report.checks, not_checked: report.not_checked, structure: report.structure }
            }),
            summary: `${report.valid ? '✅' : '❌'} 校验${report.valid ? '通过' : '失败'}：${report.checks.filter((c) => c.ok).length}/${report.checks.length} 项通过${failed.length ? `（未通过：${failed.map((c) => c.name).join('、')}）` : ''}`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '校验 PDF', kind: 'read', rawInput: args.path })
    })
  )

  /**
   * PDF 写操作：锁文件 → 增量写回 → 重新打开校验 → 原子替换。
   *
   * 与 OOXML 侧同一套事务模型，区别在于「最小修改」的落地方式：
   * OOXML 是只改目标字节区间，PDF 是**只在尾部追加**被修改对象的新版本、
   * 新的交叉引用段与带 `/Prev` 的 trailer —— 原有字节一个都不动。
   *
   * @param {object} args - 参数。
   * @param {string} args.path - 文件路径。
   * @param {string} args.requestId - 请求 ID。
   * @param {(ctx: object) => Promise<object>} args.mutate - 执行修改。
   * @param {(result: object, changes: object[], extra: object) => string} args.summaryOf - 摘要。
   * @param {(pdf: object, extra: object) => Buffer} [args.serialize] - 自定义写回方式；默认增量更新（`pdf.save()`）。
   * @returns {Promise<object>} 工具返回值。
   */
  async function mutatePdf({ path, requestId, mutate, summaryOf, serialize }) {    const absolute = resolvePath(path, true)
    const lock = await FileLock.acquire(absolute)
    const started = Date.now()
    try {
      const { buffer } = await readDocument(absolute, { maxBytes })
      const pdf = PdfDocument.open(buffer)
      const warnings = []
      const structure = pdf.structure()
      if (structure.encrypted) {
        // 加密文件的页面对象与交叉引用被加密保护，不带口令改写只会产出损坏文件。
        throw new OfficeError('PASSWORD_REQUIRED', 'PDF 已加密，无法在不知口令的情况下安全改写。', {
          needsConfirmation: true,
          solution: '请先解密（另存为不含密码的副本）后再修改。'
        })
      }
      if (structure.has_digital_signature) {
        warnings.push({ code: 'DIGITAL_SIGNATURE_PRESENT', message: 'PDF 含数字签名；任何修改都会使签名失效。' })
      }
      warnings.push({
        code: 'INCREMENTAL_UPDATE',
        message: '写回采用增量更新：原有字节不变，只在文件尾部追加新版本对象与新的交叉引用段。'
      })

      const changes = []
      const extra = await mutate({ pdf, changes, warnings })
      const next = serialize ? serialize(pdf, extra) : pdf.save()
      const nextPageCount = PdfDocument.open(next).pages().length

      const temp = await (await store()).temp('pdf')
      const transaction = await Transaction.begin(absolute, temp.path)
      for (const w of warnings) transaction.warn(w)
      try {
        const result = await transaction.commit(next, {
          overwrite: true,
          verify: (out) => {
            const reopened = PdfDocument.open(out)
            const report = reopened.validate()
            if (!report.valid) {
              throw new OfficeError('OUTPUT_VALIDATION_FAILED', '写回后的 PDF 未通过基本结构校验。', {
                checks: report.checks.filter((c) => !c.ok)
              })
            }
            if (reopened.pages().length === 0) {
              throw new OfficeError('OUTPUT_VALIDATION_FAILED', '写回后的 PDF 没有任何页面。')
            }
          }
        })
        const envelope = ok({
          requestId,
          documentId: sha256(buffer),
          outputFile: {
            name: basename(absolute),
            path: absolute,
            size: result.size,
            mime_type: 'application/pdf',
            sha256: result.sha256
          },
          warnings: result.warnings,
          changes,
          performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
          data: { ...extra, appended_bytes: next.length - buffer.length, page_count: nextPageCount }
        })
        return { ...envelope, summary: summaryOf(result, changes, extra) }
      } finally {
        await transaction.rollback()
        await temp.dispose()
      }
    } finally {
      await lock.release()
    }
  }

  tools.push(
    defineTool({
      name: 'office_rotate_pdf_pages',
      description:
        '旋转 PDF 页面。采用增量更新写回：原有字节一个都不动，只在文件尾部追加页面的新版本、新的交叉引用段与带 /Prev 的 trailer。`page` 省略时旋转全部页面；`mode=add` 在现有角度上累加，否则直接设为 `degrees`。加密 PDF 会被拒绝（改写会产出损坏文件），含数字签名会给出失效提示。',
      parameters: {
        path: { type: 'string', required: true, description: 'PDF 文件路径。' },
        degrees: { type: 'integer', required: true, description: '旋转角度（顺时针，90 的整数倍），如 90、180、270。' },
        page: { type: 'integer', description: '页面下标（0 基）；省略则旋转全部页面。' },
        mode: { type: 'string', enum: ['set', 'add'], description: 'set=设为该角度（默认）；add=在现有角度上累加。' }
      },
      output: envelopeOutput('旋转 PDF 页面'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutatePdf({
            path: args.path,
            requestId,
            mutate: async ({ pdf, changes }) => {
              if (!Number.isInteger(args.degrees) || args.degrees % 90 !== 0) {
                throw new OfficeError('INVALID_REQUEST', `degrees 必须是 90 的整数倍，实际 ${args.degrees}。`)
              }
              const indices = args.page === undefined ? pdf.pages().map((_, i) => i) : [args.page]
              for (const index of indices) {
                const current = pdf.pageInfos()[index]?.rotation ?? 0
                const target = args.mode === 'add' ? current + args.degrees : args.degrees
                changes.push(pdf.rotatePage({ page: index, degrees: target }))
              }
              return { rotated: changes.map((c) => c.page), page_count: pdf.pages().length }
            },
            summaryOf: (_s, changes, extra) =>
              `🔄 已旋转 ${changes.length} 页（共 ${extra.page_count} 页）：${changes
                .slice(0, 5)
                .map((c) => `${c.page + 1}: ${c.from}°→${c.to}°`)
                .join('、')}${changes.length > 5 ? ' 等' : ''}`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '旋转 PDF 页面', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_delete_pdf_page',
      description:
        '删除 PDF 中的一页：把该页从页面树的 /Kids 里摘掉并更新 /Count。采用增量更新，被删页面的对象仍留在文件里（增量更新不回收字节），但不再被页面树引用，因此不会再被任何阅读器当作文档的一页。只剩一页时拒绝删除（删完就不是有效文档了）。加密 PDF 会被拒绝。',
      parameters: {
        path: { type: 'string', required: true, description: 'PDF 文件路径。' },
        page: { type: 'integer', required: true, description: '要删除的页面下标（0 基）。' }
      },
      output: envelopeOutput('删除 PDF 页面'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutatePdf({
            path: args.path,
            requestId,
            mutate: async ({ pdf, changes }) => {
              changes.push(pdf.deletePage({ page: args.page }))
              return { removed: args.page, page_count: pdf.pages().length }
            },
            summaryOf: (_s, _c, extra) => `🗑️ 已删除第 ${extra.removed + 1} 页，文档现有 ${extra.page_count} 页。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '删除 PDF 页面', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_reorder_pdf_pages',
      description:
        '调整 PDF 页面顺序。两种写法：`order` 给出新页序（元素是原页序下标，如 [2,0,1]），或 `from`+`to` 把一页移到新位置。只改页面树的 /Kids 与 /Count，页面内容与资源完全不动，因此不会损伤任何页面内容。采用增量更新写回。加密 PDF 会被拒绝。',
      parameters: {
        path: { type: 'string', required: true, description: 'PDF 文件路径。' },
        order: {
          type: 'array',
          items: { type: 'integer' },
          description: '新页序：原页序下标的排列，长度必须等于页数。'
        },
        from: { type: 'integer', description: '单页移动：原下标（0 基）。' },
        to: { type: 'integer', description: '单页移动：目标下标（0 基）。' }
      },
      output: envelopeOutput('调整 PDF 页面顺序'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const hasOrder = Array.isArray(args.order)
          const hasMove = args.from !== undefined || args.to !== undefined
          if (!hasOrder && !hasMove) {
            throw new OfficeError('INVALID_REQUEST', '需要给出 order，或同时给出 from 与 to。')
          }
          if (hasMove && (args.from === undefined || args.to === undefined)) {
            throw new OfficeError('INVALID_REQUEST', 'from 与 to 必须同时给出。')
          }
          return await mutatePdf({
            path: args.path,
            requestId,
            mutate: async ({ pdf, changes }) => {
              const result = hasOrder
                ? pdf.reorderPages({ order: args.order })
                : pdf.movePage({ from: args.from, to: args.to })
              changes.push(result)
              return { order: pdf.pages().map((p) => `${p.num} ${p.gen}`), page_count: pdf.pages().length }
            },
            summaryOf: (_s, _c, extra) => `🔀 已重排页面，文档现有 ${extra.page_count} 页。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '调整 PDF 页面顺序', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_add_pdf_watermark',
      description:
        '给 PDF 加水印文字（英文/数字）。采用**叠加式编辑**：原内容流一个字节都不改，只新建一个绘制流追加到页面的 /Contents 之后，并用浅灰填充近似淡化（不使用透明度组）。字体用阅读器内置的 Helvetica，**不需要嵌入字体**，因此只支持 ASCII 可见字符——中文水印会被明确拒绝（需要嵌入字体子集，当前阶段不做）。默认斜向 45° 居中，可用 format={page}/{total} 之类占位符做页码式水印。',
      parameters: {
        path: { type: 'string', required: true, description: 'PDF 文件路径。' },
        text: { type: 'string', required: true, description: '水印文字（ASCII）；支持 {page} 与 {total} 占位符。' },
        position: { type: 'string', enum: ['center', 'top', 'bottom'], description: 'center=斜向居中（默认）；top/bottom=水平居中贴边。' },
        font_size: { type: 'integer', description: '字号（pt）；center 默认按页宽自动取值。' },
        gray: { type: 'number', description: '灰度 0（黑）–1（白），默认 0.9 的浅灰。' },
        margin: { type: 'integer', description: 'top/bottom 时距页边的距离（pt），默认 24。' },
        pages: {
          type: 'array',
          items: { type: 'integer' },
          description: '要加水的页下标（0 基）；省略表示全部页。'
        }
      },
      output: envelopeOutput('添加 PDF 水印'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutatePdf({
            path: args.path,
            requestId,
            mutate: async ({ pdf, changes }) => {
              const result = pdf.addTextOverlay({
                text: args.text,
                pages: Array.isArray(args.pages) ? args.pages : null,
                position: args.position ?? 'center',
                fontSize: args.font_size ?? null,
                gray: args.gray ?? 0.9,
                margin: args.margin ?? 24
              })
              changes.push(result)
              return { pages: result.pages, position: result.position, text: args.text }
            },
            summaryOf: (_s, _c, extra) =>
              `💧 已给 ${extra.pages.length} 页添加水印「${extra.text}」（叠加式，原内容流未改动）。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '添加 PDF 水印', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_add_pdf_page_numbers',
      description:
        '给 PDF 加页码。与水印同一套叠加式写入：新建绘制流追加到页面 /Contents 之后，原内容流不改动，字体用内置 Helvetica（无需嵌入）。页码默认放在页脚居中，`start_at` 可指定起始编号（例如从第 5 页开始标 1），页脚已有内容时可用 `margin` 抬高或改用 top。',
      parameters: {
        path: { type: 'string', required: true, description: 'PDF 文件路径。' },
        format: { type: 'string', description: '页码格式，默认「{page} / {total}」；支持 {page} 与 {total} 占位符。' },
        position: { type: 'string', enum: ['bottom', 'top'], description: '页脚（默认）或页眉。' },
        start_at: { type: 'integer', description: '第一页显示的编号，默认 1。' },
        font_size: { type: 'integer', description: '字号（pt），默认 10。' },
        gray: { type: 'number', description: '灰度 0（黑）–1（白），默认 0.2（近黑）。' },
        margin: { type: 'integer', description: '距页边距离（pt），默认 24。' },
        pages: {
          type: 'array',
          items: { type: 'integer' },
          description: '要加页码的页下标（0 基）；省略表示全部页。'
        }
      },
      output: envelopeOutput('添加 PDF 页码'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const startAt = args.start_at ?? 1
          if (!Number.isInteger(startAt)) {
            throw new OfficeError('INVALID_REQUEST', `start_at 必须是整数，实际 ${args.start_at}。`)
          }
          const offset = startAt - 1
          return await mutatePdf({
            path: args.path,
            requestId,
            mutate: async ({ pdf, changes }) => {
              const all = pdf.pages().length
              const targets = Array.isArray(args.pages) ? args.pages : null
              const format = args.format ?? '{page} / {total}'
              // 编号按「所选页中的第几页」顺排：start_at 是所选的**第一页**显示的编号。
              // {total} 始终是文档总页数（不是所选页数），便于「第 3 页 / 共 10 页」这类写法。
              const result = pdf.addTextOverlay({
                text: format,
                pages: targets,
                position: args.position ?? 'bottom',
                fontSize: args.font_size ?? 10,
                gray: args.gray ?? 0.2,
                margin: args.margin ?? 24,
                labelFor: (index, order) =>
                  format
                    .replace(/\{page\}/g, String(startAt + order))
                    .replace(/\{n\}/g, String(order + 1))
                    .replace(/\{total\}/g, String(all))
              })
              changes.push(result)
              return { pages: result.pages, start_at: startAt, format }
            },
            summaryOf: (_s, _c, extra) =>
              `🔢 已给 ${extra.pages.length} 页添加页码（格式「${extra.format}」，起始 ${extra.start_at}）。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '添加 PDF 页码', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_set_pdf_metadata',
      description:
        '写入 PDF 文档元数据：标题、作者、主题、关键词、创建程序、生成程序。采用增量更新，只改 /Info 字典（原本没有就新建并接进 trailer），页面内容与其余对象一律不动。中文等非 ASCII 文本按 PDF 规范写成 UTF-16BE + BOM 十六进制串。**验证边界**：本机没有能读取 PDF 元数据的第三方工具（Word 的 PDF 重排不映射元数据、WPS PDF 的 COM 接口无法自动化驱动），因此元数据写入只做结构校验与插件自身回读，未做第三方确认。',
      parameters: {
        path: { type: 'string', required: true, description: 'PDF 文件路径。' },
        title: { type: 'string', description: '标题（Title）。' },
        author: { type: 'string', description: '作者（Author）。' },
        subject: { type: 'string', description: '主题（Subject）。' },
        keywords: { type: 'string', description: '关键词（Keywords），逗号分隔。' },
        creator: { type: 'string', description: '创建程序（Creator）。' },
        producer: { type: 'string', description: '生成程序（Producer）。' }
      },
      output: envelopeOutput('写入 PDF 元数据'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutatePdf({
            path: args.path,
            requestId,
            mutate: async ({ pdf, changes }) => {
              const result = pdf.updateMetadata({
                title: args.title,
                author: args.author,
                subject: args.subject,
                keywords: args.keywords,
                creator: args.creator,
                producer: args.producer
              })
              changes.push(result)
              const after = pdf.metadata()
              return { fields: result.fields, created_info: result.created_info, metadata: after }
            },
            summaryOf: (_s, _c, extra) =>
              `🏷️ 已写入 ${extra.fields.length} 个 PDF 元数据字段：${extra.fields.map((f) => f.field).join('、')}${
                extra.created_info ? '（原文件没有 /Info，已新建）' : ''
              }`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '写入 PDF 元数据', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_optimize_pdf',
      description:
        '把 PDF 重写成一份紧凑的新文件（对象图重写 / 字节回收）：只保留从 trailer 出发可达的对象，重新编号并重建经典 xref 表，于是反复增量更新累积的旧对象、旧交叉引用段与对象流容器全部消失。内容流保持原压缩字节不动，页面、字体、图片、批注与元数据都保留。省略 output_path 时**原地改写**（走事务：暂存 → 重开校验 → 原子替换）；给出 output_path 时另存为新文件，源文件不动。加密 PDF 会被拒绝。',
      parameters: {
        path: { type: 'string', required: true, description: 'PDF 文件路径。' },
        output_path: { type: 'string', description: '另存目标路径；省略表示原地改写。' },
        overwrite: { type: 'boolean', description: '另存且目标已存在时是否覆盖，默认 false。' }
      },
      output: envelopeOutput('重写 PDF（回收字节）'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          if (args.output_path) {
            return await writePdfArtifact({
              requestId,
              target: args.output_path,
              overwrite: args.overwrite === true,
              produce: async () => {
                const absolute = resolvePath(args.path, true)
                const { buffer } = await readDocument(absolute, { maxBytes })
                const pdf = PdfDocument.open(buffer)
                const result = pdf.rewrite()
                const warnings = sizeGainWarning(buffer, result.buffer, '重写')
                return {
                  output: result.buffer,
                  documentId: sha256(buffer),
                  sourceBytes: buffer.length,
                  expectedPages: pdf.pages().length,
                  warnings,
                  data: {
                    action: '重写',
                    source: basename(absolute),
                    saved_bytes: buffer.length - result.buffer.length,
                    objects: result.objects,
                    page_count: pdf.pages().length
                  },
                  summary: `🧹 已重写并另存为 ${basename(args.output_path)}：${result.objects} 个可达对象，${
                    buffer.length
                  } → ${result.buffer.length} 字节。`
                }
              }
            })
          }
          return await mutatePdf({
            path: args.path,
            requestId,
            serialize: (pdf, extra) => {
              const result = pdf.rewrite()
              extra.objects = result.objects
              extra.rewritten_bytes = result.buffer.length
              return result.buffer
            },
            mutate: async ({ pdf, changes, warnings }) => {
              const before = pdf.pages().length
              changes.push({ type: 'rewrite', pages: before })
              warnings.push({
                code: 'OBJECT_GRAPH_REWRITTEN',
                message: '整份文件已按可达对象重写：对象被重新编号，对象流被展平为普通对象，交叉引用流换成经典 xref 表。'
              })
              return { page_count: before, objects: 0 }
            },
            summaryOf: (_s, _c, extra) =>
              `🧹 已重写 PDF（${extra.objects} 个可达对象，${extra.page_count} 页，${extra.appended_bytes} 字节变化）；对象重新编号、旧版本与对象流已清除。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '重写 PDF（回收字节）', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_split_pdf',
      description:
        '按页面拆分 PDF：把选中的页写成一份**新文件**，源文件不动。与「删页」不同，这里走对象图重写 —— 未选页面的对象与内容**不会留在字节里**，因此可以安全地对外分发。pages 同时决定新文件的页序（[2,0] 表示新第 1 页是原第 3 页）。加密 PDF 会被拒绝。',
      parameters: {
        path: { type: 'string', required: true, description: '源 PDF 路径。' },
        pages: {
          type: 'array',
          required: true,
          items: { type: 'integer' },
          description: '要保留的页下标（0 基），顺序即新文件的页序。'
        },
        output_path: { type: 'string', required: true, description: '输出 PDF 路径。' },
        overwrite: { type: 'boolean', description: '目标已存在时是否覆盖，默认 false。' }
      },
      output: envelopeOutput('拆分 PDF'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          if (!Array.isArray(args.pages) || args.pages.length === 0) {
            throw new OfficeError('INVALID_REQUEST', 'pages 必须是非空下标数组。')
          }
          return await writePdfArtifact({
            requestId,
            target: args.output_path,
            overwrite: args.overwrite === true,
            produce: async () => {
              const absolute = resolvePath(args.path, true)
              const { buffer } = await readDocument(absolute, { maxBytes })
              const pdf = PdfDocument.open(buffer)
              const result = pdf.rewrite({ keepPages: args.pages })
              const warnings = sizeGainWarning(buffer, result.buffer, '拆分')
              return {
                output: result.buffer,
                documentId: sha256(buffer),
                sourceBytes: buffer.length,
                expectedPages: args.pages.length,
                warnings,
                data: {
                  action: '拆分',
                  source: basename(absolute),
                  saved_bytes: buffer.length - result.buffer.length,
                  objects: result.objects,
                  page_count: args.pages.length,
                  dropped_pages: result.dropped,
                  pages: result.pages
                },
                summary: `✂️ 已拆出 ${args.pages.length} 页到 ${basename(args.output_path)}（丢弃 ${
                  result.dropped
                } 页，其内容不再留在字节里）；源文件未改动。`
              }
            }
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '拆分 PDF', kind: 'write', rawInput: args })
    })
  )

  /**
   * 重写/拆分后体积没变小就给一条提示（内容流本来就压缩，普通对象不再走对象流）。
   * @param {Buffer} before - 原字节。
   * @param {Buffer} after - 新字节。
   * @param {string} action - 动作名。
   * @returns {object[]} 警告列表。
   */
  function sizeGainWarning(before, after, action) {
    if (after.length < before.length) return []
    return [
      {
        code: 'NO_SIZE_GAIN',
        message: `${action}后体积没有变小（${before.length} → ${after.length} 字节）：内容流本来就已压缩，而普通对象不再走对象流压缩。`
      }
    ]
  }

  /**
   * 写出一个 PDF 产物：先写文件 → **重新打开它做校验**（页数与结构）→
   * 校验不过就删掉自己刚写的文件并报错，绝不留下半个产物。
   *
   * 重写另存、拆分、合并三个工具共用这条路径，区别只在 `produce()` 怎么造字节。
   *
   * @param {object} args - 参数。
   * @param {string} args.requestId - 请求 ID。
   * @param {string} args.target - 目标路径。
   * @param {boolean} args.overwrite - 是否允许覆盖。
   * @param {() => Promise<{output: Buffer, documentId: string, sourceBytes: number, expectedPages: number, data: object, warnings: object[], summary: string}>} args.produce - 造字节并给出元信息。
   * @returns {Promise<object>} 工具返回值。
   */
  async function writePdfArtifact({ requestId, target: targetArg, overwrite, produce }) {
    const started = Date.now()
    const target = resolveSafePath(targetArg, { workspaceRoot, mustExist: false })
    if (existsSync(target) && !overwrite) {
      throw new OfficeError('PERMISSION_DENIED', `目标文件已存在：${targetArg}。如需覆盖请显式传 overwrite=true。`, {
        needsConfirmation: true
      })
    }
    const produced = await produce()
    const output = produced.output
    await mkdir(dirname(target), { recursive: true })
    await writeFile(target, output)
    try {
      const reopened = PdfDocument.open((await readDocument(target, { maxBytes })).buffer)
      const report = reopened.validate()
      if (reopened.pages().length !== produced.expectedPages) {
        throw new OfficeError(
          'OUTPUT_VALIDATION_FAILED',
          `输出页数 ${reopened.pages().length} 与预期 ${produced.expectedPages} 不一致。`
        )
      }
      if (!report.valid) {
        throw new OfficeError('OUTPUT_VALIDATION_FAILED', '输出 PDF 未通过基本结构校验。', {
          checks: report.checks.filter((c) => !c.ok)
        })
      }
    } catch (err) {
      await rm(target, { force: true })
      throw err
    }
    return {
      ...ok({
        requestId,
        documentId: produced.documentId,
        outputFile: {
          name: basename(target),
          path: target,
          size: output.length,
          mime_type: 'application/pdf',
          sha256: sha256(output)
        },
        warnings: produced.warnings,
        performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
        data: { ...produced.data, target, output_bytes: output.length, source_bytes: produced.sourceBytes }
      }),
      summary: produced.summary
    }
  }

  tools.push(
    defineTool({
      name: 'office_merge_pdfs',
      description:
        '把多份 PDF 合并成一份新文件（源文件都不动）。合并是**跨文件对象图搬运**：每份文档的整体对象编号错开，内部引用跟着改写，页面树用嵌套写法挂在新建的 `/Pages` 根下，因此各份文档的页尺寸、字体与图片各自独立、互不干扰。**只搬运页面内容与资源**：目录（书签）、命名目标、表单、结构树与页面标签不会带过来（跨文档合并它们需要重映射，硬带只会得到互相打架的引用）；元数据取第一份。加密 PDF 会被拒绝。',
      parameters: {
        paths: {
          type: 'array',
          required: true,
          items: { type: 'string' },
          description: '要合并的 PDF 路径列表，顺序即页序（至少两份）。'
        },
        output_path: { type: 'string', required: true, description: '输出 PDF 路径。' },
        overwrite: { type: 'boolean', description: '目标已存在时是否覆盖，默认 false。' }
      },
      output: envelopeOutput('合并 PDF'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          if (!Array.isArray(args.paths) || args.paths.length < 2) {
            throw new OfficeError('INVALID_REQUEST', 'paths 至少需要两个 PDF 路径。')
          }
          return await writePdfArtifact({
            requestId,
            target: args.output_path,
            overwrite: args.overwrite === true,
            produce: async () => {
              const buffers = []
              const names = []
              for (const path of args.paths) {
                const absolute = resolvePath(path, true)
                buffers.push((await readDocument(absolute, { maxBytes })).buffer)
                names.push(basename(absolute))
              }
              const merged = PdfDocument.merge(buffers)
              const warnings = [
                {
                  code: 'MERGE_SCOPE',
                  message: '合并只带页面内容与资源：各份文档的书签目录、命名目标、表单、结构树与页面标签都不会保留。'
                }
              ]
              if (merged.removed_page_refs > 0) {
                warnings.push({
                  code: 'DANGLING_REFS_REMOVED',
                  message: `清理了 ${merged.removed_page_refs} 处指向被丢弃页面的引用（书签目标/链接批注/结构树）。`
                })
              }
              return {
                output: merged.buffer,
                documentId: sha256(Buffer.concat(buffers)),
                sourceBytes: buffers.reduce((sum, b) => sum + b.length, 0),
                expectedPages: merged.pages,
                warnings,
                data: {
                  action: '合并',
                  sources: names,
                  objects: merged.objects,
                  page_count: merged.pages
                },
                summary: `🔗 已合并 ${merged.sources} 份 PDF（共 ${merged.pages} 页，${merged.objects} 个对象）到 ${basename(
                  args.output_path
                )}；源文件均未改动。`
              }
            }
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '合并 PDF', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_extract_pdf_images',
      description:
        '提取 PDF 页面上的图片到目录。只解出 PDF 里**已经存在的图像流**，不做页面光栅化：JPEG（DCTDecode）与 JPEG 2000（JPXDecode）原样导出不重新编码；FlateDecode + 8 位分量会还原预测器后包成 PNG；索引色展开为 RGB，软掩码（/SMask）合成成带透明通道的 PNG。1/2/4 位分量、CMYK、CCITT/JBIG2 等形态会**明确跳过并给出原因**，而不是导出打不开的文件。同一张图被多页共用时只导出一次，并在结果里列出页码。',
      parameters: {
        path: { type: 'string', required: true, description: 'PDF 文件路径。' },
        output_dir: { type: 'string', required: true, description: '输出目录（不存在时自动创建）。' },
        pages: {
          type: 'array',
          items: { type: 'integer' },
          description: '只提取这些页（0 基）；省略表示全部页。'
        },
        prefix: { type: 'string', description: '文件名前缀，默认 page，最终形如 page1-Im7.png。' },
        overwrite: { type: 'boolean', description: '同名文件已存在时是否覆盖，默认 false（跳过并报告）。' }
      },
      output: envelopeOutput('提取 PDF 图片'),
      async execute(args) {
        const requestId = newRequestId()
        const started = Date.now()
        try {
          const absolute = resolvePath(args.path, true)
          const outDir = resolveSafePath(args.output_dir, { workspaceRoot, mustExist: false })
          await mkdir(outDir, { recursive: true })
          const { buffer } = await readDocument(absolute, { maxBytes })
          const pdf = PdfDocument.open(buffer)
          const { images, skipped } = pdf.extractImages({ pages: Array.isArray(args.pages) ? args.pages : null })
          const prefix = typeof args.prefix === 'string' && args.prefix.length > 0 ? args.prefix : 'page'
          const files = []
          const existing = []
          const warnings = []
          for (const image of images) {
            const safeName = String(image.name ?? 'image').replace(/[^A-Za-z0-9_-]/g, '_')
            const fileName = `${prefix}${image.pages[0]}-${safeName}.${image.extension}`
            const target = join(outDir, fileName)
            if (existsSync(target) && args.overwrite !== true) {
              existing.push({ name: fileName, reason: '同名文件已存在（未传 overwrite=true）' })
              continue
            }
            await writeFile(target, image.bytes)
            // 写完自检：类型与尺寸必须与刚才算出来的一致，否则删掉，不留坏文件
            const detected = detectImageType(image.bytes)
            const size = readImageSize(image.bytes, image.extension)
            if (!detected || size.width !== image.width || size.height !== image.height) {
              await rm(target, { force: true })
              throw new OfficeError('OUTPUT_VALIDATION_FAILED', `写出的图片未通过自检：${fileName}`)
            }
            files.push({
              name: fileName,
              path: target,
              format: image.format,
              width: image.width,
              height: image.height,
              bytes: image.bytes.length,
              pages: image.pages,
              has_alpha: Boolean(image.has_alpha)
            })
          }
          if (skipped.length > 0) {
            warnings.push({
              code: 'UNSUPPORTED_FEATURE',
              message: `有 ${skipped.length} 个图像对象被跳过（位深/色彩空间/过滤器不支持），原因见 skipped。`
            })
          }
          if (existing.length > 0) {
            warnings.push({ code: 'FILE_EXISTS', message: `有 ${existing.length} 个同名文件未覆盖。` })
          }
          if (images.length === 0 && skipped.length === 0) {
            warnings.push({ code: 'NO_IMAGES', message: '这份 PDF 的页面资源里没有图像对象。' })
          }
          return {
            ...ok({
              requestId,
              documentId: sha256(buffer),
              warnings,
              performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
              data: {
                action: '提取图片',
                source: basename(absolute),
                output_dir: outDir,
                image_count: files.length,
                files,
                skipped,
                skipped_existing: existing
              }
            }),
            summary: `🖼️ 从 ${basename(absolute)} 提取出 ${files.length} 张图片到 ${args.output_dir}${
              skipped.length ? `（${skipped.length} 个对象被跳过）` : ''
            }`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '提取 PDF 图片', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_validate_docx',
      description:
        '校验 Word 文档结构：ZIP 完整性、必需部件、document.xml 与页眉页脚是否可解析、关系目标是否存在、图片/超链接/页眉页脚引用能否解析、样式与编号引用是否有定义、目录与域代码状态、批注与修订是否保留，并可输出文档使用的字体清单。**不做**空白页、元素重叠、表格越界等需要渲染才能判定的检查——这些会如实在 not_checked 中列出，不会假装通过。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        baseline_path: { type: 'string', description: '可选：改动前的文档，用于对比部件级差异。' }
      },
      output: envelopeOutput('校验 Word 文档'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const absolute = resolvePath(args.path, true)
          const { buffer } = await readDocument(absolute, { maxBytes })
          const doc = DocxDocument.open(buffer)
          const report = doc.validate()
          const warnings = report.checks.filter((c) => c.warning === true).map((c) => ({ code: 'REVIEW_REQUIRED', message: `${c.name}：${c.detail}` }))

          let changedParts = null
          if (args.baseline_path) {
            const { buffer: baseBuffer } = await readDocument(resolvePath(args.baseline_path, true), { maxBytes })
            const basePkg = ZipPackage.open(baseBuffer)
            const pkg = ZipPackage.open(buffer)
            changedParts = []
            for (const name of basePkg.names()) {
              if (!pkg.has(name)) changedParts.push({ part: name, change: 'removed' })
              else if (!pkg.read(name).equals(basePkg.read(name))) changedParts.push({ part: name, change: 'modified' })
            }
            for (const name of pkg.names()) {
              if (!basePkg.has(name)) changedParts.push({ part: name, change: 'added' })
            }
          }

          const failed = report.checks.filter((c) => !c.ok)
          return {
            ...ok({
              requestId,
              documentId: sha256(buffer),
              warnings,
              performance: null,
              data: {
                valid: report.valid,
                checks: report.checks,
                not_checked: report.not_checked,
                fonts_used: report.fonts,
                ...(changedParts ? { changed_parts: changedParts } : {})
              }
            }),
            summary: `${report.valid ? '✅' : '❌'} 校验${report.valid ? '通过' : '失败'}：${report.checks.filter((c) => c.ok).length}/${report.checks.length} 项通过${failed.length ? `（未通过：${failed.map((c) => c.name).join('、')}）` : ''}${changedParts ? `｜与基线相比 ${changedParts.length} 个部件变化` : ''}`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '校验 Word 文档', kind: 'read', rawInput: args.path })
    })
  )

  // ───────────────────────── 阶段 6：本地 Office/WPS 联动 ─────────────────────────

  /**
   * 重开校验：确认引擎重写出来的文件仍然能被自己的读取器解析，且规模与源文件一致。
   *
   * 引擎重写会改变字节（也会改变很多我们无法逐项核对的内部细节），因此这里核对的是
   * **可核对的结构性事实**：能打开、工作簿/段落/幻灯片清单与源文件一致。做不到的就
   * 如实写在 not_verified 里，不假装「完全一致」。
   *
   * @param {Buffer} output - 引擎产出的字节。
   * @param {string} kind - 文件类型。
   * @param {object} source - `{sheets, paragraphs, slides}` 源文件的规模。
   * @returns {object} `{ok, checks, not_verified}`。
   */
  function verifyEngineOutput(output, kind, source) {
    const checks = []
    if (kind === 'xlsx' || kind === 'xlsm') {
      const wb = Workbook.open(output)
      const sheets = wb.sheetNames()
      checks.push({ name: '工作簿可打开', ok: true, detail: `${sheets.length} 张表` })
      checks.push({
        name: '工作表清单与源文件一致',
        ok: JSON.stringify(sheets) === JSON.stringify(source.sheets),
        detail: `${sheets.join(',')} ↔ 源 ${(source.sheets ?? []).join(',')}`
      })
    } else if (kind === 'docx' || kind === 'docm') {
      const doc = DocxDocument.open(output)
      const total = doc.paragraphs().total
      checks.push({ name: '文档可打开', ok: true, detail: `${total} 个段落` })
      checks.push({ name: '段落数与源文件一致', ok: total === source.paragraphs, detail: `${total} ↔ 源 ${source.paragraphs}` })
    } else {
      const deck = PptxPresentation.open(output)
      const slides = deck.slides({ includeShapes: false }).count
      checks.push({ name: '演示文稿可打开', ok: true, detail: `${slides} 页` })
      checks.push({ name: '幻灯片数与源文件一致', ok: slides === source.slides, detail: `${slides} ↔ 源 ${source.slides}` })
    }
    return {
      ok: checks.every((c) => c.ok),
      checks,
      not_verified: [
        '引擎重写后与源文件的字节差异（引擎会重排内部结构，本插件不逐字节比对）',
        '视觉与排版是否发生变化（需要渲染成图片才能判定）',
        '引擎自己新增/删除的部件（如 printerSettings、theme 微调）未逐项核对'
      ]
    }
  }

  /**
   * 记录源文件的可核对规模（供重开校验比对）。
   * @param {Buffer} buffer - 源文件字节。
   * @param {string} kind - 文件类型。
   * @returns {object} `{sheets, paragraphs, slides}`。
   */
  function sourceShape(buffer, kind) {
    if (kind === 'xlsx' || kind === 'xlsm') return { sheets: Workbook.open(buffer).sheetNames() }
    if (kind === 'docx' || kind === 'docm') return { paragraphs: DocxDocument.open(buffer).paragraphs().total }
    const deck = PptxPresentation.open(buffer)
    return { slides: deck.slides({ includeShapes: false }).count }
  }

  /**
   * 用本地引擎做一次「打开 → 处理 → 另存」，并把产物按事务写回目标路径。
   * @param {object} args - 工具参数。
   * @param {string} requestId - 请求标识。
   * @param {'recalc'|'rerender'} action - 动作。
   * @returns {Promise<object>} 响应信封。
   */
  async function runEngineRewrite(args, requestId, action) {
    const started = Date.now()
    const absolute = resolvePath(args.path, true)
    const { buffer, type } = await readDocument(absolute, { maxBytes })
    const kind = type.ext
    if (!AUTOMATION_KINDS.includes(kind)) {
      throw new OfficeError('UNSUPPORTED_FILE_TYPE', `本地引擎联动只支持 ${AUTOMATION_KINDS.join(' / ')}，实际是 ${kind}。`, {
        supported: [...AUTOMATION_KINDS]
      })
    }
    if (type.hasMacro) {
      // 宏文件可以交给引擎重存（保持宏格式），但绝不执行宏
    }
    const suffix = action === 'recalc' ? 'recalculated' : 'rerendered'
    const rawTarget = args.target_path ?? `${basename(absolute).replace(/\.[^.]+$/, '')}-${suffix}.${kind}`
    const withExt = /\.[a-z0-9]+$/i.test(rawTarget) ? rawTarget : `${rawTarget}.${kind}`
    const target = args.target_path ? resolveSafePath(withExt, { workspaceRoot, mustExist: false }) : (await store()).outputPath(withExt)
    if (target === absolute) {
      throw new OfficeError('INVALID_REQUEST', '输出路径不能与输入相同：本地引擎从不改动源文件，请换一个 target_path。')
    }
    if (existsSync(target) && args.overwrite !== true) {
      throw new OfficeError('PERMISSION_DENIED', `目标文件已存在：${withExt}。如需覆盖请显式传 overwrite=true。`, {
        needsConfirmation: true
      })
    }

    const lock = await FileLock.acquire(absolute)
    const temp = await (await store()).temp('automation')
    try {
      const engineOutput = temp.file(`engine-output.${kind}`)
      const run = action === 'recalc' ? automation.recalculate.bind(automation) : automation.rerender.bind(automation)
      const result = await run({
        inputPath: absolute,
        outputPath: engineOutput,
        kind,
        engine: args.engine ?? 'auto',
        timeoutMs: args.timeout_ms ?? automation.timeoutMs
      })
      const bytes = await readFile(engineOutput)
      const shape = sourceShape(buffer, kind)
      const verification = verifyEngineOutput(bytes, kind, shape)
      const warnings = []
      for (const w of result.warnings ?? []) warnings.push({ code: 'ENGINE_WARNING', message: w })
      warnings.push({
        code: 'ENGINE_REWRITE',
        message: '产物由本地引擎重写整份文件（不是本插件的最小修改）：内部部件顺序、主题与打印设置等细节可能与源文件不同，源文件未被改动。'
      })
      if (action === 'recalc') {
        warnings.push({ code: 'FORMULA_CACHE_REFRESHED', message: '公式缓存与域已由本地引擎重新计算（这是 FORMULA_NOT_RECALCULATED 的正解）。' })
      }
      if (type.hasMacro) {
        warnings.push({ code: 'MACRO_NOT_EXECUTED', message: '文件含宏：已按宏格式重存，宏代码按字节保留，但**没有执行**任何宏。' })
      }
      if ((result.leftover_pids ?? []).length > 0) {
        warnings.push({
          code: 'PROCESS_RESIDUE',
          message: `本次启动的引擎进程仍有残留：PID ${result.leftover_pids.join(', ')}（插件只清理自己启动的进程，不会去杀用户开着的 Office/WPS）。`
        })
      }
      if (!verification.ok) {
        throw new OfficeError('OUTPUT_VALIDATION_FAILED', `引擎产物未通过重开校验：${verification.checks.filter((c) => !c.ok).map((c) => c.name).join('、')}`, {
          checks: verification.checks
        })
      }

      const transaction = await Transaction.begin(target, temp.path)
      for (const w of warnings) transaction.warn(w)
      const committed = await transaction.commit(bytes, { overwrite: true })
      const relativeTarget = relative(workspaceRoot, target)
      return {
        ...ok({
          requestId,
          documentId: sha256(buffer),
          outputFile: {
            name: basename(target),
            path: relativeTarget === '' ? target : relativeTarget,
            size: committed.size,
            mime_type:
              kind === 'xlsx' || kind === 'xlsm'
                ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                : kind === 'docx' || kind === 'docm'
                  ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                  : 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            sha256: committed.sha256
          },
          warnings: committed.warnings,
          performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
          data: {
            action,
            engine: result.engine,
            engine_version: result.engine_version,
            engine_elapsed_ms: result.elapsed_ms,
            source_unchanged: true,
            source_sha256: sha256(buffer),
            verification,
            audit_logged: true
          }
        }),
        summary: `🔄 ${action === 'recalc' ? '重算' : '重渲染'}完成（${result.engine} ${result.engine_version ?? ''}）｜${result.elapsed_ms} ms｜输出 ${committed.size} 字节｜源文件未改动`
      }
    } finally {
      await temp.dispose()
      await lock.release()
    }
  }

  tools.push(
    defineTool({
      name: 'office_extract_pdf_tables',
      description:
        '从 PDF 里提取**表格**：按文本片段的 y 坐标聚类成行、x 坐标聚类成列（不依赖框线），返回二维单元格文本。适用于文本可选、列大致对齐的表格（Office/LibreOffice 导出的表格、报表类 PDF）。**不做**：跨页表格合并、合并单元格的跨行跨列还原、扫描件与纯图片表格（需要 OCR，当前不支持）、单元格内换行。提取结果里的 `not_done` 会如实列出这些边界。',
      parameters: {
        path: { type: 'string', required: true, description: 'PDF 路径。' },
        page: { type: 'integer', description: '只处理某一页（0 基）；省略处理全部页。' },
        row_tolerance: { type: 'number', description: '同一行的 y 容差（点）；省略则按行距自动推导（推荐省略）。' },
        column_gap: { type: 'number', description: '同一列的 x 容差（点），默认 8。' },
        min_rows: { type: 'integer', description: '至少几行才算表格，默认 2。' },
        min_columns: { type: 'integer', description: '至少几列才算表格，默认 2。' }
      },
      output: envelopeOutput('提取 PDF 表格'),
      async execute(args) {
        const requestId = newRequestId()
        const started = Date.now()
        try {
          const absolute = resolvePath(args.path, true)
          const { buffer } = await readDocument(absolute, { maxBytes })
          const pdf = PdfDocument.open(buffer)
          const result = pdf.extractTables({
            page: args.page ?? null,
            rowTolerance: args.row_tolerance ?? null,
            columnGap: args.column_gap ?? 8,
            minRows: args.min_rows ?? 2,
            minColumns: args.min_columns ?? 2
          })
          const warnings = []
          if (result.table_count === 0) {
            warnings.push({
              code: 'NO_TABLE_FOUND',
              message: '没有找到符合「至少 2 行 2 列」的表格。扫描件或纯图片表格没有文本层，需要 OCR（当前不支持）。'
            })
          }
          return {
            ...ok({
              requestId,
              documentId: sha256(buffer),
              warnings,
              performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
              data: { path: absolute, name: basename(absolute), ...result }
            }),
            summary:
              result.table_count === 0
                ? `📄 ${basename(absolute)}｜未找到表格`
                : `📄 ${basename(absolute)}｜找到 ${result.table_count} 张表：${result.tables.map((t) => `第 ${t.page + 1} 页 ${t.row_count}×${t.column_count}`).join('、')}`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '提取 PDF 表格', kind: 'read', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_create_pdf',
      description:
        '**从零生成**一份 PDF（不依赖任何模板或第三方库）：自写页面树、内容流、字体资源与 xref，同时写入 /ToUnicode，因此生成的文本**可被提取**（自己与第三方解析器都能读回）。拉丁文本用 PDF 标准 14 字体（无需嵌入）；**含中文等非 WinAnsi 字符时自动嵌入一份子集化的 TrueType 字体**（保留原字形编号 + CIDToGIDMap，只带用到的字形，因此文件仍在几十 KB 量级）。支持自动折行与分页。',
      parameters: {
        path: { type: 'string', required: true, description: '目标 PDF 路径。' },
        lines: { type: 'array', items: { type: 'string' }, required: true, description: "正文行；'' 表示空行；行首加 \\f 强制分页。" },
        page_size: { type: 'string', enum: ['A4', 'A3', 'A5', 'Letter', 'Legal'], description: '纸张，默认 A4。' },
        orientation: { type: 'string', enum: ['portrait', 'landscape'], description: '方向，默认 portrait。' },
        margin_pt: { type: 'number', description: '页边距（点），默认 56.7（2 cm）。' },
        font_size_pt: { type: 'number', description: '字号（点），默认 11。' },
        font_family: { type: 'string', enum: ['helvetica', 'times', 'courier'], description: '拉丁字族，默认 helvetica。' },
        line_height_pt: { type: 'number', description: '行距（点），默认字号 × 1.4。' },
        cjk_font_path: { type: 'string', description: '嵌入用的中文字体文件（.ttf/.ttc）；省略时自动找本机的 SimHei / 等线 / 仿宋 / 楷体 / 微软雅黑 / 宋体 / Noto Sans SC。' },
        title: { type: 'string', description: '文档标题（写入 /Info）。' },
        author: { type: 'string', description: '作者（写入 /Info）。' },
        subject: { type: 'string', description: '主题（写入 /Info）。' },
        keywords: { type: 'string', description: '关键词（写入 /Info）。' },
        compress: { type: 'boolean', description: '是否压缩内容流，默认 true。' },
        overwrite: { type: 'boolean', description: '目标已存在时是否覆盖，默认 false。' }
      },
      output: envelopeOutput('从零生成 PDF'),
      async execute(args) {
        const requestId = newRequestId()
        const started = Date.now()
        try {
          const rawTarget = args.path
          const withExt = /\.pdf$/i.test(rawTarget) ? rawTarget : `${rawTarget}.pdf`
          const target = resolveSafePath(withExt, { workspaceRoot, mustExist: false })
          if (existsSync(target) && args.overwrite !== true) {
            throw new OfficeError('PERMISSION_DENIED', `目标文件已存在：${withExt}。如需覆盖请显式传 overwrite=true。`, {
              needsConfirmation: true
            })
          }
          const built = buildPdf({
            lines: args.lines,
            pageSize: args.page_size ?? 'A4',
            orientation: args.orientation ?? 'portrait',
            marginPt: args.margin_pt ?? 56.7,
            fontSizePt: args.font_size_pt ?? 11,
            fontFamily: args.font_family ?? 'helvetica',
            lineHeightPt: args.line_height_pt ?? null,
            cjkFontPath: args.cjk_font_path ?? null,
            metadata: {
              title: args.title ?? null,
              author: args.author ?? null,
              subject: args.subject ?? null,
              keywords: args.keywords ?? null
            },
            compress: args.compress !== false
          })

          const temp = await (await store()).temp('pdf')
          try {
            const transaction = await Transaction.begin(target, temp.path)
            const committed = await transaction.commit(built.buffer, {
              overwrite: true,
              // 写完必须自己读回来：页数与首页第一行文本都要对得上，否则不留文件
              verify: (bytes) => {
                const doc = PdfDocument.open(bytes)
                const pages = doc.pages().length
                if (pages !== built.pages) {
                  throw new OfficeError('OUTPUT_VALIDATION_FAILED', `生成的 PDF 读回 ${pages} 页，与预期 ${built.pages} 页不一致。`)
                }
                const firstText = args.lines.find((line) => line.trim() !== '')
                if (firstText) {
                  const sample = firstText.replace(/^\f/, '').trim().slice(0, 20)
                  const extracted = doc.extractText().text
                  if (sample !== '' && !extracted.includes(sample)) {
                    throw new OfficeError('OUTPUT_VALIDATION_FAILED', `生成的 PDF 读不回第一行文本「${sample}」。`)
                  }
                }
              }
            })
            const warnings = []
            if (built.lines > args.lines.length) {
              warnings.push({
                code: 'PDF_TEXT_WRAPPED',
                message: `有行超出可用宽度被自动折行：传入 ${args.lines.length} 行，实际排版 ${built.lines} 行。`
              })
            }
            if (built.embedded_font) {
              warnings.push({
                code: 'FONT_EMBEDDED',
                message: `正文含非 WinAnsi 字符，已嵌入字体子集「${built.embedded_font.name}」（${built.embedded_font.characters} 个字符 / ${built.embedded_font.glyphs} 个字形，未压缩 ${Math.round(built.embedded_font.subset_bytes / 1024)} KB）。字体来自本机，只保留用到的字形。`
              })
            }
            return {
              ...ok({
                requestId,
                outputFile: {
                  name: basename(target),
                  path: relative(workspaceRoot, target) === '' ? target : relative(workspaceRoot, target),
                  size: committed.size,
                  mime_type: 'application/pdf',
                  sha256: committed.sha256
                },
                warnings,
                performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
                data: {
                  pages: built.pages,
                  lines: built.lines,
                  input_lines: args.lines.length,
                  font: args.font_family ?? 'helvetica',
                  font_size_pt: args.font_size_pt ?? 11,
                  page_size: args.page_size ?? 'A4',
                  orientation: args.orientation ?? 'portrait',
                  encoding: built.embedded_font ? 'Identity-H（嵌入子集字体）' : 'WinAnsiEncoding（标准 14 字体，无需嵌入）',
                  embedded_font: built.embedded_font,
                  text_extractable: true
                }
              }),
              summary: `📄 已生成 PDF：${basename(target)}｜${built.pages} 页 / ${built.lines} 行｜${committed.size} 字节${built.embedded_font ? `｜嵌入字体子集 ${built.embedded_font.name}` : ''}`
            }
          } finally {
            await temp.dispose()
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '从零生成 PDF', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_detect_engines',
      description:
        '检测本机装了哪些 Microsoft Office / WPS 组件及其版本。默认**只读注册表、不启动任何程序**；probe_com=true 会真的创建一次 COM 实例来验证自动化是否可用（需要插件配置 allowLocalAutomation=true，会短暂启动并随即关闭对应程序）。',
      parameters: {
        probe_com: { type: 'boolean', description: '是否实际创建 COM 实例验证（默认 false）。' }
      },
      output: envelopeOutput('检测本地 Office/WPS'),
      async execute(args) {
        const requestId = newRequestId()
        const started = Date.now()
        try {
          const result = await automation.detectEngines({ probeCom: args.probe_com === true })
          const installed = result.engines.filter((e) => e.installed)
          const warnings = []
          if (result.leftover_pids.length > 0) {
            warnings.push({
              code: 'PROCESS_RESIDUE',
              message: `检测后仍有引擎进程在运行：PID ${result.leftover_pids.join(', ')}。插件只清理自己启动的进程，不会去杀用户开着的 Office/WPS。`
            })
          }
          if (!result.microsoft_office && !result.wps) {
            warnings.push({
              code: 'OFFICE_NOT_INSTALLED',
              message: '未检测到本机 Office/WPS：重算与重渲染会返回 OFFICE_NOT_INSTALLED；插件其余能力不受影响（独立文件处理模式）。'
            })
          }
          const names = installed.map((e) => `${e.label} ${e.com_version ?? e.file_version ?? ''}`.trim())
          return {
            ...ok({
              requestId,
              warnings,
              performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
              data: {
                engines: result.engines,
                microsoft_office: result.microsoft_office,
                wps: result.wps,
                probed_com: result.probed_com,
                installed_count: installed.length,
                local_automation_enabled: automation.enabled
              }
            }),
            summary: `🔍 本机引擎：${installed.length > 0 ? names.join('、') : '未检测到 Office/WPS'}${result.probed_com ? '（已实测 COM 可用性）' : '（仅注册表检测）'}`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: () => ({ card: 'generic', title: '检测本地 Office/WPS', kind: 'read', rawInput: {} })
    })
  )

  tools.push(
    defineTool({
      name: 'office_recalculate',
      description:
        '用**本机真实的 Office/WPS** 打开文件的副本、重新计算（Excel 公式缓存、Word 域）后另存为新文件：这是「公式不重算」的正解。**源文件永不被改动**；引擎产物会再经本插件重开校验。需要插件配置 allowLocalAutomation=true；被禁宏、禁外链更新、禁弹窗，超时会被终止并清理本次启动的进程（只清理自己启动的）。',
      parameters: {
        path: { type: 'string', required: true, description: '源文件路径（xlsx/xlsm/docx/docm/pptx/pptm）。' },
        target_path: { type: 'string', description: '输出路径；省略时写到受管输出目录 <名字>-recalculated.<扩展名>。' },
        engine: { type: 'string', enum: ['auto', 'excel', 'word', 'powerpoint', 'wps'], description: '用哪个引擎，默认 auto（按类型选 Office，再退回 WPS）。' },
        timeout_ms: { type: 'integer', description: `超时毫秒数，默认 ${DEFAULT_AUTOMATION_TIMEOUT_MS}（5 秒–10 分钟）。` },
        overwrite: { type: 'boolean', description: '目标已存在时是否覆盖，默认 false。' }
      },
      output: envelopeOutput('本地引擎重算'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await runEngineRewrite(args, requestId, 'recalc')
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '本地引擎重算', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_rerender',
      description:
        '用**本机真实的 Office/WPS** 打开文件的副本并整份重写（等价于「用 Office 重新保存一遍」）后另存为新文件：用于修复第三方工具生成的、Office 能打开但会提示修复的文件。**源文件永不被改动**，产物会再经本插件重开校验。安全约定同 office_recalculate（禁宏、禁外链、禁弹窗、超时清理）。',
      parameters: {
        path: { type: 'string', required: true, description: '源文件路径（xlsx/xlsm/docx/docm/pptx/pptm）。' },
        target_path: { type: 'string', description: '输出路径；省略时写到受管输出目录 <名字>-rerendered.<扩展名>。' },
        engine: { type: 'string', enum: ['auto', 'excel', 'word', 'powerpoint', 'wps'], description: '用哪个引擎，默认 auto。' },
        timeout_ms: { type: 'integer', description: `超时毫秒数，默认 ${DEFAULT_AUTOMATION_TIMEOUT_MS}。` },
        overwrite: { type: 'boolean', description: '目标已存在时是否覆盖，默认 false。' }
      },
      output: envelopeOutput('本地引擎重渲染'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await runEngineRewrite(args, requestId, 'rerender')
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '本地引擎重渲染', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_compare_docx',
      description:
        '比较两份 Word 文档的**内容差异**（只读，不改文件、不写修订标记）：段落级新增/删除/修改（并区分是文字变了、样式变了还是两者都变）、表格个数与逐格文本变化、段落样式集合差异、元数据差异。段落用最长公共子序列对齐，因此中间插入一段不会把后面所有段落都报成「修改」。需要把差异**写进文档**（Word 修订标记）请另说——那是另一件事。',
      parameters: {
        path_a: { type: 'string', required: true, description: '作为「旧版」的 Word 文档路径。' },
        path_b: { type: 'string', required: true, description: '作为「新版」的 Word 文档路径。' },
        max_items: { type: 'integer', description: '每类差异最多返回多少条，默认 200（total_* 仍是真实总数）。' },
        include_tables: { type: 'boolean', description: '是否比较表格，默认 true。' },
        include_styles: { type: 'boolean', description: '是否比较段落样式集合，默认 true。' },
        include_metadata: { type: 'boolean', description: '是否比较元数据，默认 true。' }
      },
      output: envelopeOutput('比较两份 Word 文档'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const absoluteA = resolvePath(args.path_a, true)
          const absoluteB = resolvePath(args.path_b, true)
          const { buffer: bufferA, type: typeA } = await readDocument(absoluteA, { maxBytes })
          const { buffer: bufferB, type: typeB } = await readDocument(absoluteB, { maxBytes })
          for (const [absolute, type] of [
            [absoluteA, typeA],
            [absoluteB, typeB]
          ]) {
            if (!['docx', 'docm', 'dotx'].includes(type.ext)) {
              throw new OfficeError('UNSUPPORTED_FILE_TYPE', `比较只接受 Word 文档，${basename(absolute)} 识别为 ${type.ext}。`, {
                file: basename(absolute),
                detected: type.ext,
                supported: ['docx', 'docm', 'dotx']
              })
            }
          }
          const docA = DocxDocument.open(bufferA)
          const docB = DocxDocument.open(bufferB)
          const diff = docA.compare(docB, {
            maxItems: args.max_items ?? 200,
            includeTables: args.include_tables !== false,
            includeStyles: args.include_styles !== false,
            includeMetadata: args.include_metadata !== false
          })
          const p = diff.paragraphs
          const summary =
            `📄 ${basename(absoluteA)} ↔ ${basename(absoluteB)}｜` +
            (diff.identical
              ? '内容一致（段落/表格/样式/元数据均无差异）'
              : `段落 +${p.total_added} / -${p.total_removed} / ~${p.total_changed}｜表格 ${diff.tables.total_changed} 处｜样式 +${diff.styles.added.length} / -${diff.styles.removed.length}｜元数据 ${diff.metadata.changed.length} 处`) +
            (diff.truncated ? '（已截断，见 total_*）' : '')
          return {
            ...ok({
              requestId,
              documentId: sha256(bufferA),
              warnings: diff.warnings,
              performance: null,
              data: {
                path_a: absoluteA,
                path_b: absoluteB,
                sha256_a: sha256(bufferA),
                sha256_b: sha256(bufferB),
                ...diff
              }
            }),
            summary
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '比较 Word 文档', kind: 'read', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_export_docx',
      description:
        '把 Word 文档导出为新文件，不修改源文件。format=docx **原样复制字节**（留存当前状态/另存交付，内容与源文件逐字节一致）；format=text 导出纯文本（strip 掉样式与图片，带 BOM 的 UTF-8，记事本打开中文不乱码）。目标已存在时默认拒绝覆盖，需显式传 overwrite=true。',
      parameters: {
        path: { type: 'string', required: true, description: '源 Word 文档路径。' },
        target_path: { type: 'string', required: true, description: '导出目标路径；未写扩展名时按 format 自动补全（.docx / .txt）。' },
        format: { type: 'string', enum: ['docx', 'text'], description: '导出格式，默认 docx。' },
        overwrite: { type: 'boolean', description: '目标已存在时是否覆盖，默认 false。' }
      },
      output: envelopeOutput('导出 Word 文档'),
      async execute(args) {
        const requestId = newRequestId()
        const started = Date.now()
        try {
          const format = args.format ?? 'docx'
          const absolute = resolvePath(args.path, true)
          const { buffer, type } = await readDocument(absolute, { maxBytes })
          if (type.ext !== 'docx') {
            throw new OfficeError('UNSUPPORTED_FILE_TYPE', `Word 导出只接受 .docx，实际是 ${type.ext}。`, { supported: ['docx'] })
          }
          const extension = format === 'text' ? 'txt' : 'docx'
          const withExt = /\.[a-z0-9]+$/i.test(args.target_path) ? args.target_path : `${args.target_path}.${extension}`
          const target = resolveSafePath(withExt, { workspaceRoot, mustExist: false })
          if (existsSync(target) && args.overwrite !== true) {
            throw new OfficeError('PERMISSION_DENIED', `目标文件已存在：${withExt}。如需覆盖请显式传 overwrite=true。`, {
              needsConfirmation: true
            })
          }

          const warnings = []
          let bytes
          let mime
          if (format === 'text') {
            const { text } = DocxDocument.open(buffer).text()
            // 带 BOM：Windows 记事本/Excel 打开纯文本时靠它认 UTF-8，否则中文会乱码
            bytes = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(text, 'utf8')])
            mime = 'text/plain'
            warnings.push({
              code: 'LOSSY_CONVERT',
              message: '纯文本导出不含样式、编号、表格结构、图片、批注与修订；表格内容会按单元格文本平铺。'
            })
          } else {
            // 原样复制：不做任何解析与重写，保证与源文件逐字节一致
            bytes = buffer
            mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
          }

          await mkdir(dirname(target), { recursive: true })
          await writeFile(target, bytes)
          const identical = format === 'docx' ? bytes.equals(buffer) : null
          return {
            ...ok({
              requestId,
              documentId: sha256(buffer),
              outputFile: { name: basename(target), path: target, size: bytes.length, mime_type: mime, sha256: sha256(bytes) },
              warnings,
              performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
              data: {
                format,
                source: basename(absolute),
                target: withExt,
                byte_identical: identical,
                source_sha256: sha256(buffer)
              }
            }),
            summary:
              format === 'text'
                ? `📤 已导出纯文本：${withExt}（${bytes.length} 字节，UTF-8 带 BOM）｜⚠️ 样式/表格结构/图片不保留`
                : `📤 已导出 Word 副本：${withExt}（${bytes.length} 字节，与源文件逐字节一致）`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `导出为 ${args.format ?? 'docx'}`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_convert_document',
      description:
        '把 Office 文档（docx/xlsx/pptx 及其旧版二进制格式）转换为 PDF，使用宿主内置的 LibreOffice 引擎，不需要本机安装 Office。返回所用引擎（native/wasm）、页数，以及**缺失字体清单**——后者是判断是否发生字体替换的唯一可靠依据，读 XML 无法回答。转换不会覆盖已存在的输出文件。转换是重操作（引擎冷启动通常数秒），支持取消。',
      parameters: {
        path: { type: 'string', required: true, description: '输入文档路径。' },
        target_path: { type: 'string', description: '输出 PDF 路径；省略则在受管输出目录生成同名 .pdf。' },
        overwrite: { type: 'boolean', description: '输出已存在时是否覆盖，默认 false。' },
        timeout_ms: { type: 'integer', description: '超时毫秒数，默认 180000。' }
      },
      output: envelopeOutput('转换为 PDF'),
      async execute(args, exec) {
        const requestId = newRequestId()
        const started = Date.now()
        try {
          const absolute = resolvePath(args.path, true)
          const { type } = await readDocument(absolute, { maxBytes })
          if (!CONVERTIBLE_EXTENSIONS.includes(type.ext)) {
            throw new OfficeError('UNSUPPORTED_FILE_TYPE', `类型 ${type.ext} 不能转换为 PDF。`, { supported: CONVERTIBLE_EXTENSIONS })
          }
          const target = args.target_path
            ? resolveSafePath(/\.[a-z0-9]+$/i.test(args.target_path) ? args.target_path : `${args.target_path}.pdf`, { workspaceRoot, mustExist: false })
            : (await store()).outputPath(`${basename(absolute).replace(/\.[^.]+$/, '')}.pdf`)

          if (existsSync(target) && args.overwrite !== true) {
            throw new OfficeError('PERMISSION_DENIED', `输出文件已存在：${basename(target)}。如需覆盖请显式传 overwrite=true。`, {
              needsConfirmation: true
            })
          }
          if (existsSync(target)) {
            const { rm } = await import('node:fs/promises')
            await rm(target, { force: true })
          }

          const result = await convertToPdf({
            inputPath: absolute,
            outputPath: target,
            signal: exec?.signal,
            timeoutMs: args.timeout_ms
          })

          const warnings = []
          if (result.missingFonts.length > 0) {
            warnings.push({
              code: 'FONT_NOT_FOUND',
              message: `以下字体在本机未安装，PDF 中已发生替换：${result.missingFonts.join('、')}`,
              fonts: result.missingFonts
            })
          }
          if (result.backend === 'wasm') {
            warnings.push({ code: 'ENGINE_DEGRADED', message: '使用 WASM 引擎（未找到本机原生引擎），转换质量与速度可能低于原生。' })
          }

          return {
            ...ok({
              requestId,
              outputFile: {
                name: basename(target),
                path: target,
                size: result.size,
                mime_type: 'application/pdf',
                sha256: sha256(await (await import('node:fs/promises')).readFile(target))
              },
              warnings,
              performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
              data: {
                source: basename(absolute),
                source_type: type.ext,
                engine: result.backend,
                pages: result.pages,
                missing_fonts: result.missingFonts
              }
            }),
            summary: `📕 已转换为 PDF：${basename(target)}（${result.size} 字节，${result.pages ?? '?'} 页，引擎 ${result.backend}）${result.missingFonts.length ? `｜⚠️ ${result.missingFonts.length} 种字体缺失已替换：${result.missingFonts.join('、')}` : '｜字体齐全'}`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '转换为 PDF', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_export_presentation',
      description:
        '把 PowerPoint 演示文稿导出为 PDF（宿主内置 LibreOffice 引擎，不需要本机装 Office）。与 office_convert_document 的区别在于**只接受 pptx 且会先读一遍演示文稿**：返回幻灯片数、被丢弃的东西（动画/切换/音视频/嵌入对象）与缺失字体清单，并读回生成的 PDF 页数以核对「一页幻灯片 = 一页 PDF」。**有损**：动画、切换、媒体与嵌入式对象不会出现在 PDF 里；显示效果由引擎排版决定，可能与 PowerPoint 有细微差异。不覆盖已存在的输出文件。',
      parameters: {
        path: { type: 'string', required: true, description: '演示文稿路径（.pptx）。' },
        target_path: { type: 'string', description: '输出 PDF 路径；省略则在受管输出目录生成同名 .pdf。' },
        overwrite: { type: 'boolean', description: '输出已存在时是否覆盖，默认 false。' },
        timeout_ms: { type: 'integer', description: '超时毫秒数，默认 180000。' }
      },
      output: envelopeOutput('导出演示文稿为 PDF'),
      async execute(args, exec) {
        const requestId = newRequestId()
        const started = Date.now()
        try {
          const absolute = resolvePath(args.path, true)
          const { buffer, type } = await readDocument(absolute, { maxBytes })
          if (type.ext !== 'pptx') {
            throw new OfficeError('UNSUPPORTED_FILE_TYPE', `演示文稿导出只接受 .pptx，实际是 ${type.ext}。`, { supported: ['pptx'] })
          }
          // 先读一遍：把「导不进去的东西」在导出前就说清楚，而不是导出后让用户自己发现
          const pres = PptxPresentation.open(buffer)
          const structure = pres.structure()
          const lost = []
          if (structure.animated_slides > 0) lost.push(`${structure.animated_slides} 张幻灯片动画`)
          if (structure.slides_with_transition > 0) lost.push(`${structure.slides_with_transition} 张切换效果`)
          if (structure.media_kinds.video > 0) lost.push(`${structure.media_kinds.video} 个视频`)
          if (structure.media_kinds.audio > 0) lost.push(`${structure.media_kinds.audio} 个音频`)
          if (structure.embedded_objects > 0) lost.push(`${structure.embedded_objects} 个嵌入对象`)
          if (structure.hidden_slides.length > 0) lost.push(`${structure.hidden_slides.length} 张隐藏幻灯片不会出现在 PDF 里`)

          const target = args.target_path
            ? resolveSafePath(/\.[a-z0-9]+$/i.test(args.target_path) ? args.target_path : `${args.target_path}.pdf`, {
                workspaceRoot,
                mustExist: false
              })
            : await (async () => (await store()).outputPath(`${basename(absolute).replace(/\.[^.]+$/, '')}.pdf`))()

          if (existsSync(target) && args.overwrite !== true) {
            throw new OfficeError('PERMISSION_DENIED', `输出文件已存在：${basename(target)}。如需覆盖请显式传 overwrite=true。`, {
              needsConfirmation: true
            })
          }
          if (existsSync(target)) {
            const { rm } = await import('node:fs/promises')
            await rm(target, { force: true })
          }

          const result = await convertToPdf({
            inputPath: absolute,
            outputPath: target,
            signal: exec?.signal,
            timeoutMs: args.timeout_ms
          })

          const warnings = []
          if (lost.length > 0) {
            warnings.push({ code: 'LOSSY_CONVERT', message: `以下内容不会出现在 PDF 里：${lost.join('、')}`, lost })
          }
          if (result.missingFonts.length > 0) {
            warnings.push({ code: 'FONT_NOT_FOUND', message: `以下字体未安装，PDF 中已发生替换：${result.missingFonts.join('、')}`, fonts: result.missingFonts })
          }
          if (result.backend === 'wasm') {
            warnings.push({ code: 'ENGINE_DEGRADED', message: '使用 WASM 引擎（未找到本机原生引擎），质量与速度可能低于原生。' })
          }

          // 读回生成的 PDF 核对页数：一页幻灯片应对应一页 PDF（隐藏页除外）
          let pdfPages = null
          try {
            const pdfBytes = await (await import('node:fs/promises')).readFile(target)
            pdfPages = PdfDocument.open(pdfBytes).pages().length
          } catch {
            pdfPages = null
          }
          const visibleSlides = structure.slide_count - structure.hidden_slides.length
          if (pdfPages !== null && pdfPages !== visibleSlides) {
            warnings.push({
              code: 'PAGE_COUNT_MISMATCH',
              message: `PDF 页数 ${pdfPages} 与可见幻灯片数 ${visibleSlides} 不一致，请人工核对分页。`
            })
          }

          return {
            ...ok({
              requestId,
              outputFile: {
                name: basename(target),
                path: target,
                size: result.size,
                mime_type: 'application/pdf',
                sha256: sha256(await (await import('node:fs/promises')).readFile(target))
              },
              warnings,
              performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
              data: {
                source: basename(absolute),
                slides: structure.slide_count,
                hidden_slides: structure.hidden_slides.length,
                engine: result.backend,
                pdf_pages: pdfPages,
                missing_fonts: result.missingFonts,
                lost
              }
            }),
            summary: `📕 已导出 PDF：${basename(target)}（${structure.slide_count} 张幻灯片 → ${pdfPages ?? '?'} 页，引擎 ${result.backend}）${lost.length ? `｜⚠️ 有损：${lost.join('、')}` : ''}${result.missingFonts.length ? `｜${result.missingFonts.length} 种字体被替换` : ''}`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '导出演示文稿为 PDF', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_insert_bookmark',
      description:
        '在指定段落上插入一个书签（`w:bookmarkStart` + `w:bookmarkEnd` 一对标记，`w:id` 全文档唯一）。书签名需以字母（含中文）或下划线开头，后续只能是文字、数字、下划线，最长 40 字符，且不区分大小写地在文档内唯一——重名或重号会让目录、交叉引用与书签跳转失效。只插入标记，不改段落文本与样式。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        paragraph: { type: 'integer', required: true, description: '段落下标（0 基，只数正文直接段落）。' },
        name: { type: 'string', required: true, description: '书签名（如 intro、章节_1）。' }
      },
      output: envelopeOutput('插入书签'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes }) => {
              const result = doc.insertBookmark({ paragraph: args.paragraph, name: args.name })
              changes.push(result)
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `🔖 已在第 ${extra.paragraph + 1} 段插入书签「${extra.name}」（id ${extra.id}，文档共 ${extra.bookmark_count} 个）：${extra.text.slice(0, 20)}`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `第 ${args.paragraph + 1} 段插书签`, kind: 'write', rawInput: args })
    })
  )

  /**
   * 声明一个页眉/页脚写入工具。
   * @param {string} toolName - 工具名。
   * @param {'header'|'footer'} kind - 类型。
   * @param {string} label - 中文标签。
   * @returns {object} 工具定义。
   */
  const headerFooterTool = (toolName, kind, label) =>
    defineTool({
      name: toolName,
      description: `设置 Word 文档的${label}文本，可选在末尾附加页码域（Word 会自动按页更新，插件不写死页码）。文档还没有${label}时会自动新建部件并接线（媒体部件 → 内容类型 Override → 关系 → sectPr 引用四处一次补齐）。只改写${label}部件的段落内容，文档其它部分不受影响。`,
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        text: { type: 'string', required: true, description: `${label}文本。` },
        page_number: { type: 'boolean', description: '是否在文本后附加 PAGE 页码域，默认 false。' },
        index: { type: 'integer', description: `第几个${label}部件（0 基），默认 0 即默认${label}。` }
      },
      output: envelopeOutput(`设置${label}`),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes }) => {
              const result = doc.setHeaderFooter({
                kind,
                text: args.text,
                pageNumber: args.page_number === true,
                index: args.index ?? 0
              })
              changes.push({ type: result.type, part: result.part, from: result.from, to: result.to })
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `📄 已设置${label}：「${extra.from}」→「${extra.to}」${extra.page_number_field ? '（含自动页码域）' : ''}`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `设置${label}`, kind: 'write', rawInput: args })
    })

  tools.push(headerFooterTool('office_set_docx_header', 'header', '页眉'))
  tools.push(headerFooterTool('office_set_docx_footer', 'footer', '页脚'))

  tools.push(
    defineTool({
      name: 'office_set_docx_page_layout',
      description:
        '设置 Word 文档的纸张与方向（末节）。尺寸以 twips 存储并回读为厘米；方向通过交换宽高实现，横向时写 `w:orient="landscape"`。只改 body 直接子 `sectPr`（真正的最后一节）——段落内部的 `sectPr` 是分节符，改它只影响前面那一节。文档没有 `sectPr` 时会补一个。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        paper: { type: 'string', description: '纸张预设：A4（默认）/ A3 / A5 / Letter / Legal。' },
        orientation: { type: 'string', description: '方向：portrait（纵向，默认）或 landscape（横向）。' },
        width_cm: { type: 'number', description: '自定义页宽（厘米）；给定时覆盖纸张预设。' },
        height_cm: { type: 'number', description: '自定义页高（厘米）。' }
      },
      output: envelopeOutput('设置纸张与方向'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes }) => {
              const result = doc.setPageLayout({
                paper: args.paper,
                orientation: args.orientation,
                widthCm: args.width_cm,
                heightCm: args.height_cm
              })
              changes.push(result)
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `📐 页面已设为 ${extra.to.paper ?? '自定义'} / ${extra.to.orientation === 'landscape' ? '横向' : '纵向'}：${extra.to.width_cm} × ${extra.to.height_cm} cm`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '设置纸张与方向', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_set_docx_margins',
      description:
        '设置 Word 文档末节的页边距（厘米）。已有 `w:pgMar` 时逐项改写，只覆盖传入的项；没有时新建并**写齐全部七项**（上下 2.54 / 左右 3.17 / 页眉页脚 1.5 / 装订线 0 cm）——缺属性的 `pgMar` 会让不同阅读器各取各的默认值。左右合计超过页宽会被拒绝。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        top_cm: { type: 'number', description: '上边距（厘米）。' },
        bottom_cm: { type: 'number', description: '下边距（厘米）。' },
        left_cm: { type: 'number', description: '左边距（厘米）。' },
        right_cm: { type: 'number', description: '右边距（厘米）。' },
        header_cm: { type: 'number', description: '页眉距边界（厘米）。' },
        footer_cm: { type: 'number', description: '页脚距边界（厘米）。' },
        gutter_cm: { type: 'number', description: '装订线（厘米）。' }
      },
      output: envelopeOutput('设置页边距'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes }) => {
              const result = doc.setMargins({
                topCm: args.top_cm,
                bottomCm: args.bottom_cm,
                leftCm: args.left_cm,
                rightCm: args.right_cm,
                headerCm: args.header_cm,
                footerCm: args.footer_cm,
                gutterCm: args.gutter_cm
              })
              changes.push(result)
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `📏 页边距：上 ${extra.to.top_cm} / 下 ${extra.to.bottom_cm} / 左 ${extra.to.left_cm} / 右 ${extra.to.right_cm} cm（${extra.updated === 'created' ? '新建 pgMar' : '就地改写'}）`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '设置页边距', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_insert_docx_page_number',
      description:
        '在页脚（默认）或页眉插入页码域。页码是 `PAGE` 域（fldChar 三段式），由 Word 按页自动更新，插件不写死数字——读完请按 F9 或打印预览让 Word 刷新。复用默认页脚/页眉部件，没有就新建并接好关系与内容类型；可加前后缀（如「第 」/「 页」）并设置对齐。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        position: { type: 'string', description: 'footer（默认）或 header。' },
        align: { type: 'string', description: '对齐：left / center（默认）/ right。' },
        prefix: { type: 'string', description: '页码前文本，如「第 」。' },
        suffix: { type: 'string', description: '页码后文本，如「 页」。' }
      },
      output: envelopeOutput('插入页码域'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes, warnings }) => {
              const result = doc.insertPageNumber({
                position: args.position ?? 'footer',
                align: args.align ?? 'center',
                prefix: args.prefix ?? '',
                suffix: args.suffix ?? ''
              })
              changes.push(result)
              warnings.push({
                code: 'FIELD_NEEDS_UPDATE',
                message: '页码是 PAGE 域，数字由 Word 计算：打开后按 Ctrl+A 再按 F9（或打印预览）即可刷新。'
              })
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `🔢 已在${extra.position === 'header' ? '页眉' : '页脚'}插入 PAGE 页码域（${extra.align}${extra.created ? '，新建部件' : ''}）→ ${extra.part}`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '插入页码域', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_update_docx_table_cell',
      description:
        '改写 Word 文档表格中一个单元格的文本。只改目标单元格里第一个段落的 run 内容，保留单元格宽度、边框、底纹与段落属性；表格其余单元格与其它部件不动。行列号从 0 开始，只数直接子元素，因此嵌套表格不会错位。先用 office_read_docx(detail="tables") 看清表格结构再写。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        table: { type: 'integer', description: '第几个表格（0 基），默认 0。' },
        row: { type: 'integer', required: true, description: '行号（0 基）。' },
        column: { type: 'integer', required: true, description: '列号（0 基）。' },
        text: { type: 'string', required: true, description: '新文本；空字符串表示清空该单元格文本。' }
      },
      output: envelopeOutput('改写表格单元格'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes }) => {
              const result = doc.updateTableCell({
                table: args.table ?? 0,
                row: args.row,
                column: args.column,
                text: args.text
              })
              changes.push(result)
              return result
            },
            summaryOf: (_s, _c, extra) => `📊 已改写表格 ${extra.table} 第 ${extra.row} 行第 ${extra.column} 列：「${extra.from}」→「${extra.to}」。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '改写表格单元格', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_create_docx_table',
      description:
        '在 Word 文档的指定段落之后插入一张新表格。表格是 body 级元素，插到「第 N 段之后」就是把 `<w:tbl>` 插到那个段落后面，既有段落编号不受影响。会自动写 `<w:tblGrid>` 每列宽度与每个单元格的 `w:tcW`（不写 Word 会按内容猜列宽），首行可加粗并标记为重复表头；表格成为正文最后一个内容元素时会补一个空段落（Word 自己也这么做）。样式 ID 在文档里不存在时会补一个最小定义。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        after: { type: 'integer', required: true, description: '插入到第几段之后（0 基）；-1 表示插到正文最前。' },
        rows: {
          type: 'array',
          required: true,
          items: { type: 'array', items: { type: 'string' } },
          description: '二维文本数组，第一维是行（第一行通常是表头）。'
        },
        header: { type: 'boolean', description: '是否把首行当表头（加粗 + 重复表头标记），默认 true。' },
        style_id: { type: 'string', description: '表格样式 ID，默认 TableGrid（如 LightShading-Accent1）。' },
        column_widths: {
          type: 'array',
          items: { type: 'integer' },
          description: '每列宽度（dxa，1/20 磅）；省略时按正文可用宽度均分。'
        }
      },
      output: envelopeOutput('新建表格'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes, warnings }) => {
              const result = doc.createTable({
                after: args.after,
                rows: args.rows ?? [],
                header: args.header !== false,
                styleId: args.style_id,
                columnWidths: Array.isArray(args.column_widths) ? args.column_widths : null
              })
              changes.push(result)
              if (result.style_defined_in_document) {
                warnings.push({
                  code: 'STYLE_CREATED',
                  message: `文档里原本没有样式「${result.style_id}」，已补一个最小定义（名称 + 单线边框）。`
                })
              }
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `🧱 已插入表格（${extra.rows} 行 × ${extra.columns} 列${extra.header ? '，首行为表头' : ''}，样式 ${extra.style_id}）。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '新建表格', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_set_docx_table_style',
      description:
        '给 Word 表格套用样式并设置「表格样式选项」。两件事必须一起做：`w:tblStyle` 指向 styles.xml 里的表格样式，`w:tblLook` 决定首行/首列/镶边行等条件格式是否生效。指定的样式 ID 在文档里不存在时会**补一个最小定义**（名称 + 单线边框），否则 Word 会静默退回「普通表格」。常用的内置样式 ID：TableGrid、LightShading-Accent1、LightGrid-Accent1、MediumShading1-Accent1。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        table: { type: 'integer', description: '第几个表格（0 基），默认 0。' },
        style_id: { type: 'string', required: true, description: '表格样式 ID，如 TableGrid。' },
        first_row: { type: 'boolean', description: '强调首行，默认 true。' },
        last_row: { type: 'boolean', description: '强调末行，默认 false。' },
        first_column: { type: 'boolean', description: '强调首列，默认 false。' },
        last_column: { type: 'boolean', description: '强调末列，默认 false。' },
        banded_rows: { type: 'boolean', description: '镶边行（隔行底色），默认 true。' }
      },
      output: envelopeOutput('设置表格样式'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes, warnings }) => {
              const result = doc.setTableStyle({
                table: args.table ?? 0,
                styleId: args.style_id,
                firstRow: args.first_row !== false,
                lastRow: args.last_row === true,
                firstColumn: args.first_column === true,
                lastColumn: args.last_column === true,
                bandedRows: args.banded_rows !== false
              })
              changes.push(result)
              if (result.style_defined_in_document) {
                warnings.push({
                  code: 'STYLE_CREATED',
                  message: `文档里原本没有样式「${args.style_id}」，已补一个最小定义（名称 + 单线边框）。`
                })
              }
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `🎨 已给表格 ${extra.table} 套用样式「${extra.style_id}」${
                extra.style_defined_in_document ? '（文档里没有，已补最小定义）' : ''
              }。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `设置表格样式 ${args.style_id}`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_insert_docx_table_row',
      description:
        '在 Word 表格里插入一行。新行的格式从**参照行**复制（行属性与每个单元格的单元格属性），因此列宽、边框、底纹与行高都跟着走 —— 自己拼一个新行而不带单元格属性，Word 会按默认宽度重排整张表。`after=-1` 表示插到表格最前。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        table: { type: 'integer', description: '第几个表格（0 基），默认 0。' },
        after: { type: 'integer', required: true, description: '插到第几行之后（0 基）；-1 表示插到最前。' },
        cells: {
          type: 'array',
          items: { type: 'string' },
          description: '各单元格文本；缺省则插入空单元格。'
        }
      },
      output: envelopeOutput('插入表格行'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes }) => {
              const result = doc.insertTableRow({
                table: args.table ?? 0,
                after: args.after,
                cells: Array.isArray(args.cells) ? args.cells : []
              })
              changes.push(result)
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `➕ 已在表格 ${extra.table} 第 ${extra.row} 行位置插入新行（${extra.columns} 列，共 ${extra.row_count} 行）。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '插入表格行', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_delete_docx_table_row',
      description:
        '删除 Word 表格里的一行。与删除段落一样做**保护性检查**：行内含书签、批注锚点、修订标记或域代码时默认拒绝（这些会随行消失，Word 会静默丢弃对应功能），需显式传 allow_markup_loss=true。表格只剩一行时拒绝删除。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        table: { type: 'integer', description: '第几个表格（0 基），默认 0。' },
        row: { type: 'integer', required: true, description: '要删除的行号（0 基）。' },
        allow_markup_loss: { type: 'boolean', description: '明知会丢失该行的结构化标记仍继续，默认 false。' }
      },
      output: envelopeOutput('删除表格行'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes, warnings }) => {
              const result = doc.deleteTableRow({
                table: args.table ?? 0,
                row: args.row,
                allowMarkupLoss: args.allow_markup_loss === true
              })
              changes.push(result)
              if (result.markup_lost.length > 0) {
                warnings.push({ code: 'MARKUP_LOST', message: `该行含 ${result.markup_lost.join('、')}，已随行删除。` })
              }
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `➖ 已删除表格 ${extra.table} 第 ${extra.row} 行（内容「${extra.removed_text}」），表格现有 ${extra.row_count} 行。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '删除表格行', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_merge_docx_table_cells',
      description:
        '合并 Word 表格里的一个矩形单元格区域。OOXML 的合并分两步（横向写 gridSpan 并删掉被并入的单元格、纵向写 vMerge restart/继续），缺一步 Word 就显示成没合并。被合并掉的单元格文字会按 Word 的行为**并入左上角单元格**（用换行分隔）并清空原格，不丢内容。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        table: { type: 'integer', description: '第几个表格（0 基），默认 0。' },
        row: { type: 'integer', required: true, description: '区域左上角行（0 基）。' },
        column: { type: 'integer', required: true, description: '区域左上角列（0 基）。' },
        row_span: { type: 'integer', description: '纵向跨几行，默认 1。' },
        col_span: { type: 'integer', description: '横向跨几列，默认 1。' }
      },
      output: envelopeOutput('合并表格单元格'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes }) => {
              const result = doc.mergeTableCells({
                table: args.table ?? 0,
                row: args.row,
                column: args.column,
                rowSpan: args.row_span ?? 1,
                colSpan: args.col_span ?? 1
              })
              changes.push(result)
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `🔗 已合并表格 ${extra.table} 的 ${extra.row_span}×${extra.col_span} 区域（左上角 ${extra.row},${extra.column}）${
                extra.merged_text.length > 0 ? `，并入文字：${extra.merged_text.join('、')}` : ''
              }。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '合并表格单元格', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_set_paragraph_style',
      description:
        '给某个段落套用段落样式（`w:pPr/w:pStyle`）。只改样式引用，不动 run 内容与段落里的其它属性（缩进、间距、对齐、大纲级别都保留）。样式 ID 在文档里不存在时会补一个最小定义 —— 否则 Word 不报错但会**静默按默认样式显示**，看起来像「设置了没生效」。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        index: { type: 'integer', required: true, description: '段落下标（0 基，只计正文直接段落）。' },
        style_id: { type: 'string', required: true, description: '段落样式 ID，如 Heading2、Quote、Title。' }
      },
      output: envelopeOutput('设置段落样式'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes, warnings }) => {
              const result = doc.setParagraphStyle({ index: args.index, styleId: args.style_id })
              changes.push(result)
              if (result.style_defined_in_document) {
                warnings.push({
                  code: 'STYLE_CREATED',
                  message: `文档里原本没有段落样式「${args.style_id}」，已补一个最小定义（只有名称）。`
                })
              }
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `📐 已把第 ${extra.index} 段样式设为「${extra.style_id}」${extra.from ? `（原样式 ${extra.from}）` : ''}。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `设置段落样式 ${args.style_id}`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_set_character_style',
      description:
        '给某个段落里的所有 run 套用字符样式（`w:rPr/w:rStyle`）。只加字符样式引用，不改文字与段落样式。样式 ID 不存在时会补一个最小定义。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        index: { type: 'integer', required: true, description: '段落下标（0 基）。' },
        style_id: { type: 'string', required: true, description: '字符样式 ID，如 Strong、Emphasis。' }
      },
      output: envelopeOutput('设置字符样式'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes, warnings }) => {
              const result = doc.setCharacterStyle({ index: args.index, styleId: args.style_id })
              changes.push(result)
              if (result.style_defined_in_document) {
                warnings.push({
                  code: 'STYLE_CREATED',
                  message: `文档里原本没有字符样式「${args.style_id}」，已补一个最小定义（只有名称）。`
                })
              }
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `🔤 已给第 ${extra.index} 段的 ${extra.runs} 个 run 套用字符样式「${extra.style_id}」。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `设置字符样式 ${args.style_id}`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_insert_page_break',
      description:
        '在指定段落之后插入分页符。Word 的分页符就是 run 里的 `<w:br w:type="page"/>`，这里放在独立段落里最稳，不会把后续段落挤到同一行；段落编号会随之增加。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        after: { type: 'integer', required: true, description: '在第几段之后插入（0 基）；-1 表示插到正文最前。' }
      },
      output: envelopeOutput('插入分页符'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes }) => {
              const result = doc.insertPageBreak({ after: args.after })
              changes.push(result)
              return result
            },
            summaryOf: (_s, _c, extra) => `📄 已在第 ${extra.after} 段之后插入分页符（现为第 ${extra.index} 段）。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '插入分页符', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_resize_docx_image',
      description:
        '调整 Word 文档里某张图片的显示尺寸（按 96 DPI 换算）。Word 的布局尺寸取自 `<wp:extent>`，图形自身还带一份 `<a:xfrm><a:ext>` —— 两处都会改（只改一处，Word 下次编辑时会把另一处当真实尺寸，图片「弹回」原大小）。只给一个方向时按原比例缩放，不会把图拉变形。先用 office_read_docx 或图片清单确认下标（按图片在正文中出现的顺序，0 基）。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        index: { type: 'integer', required: true, description: '第几张图片（0 基，按正文出现顺序）。' },
        width_px: { type: 'integer', description: '目标宽度（像素）；与 height_px 至少给一个。' },
        height_px: { type: 'integer', description: '目标高度（像素）。' }
      },
      output: envelopeOutput('调整图片尺寸'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes }) => {
              const result = doc.resizeImage({
                index: args.index,
                widthPx: args.width_px ?? null,
                heightPx: args.height_px ?? null
              })
              changes.push(result)
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `🖼️ 已把第 ${extra.index} 张图片从 ${extra.from.width}×${extra.from.height}px 调整为 ${extra.to.width}×${extra.to.height}px。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '调整图片尺寸', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_delete_docx_image',
      description:
        '删除 Word 文档里的一张图片。图片所在段落如果只剩这一张图（没有其它文字），整段一起删掉，避免留下空行；此时若段内含书签、批注锚点、修订或域代码则默认拒绝，需传 allow_markup_loss=true。媒体部件与关系**保留**在包里（同一媒体可能被别处引用），不做垃圾回收。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        index: { type: 'integer', required: true, description: '第几张图片（0 基，按正文出现顺序）。' },
        allow_markup_loss: { type: 'boolean', description: '明知删整段会丢失结构化标记仍继续，默认 false。' }
      },
      output: envelopeOutput('删除图片'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes, warnings }) => {
              const result = doc.deleteImage({ index: args.index, allowMarkupLoss: args.allow_markup_loss === true })
              changes.push(result)
              if (result.removed_paragraph) {
                warnings.push({ code: 'PARAGRAPH_REMOVED', message: '图片所在段落只含这张图，已连同段落一起删除。' })
              }
              warnings.push({ code: 'MEDIA_KEPT', message: result.note })
              return result
            },
            summaryOf: (_s, _c, extra) =>
              `🗑️ 已删除第 ${extra.index} 张图片${extra.removed_paragraph ? '（连同空段落）' : ''}。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '删除图片', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_set_docx_image_wrap',
      description:
        '设置 Word 图片的环绕方式（行内 ↔ 浮动）。行内图是 `<wp:inline>`、浮动图是 `<wp:anchor>`——在 Word 里是两种不同的容器：浮动图带位置元素与环绕元素，且**必需属性一个都不能少**（少写一个 Word 报「内容有问题」），元素顺序也受 schema 约束（环绕元素必须排在 `wp:docPr` 之前）。支持 inline / square（四周型）/ topAndBottom（上下型）/ none / behind（衬于文字下方）/ inFront（浮于文字上方）；tight（紧密）与 through（穿越）需要多边形包围盒，明确拒绝而不是伪造矩形。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        index: { type: 'integer', required: true, description: '第几张图片（0 基，按正文出现顺序）。' },
        wrap: {
          type: 'string',
          required: true,
          description: 'inline | square | topAndBottom | none | behind | inFront。'
        },
        offset_x_emu: { type: 'number', description: '水平偏移（EMU，1 pt = 12700）；默认 0。' },
        offset_y_emu: { type: 'number', description: '垂直偏移（EMU）；默认 0。' },
        relative_from_h: { type: 'string', description: '水平参照：column（默认）/ page / margin / character。' },
        relative_from_v: { type: 'string', description: '垂直参照：paragraph（默认）/ page / margin / line。' }
      },
      output: envelopeOutput('设置图片环绕'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes }) => {
              const result = doc.setImageWrap({
                index: args.index,
                wrap: args.wrap,
                offsetXEmu: args.offset_x_emu ?? 0,
                offsetYEmu: args.offset_y_emu ?? 0,
                relativeFromH: args.relative_from_h ?? 'column',
                relativeFromV: args.relative_from_v ?? 'paragraph'
              })
              changes.push(result)
              return result
            },
            summaryOf: (_s, _c, extra) =>
              extra.changed
                ? `🧩 第 ${extra.index} 张图片环绕：${extra.from} → **${extra.to}**${extra.offset_emu && (extra.offset_emu.x || extra.offset_emu.y) ? `（偏移 ${extra.offset_emu.x},${extra.offset_emu.y} EMU）` : ''}`
                : `🧩 第 ${extra.index} 张图片本来就是行内图，未改动。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `图片环绕 → ${args.wrap}`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_insert_docx_image',
      description:
        '在 Word 文档中插入一张图片（新起一个只含图片的段落）。图片类型按**真实字节**判定而非扩展名；未指定尺寸时用图片原始像素尺寸（96 DPI），只给宽度时按比例缩放。适配器会一并补齐媒体部件、内容类型声明与关系，插入后可用 office_validate_docx 确认关系完整。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        after: { type: 'integer', required: true, description: '插入到第几段之后（0 基）；传 -1 表示插到文档最前面。' },
        image_path: { type: 'string', description: '图片文件路径（相对会话工作区或绝对路径）。与 image_base64 二选一。' },
        image_base64: { type: 'string', description: '图片的 base64 内容。与 image_path 二选一。' },
        width_px: { type: 'number', description: '显示宽度（像素）。' },
        height_px: { type: 'number', description: '显示高度（像素）。' },
        alt_text: { type: 'string', description: '替代文本（无文字说明文档的可访问性）。' }
      },
      output: envelopeOutput('插入图片'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          if (!args.image_path && !args.image_base64) {
            throw new OfficeError('INVALID_REQUEST', '必须提供 image_path 或 image_base64 之一。')
          }
          let data
          let sourceLabel
          if (args.image_path) {
            const imageAbsolute = resolvePath(args.image_path, true)
            const info = await stat(imageAbsolute)
            if (info.size > 64 * 1024 * 1024) {
              throw new OfficeError('MEMORY_LIMIT', `图片 ${info.size} 字节超过 64MB 上限。`)
            }
            const { readFile } = await import('node:fs/promises')
            data = await readFile(imageAbsolute)
            sourceLabel = basename(imageAbsolute)
          } else {
            data = Buffer.from(String(args.image_base64).replace(/^data:[^,]+,/, ''), 'base64')
            sourceLabel = 'base64'
            if (data.length === 0) throw new OfficeError('INVALID_REQUEST', 'image_base64 解码后为空。')
          }
          const extension = detectImageType(data)

          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes }) => {
              const result = doc.insertImage({
                after: args.after,
                data,
                extension,
                widthPx: args.width_px,
                heightPx: args.height_px,
                altText: args.alt_text
              })
              changes.push({ type: 'insert_image', media_part: result.media_part, bytes: result.bytes, index: result.index })
              return { ...result, source: sourceLabel }
            },
            summaryOf: (_s, _c, extra) =>
              `🖼️ 已插入图片 ${extra.source}（${extra.bytes} 字节，${extra.display_size_px.width}×${extra.display_size_px.height} px）→ ${extra.media_part}，关系 ${extra.relationship_id}`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '插入图片', kind: 'write', rawInput: args })
    })
  )

  /**
   * 带锁与事务地执行一次 DOCX 修改。
   *
   * 与 XLSX 走同一条事务流程：暂存 → 重新打开校验 → 原子替换 → 失败即回滚，
   * 因此「修改失败不破坏原文件」这条保证对 DOCX 同样成立。
   *
   * @param {object} args - 参数。
   * @returns {Promise<object>} 响应信封。
   */
  async function mutateDocx({ path, requestId, mutate, summaryOf }) {
    const absolute = resolvePath(path, true)
    const lock = await FileLock.acquire(absolute)
    const started = Date.now()
    try {
      const { buffer } = await readDocument(absolute, { maxBytes })
      const doc = DocxDocument.open(buffer)
      const warnings = []
      if (doc.info.hasMacro) warnings.push({ code: 'MACRO_DETECTED', message: '文档含 VBA 宏。插件不会执行宏，宏内容按原字节保留。' })
      if (doc.info.hasSignatures) warnings.push({ code: 'DIGITAL_SIGNATURE_PRESENT', message: '文档含数字签名，任何修改都会使签名失效。' })
      if (doc.info.embedded.length > 0) {
        warnings.push({ code: 'EMBEDDED_OBJECT_DETECTED', message: `文档含 ${doc.info.embedded.length} 个嵌入对象，已原样保留但不解析。` })
      }
      const structure = doc.structure()
      if (structure.has_revisions) {
        warnings.push({ code: 'REVISIONS_PRESENT', message: '文档含未接受的修订；段落下标按当前可见文本计算。' })
      }
      if (structure.has_comments) {
        warnings.push({ code: 'COMMENTS_PRESENT', message: `文档含 ${structure.comments} 条批注，插件不修改批注，但删改段落可能使其锚点失效。` })
      }
      if (structure.has_toc) {
        warnings.push({ code: 'TOC_PRESENT', message: '文档含目录域，插件不会更新目录；改动段落后请在 Word 中更新域。' })
      }

      const changes = []
      const extra = await mutate({ doc, changes, warnings })

      const temp = await (await store()).temp('docx')
      const transaction = await Transaction.begin(absolute, temp.path)
      for (const w of warnings) transaction.warn(w)
      try {
        const result = await transaction.commit(doc.save(), {
          overwrite: true,
          verify: (out) => {
            const reopened = DocxDocument.open(out)
            reopened.structure() // 必须能被重新解析，否则视为输出损坏
            if (reopened.bodyParagraphs().length === 0) {
              throw new OfficeError('OUTPUT_VALIDATION_FAILED', '输出文档没有任何正文段落。')
            }
          }
        })
        const envelope = ok({
          requestId,
          documentId: sha256(buffer),
          outputFile: {
            name: basename(absolute),
            path: absolute,
            size: result.size,
            mime_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            sha256: result.sha256
          },
          warnings: result.warnings,
          changes,
          performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
          data: { ...extra }
        })
        return { ...envelope, summary: summaryOf(result, changes, extra) }
      } finally {
        await transaction.rollback()
        await temp.dispose()
      }
    } finally {
      await lock.release()
    }
  }

  tools.push(
    defineTool({
      name: 'office_insert_paragraph',
      description:
        '在 Word 文档中插入一个新段落。采用最小修改：只新增一个 <w:p> 节点，样式定义、编号、页眉页脚、图片关系、批注与域代码全部原样保留。段落下标只计正文直接段落，不含表格单元格内的段落。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        after: { type: 'integer', required: true, description: '插入到第几段之后（0 基）；传 -1 表示插到文档最前面。' },
        text: { type: 'string', required: true, description: '段落文本；\\n 会转成软换行。' },
        style: { type: 'string', description: '段落样式 ID，如 Heading1、Heading2、Normal。须是文档 styles.xml 中已存在的样式。' },
        alignment: { type: 'string', enum: ['left', 'center', 'right', 'both'], description: '水平对齐。' }
      },
      output: envelopeOutput('插入段落'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes }) => {
              const result = doc.insertParagraph({ after: args.after, text: args.text, style: args.style, alignment: args.alignment })
              changes.push(result)
              return { index: result.index, paragraph_count: doc.bodyParagraphs().length }
            },
            summaryOf: (_s, _c, extra) => `📝 已在第 ${extra.index} 段位置插入新段落，文档现有 ${extra.paragraph_count} 段。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '插入段落', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_update_paragraph',
      description:
        '改写 Word 文档中一个段落的文本，保留该段的段落属性（样式、大纲级别、对齐、编号）。这是填充模板段落的首选工具——它绝不改变段落样式，因此不会破坏标题层级或编号。**保护性检查**：若该段含书签、批注锚点、修订标记、域代码或超链接，改写会让它们失效（Word 会静默丢弃对应功能），此时会拒绝并说明原因，需显式传 allow_markup_loss=true 才继续。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        index: { type: 'integer', required: true, description: '段落下标（0 基，只计正文直接段落）。' },
        text: { type: 'string', required: true, description: '新文本；空字符串表示清空该段内容而保留段落本身。' },
        allow_markup_loss: { type: 'boolean', description: '明知会丢失该段的结构化标记（书签/批注锚点/修订/域/超链接）仍继续，默认 false。' }
      },
      output: envelopeOutput('改写段落'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes }) => {
              const result = doc.updateParagraph({ index: args.index, text: args.text, allowMarkupLoss: args.allow_markup_loss === true })
              changes.push(result)
              return { index: result.index, from: result.from, to: result.to }
            },
            summaryOf: (_s, _c, extra) => `✏️ 已改写第 ${extra.index} 段：「${extra.from}」→「${extra.to}」。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `改写第 ${args.index} 段`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_delete_paragraph',
      description:
        '删除 Word 文档中的一个段落。破坏性操作：必须显式传 confirm=true。**保护性检查**：若该段含书签、批注锚点、修订标记、域代码或超链接，删除会让它们失效，此时会拒绝，需再传 allow_markup_loss=true。文档至少保留一个段落。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        index: { type: 'integer', required: true, description: '段落下标（0 基）。' },
        confirm: { type: 'boolean', required: true, description: '必须为 true 才执行删除（人工确认位）。' },
        allow_markup_loss: { type: 'boolean', description: '明知会丢失该段的结构化标记仍继续，默认 false。' }
      },
      output: envelopeOutput('删除段落'),
      async execute(args) {
        const requestId = newRequestId()
        if (args.confirm !== true) {
          return failure(requestId, new OfficeError('PERMISSION_DENIED', '删除段落是破坏性操作，需要显式传 confirm=true。', { needsConfirmation: true }))
        }
        try {
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes }) => {
              const result = doc.deleteParagraph({ index: args.index, allowMarkupLoss: args.allow_markup_loss === true })
              changes.push(result)
              return { index: result.index, removed: result.removed_preview, paragraph_count: doc.bodyParagraphs().length }
            },
            summaryOf: (_s, _c, extra) => `🗑️ 已删除第 ${extra.index} 段「${extra.removed}」，文档现有 ${extra.paragraph_count} 段。`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `删除第 ${args.index} 段`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_find_and_replace_docx',
      description:
        '在 Word 文档正文中查找替换文本。要求命中文本位于同一个格式片段（run）内；若某段合并后的文本含目标却被格式拆散在多个 run 中，会返回 TEXT_SPLIT_ACROSS_RUNS 警告而不是静默跳过。默认只改正文，可选同时改页眉页脚。建议先用 dry_run 预览命中数量。',
      parameters: {
        path: { type: 'string', required: true, description: 'Word 文档路径。' },
        find: { type: 'string', required: true, description: '要查找的文本。' },
        replace: { type: 'string', required: true, description: '替换为的文本。' },
        match_case: { type: 'boolean', description: '是否区分大小写，默认 false。' },
        include_headers: { type: 'boolean', description: '是否同时替换页眉与页脚，默认 false。' },
        dry_run: { type: 'boolean', description: '仅预览命中数量，不写入。' }
      },
      output: envelopeOutput('Word 查找替换'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          if (args.dry_run === true) {
            const absolute = resolvePath(args.path, true)
            const { buffer } = await readDocument(absolute, { maxBytes })
            const doc = DocxDocument.open(buffer)
            const result = doc.findAndReplace({
              find: args.find,
              replace: args.replace,
              matchCase: args.match_case === true,
              includeHeaders: args.include_headers === true
            })
            return {
              ...ok({
                requestId,
                performance: null,
                data: { dry_run: true, replacements: result.replacements, split_run_warnings: result.split_run_warnings }
              }),
              summary: `🔍 预览：命中 ${result.replacements} 处${result.split_run_warnings.length ? `，另有 ${result.split_run_warnings.length} 处因格式拆散无法替换` : ''}，未改动文件。`
            }
          }
          return await mutateDocx({
            path: args.path,
            requestId,
            mutate: async ({ doc, changes, warnings }) => {
              const result = doc.findAndReplace({
                find: args.find,
                replace: args.replace,
                matchCase: args.match_case === true,
                includeHeaders: args.include_headers === true
              })
              changes.push(...result.changes)
              warnings.push(...result.split_run_warnings)
              return { replacements: result.replacements, split_run_warnings: result.split_run_warnings }
            },
            summaryOf: (_s, _c, extra) =>
              `🔁 已替换 ${extra.replacements} 处「${args.find}」→「${args.replace}」${extra.split_run_warnings.length ? `｜⚠️ ${extra.split_run_warnings.length} 处因格式拆散未替换` : ''}`
          })
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `Word 查找替换：${args.find}`, kind: 'write', rawInput: args })
    })
  )

  // ───────────────────────── 创建与校验 ─────────────────────────

  tools.push(
    defineTool({
      name: 'office_create_document',
      description: '新建一个空的办公文档。format=xlsx 创建 Excel 工作簿，format=docx 创建 Word 文档（含标题段、示例表格、页眉与页码页脚）。若目标路径已存在则拒绝，避免意外覆盖。',
      parameters: {
        path: { type: 'string', required: true, description: '要创建的文件路径。' },
        format: { type: 'string', enum: ['xlsx', 'docx'], description: '文档格式，默认按扩展名推断。' },
        sheet_name: { type: 'string', description: 'xlsx：首张工作表名，默认 Sheet1。' },
        title: { type: 'string', description: 'docx：文档标题，默认「示例文档」。' },
        overwrite: { type: 'boolean', description: '是否允许覆盖已存在的文件，默认 false。' }
      },
      output: envelopeOutput('新建文档'),
      async execute(args) {
        const requestId = newRequestId()
        const started = Date.now()
        try {
          const absolute = resolveSafePath(args.path, { workspaceRoot, mustExist: false })
          if (existsSync(absolute) && args.overwrite !== true) {
            throw new OfficeError('PERMISSION_DENIED', `文件已存在：${args.path}。如需覆盖请显式传 overwrite=true。`, { needsConfirmation: true })
          }
          const inferred = /\.docx$/i.test(args.path) ? 'docx' : 'xlsx'
          const format = args.format ?? inferred
          const sheetName = args.sheet_name ?? 'Sheet1'
          const bytes = format === 'docx' ? buildBasicDocx({ title: args.title ?? '示例文档' }) : buildEmptyWorkbook(sheetName)
          const mime =
            format === 'docx'
              ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
              : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          const temp = await (await store()).temp('create')
          const transaction = await Transaction.begin(absolute, temp.path)
          try {
            const result = await transaction.commit(bytes, {
              overwrite: args.overwrite === true,
              verify: (out) => {
                if (format === 'docx') {
                  DocxDocument.open(out).structure()
                } else {
                  const wb = Workbook.open(out)
                  if (!wb.sheetNames().includes(sheetName)) {
                    throw new OfficeError('OUTPUT_VALIDATION_FAILED', '新建的工作簿重新打开后缺少预期工作表。')
                  }
                }
              }
            })
            return {
              ...ok({
                requestId,
                outputFile: { name: basename(absolute), path: absolute, size: result.size, mime_type: mime, sha256: result.sha256 },
                performance: { duration_ms: Date.now() - started, memory_bytes: process.memoryUsage().heapUsed },
                data: { format, ...(format === 'docx' ? { title: args.title ?? '示例文档' } : { sheets: [sheetName] }) }
              }),
              summary: format === 'docx' ? `🆕 已创建 Word 文档 ${basename(absolute)}（${result.size} 字节）。` : `🆕 已创建 ${basename(absolute)}（工作表：${sheetName}，${result.size} 字节）。`
            }
          } finally {
            await transaction.rollback()
            await temp.dispose()
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `新建 ${args.path}`, kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_validate_workbook',
      description:
        '对工作簿做结构与内容校验：ZIP 包完整性、关系有效性、能否重新打开、工作表数量、公式是否保留、图表与图片关系、宏是否被破坏、工作表保护状态。修改文件后用它做自检。',
      parameters: {
        path: { type: 'string', required: true, description: '工作簿路径。' },
        baseline_path: { type: 'string', description: '可选：改动前的工作簿，用于对比部件级差异。' }
      },
      output: envelopeOutput('工作簿校验'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const absolute = resolvePath(args.path, true)
          const { buffer } = await readDocument(absolute, { maxBytes })
          const checks = []
          const warnings = []

          let pkg
          try {
            pkg = ZipPackage.open(buffer)
            checks.push({ name: 'ZIP 包可解析', ok: true })
          } catch (err) {
            return {
              ...ok({ requestId, data: { checks: [{ name: 'ZIP 包可解析', ok: false, detail: err.message }], valid: false } }),
              summary: `❌ 校验失败：ZIP 包无法解析（${err.message}）`
            }
          }
          checks.push({ name: 'ZIP 包可解析', ok: true, detail: `${pkg.names().length} 个部件` })

          const required = ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels']
          const missing = required.filter((p) => !pkg.has(p))
          checks.push({ name: '必需部件齐全', ok: missing.length === 0, detail: missing.length ? `缺少 ${missing.join(', ')}` : '全部存在' })

          // XML 部件「格式良好」检查：本插件自己的解析器是宽松的（容忍多根、未声明前缀），
          // 所以「我们能读」不等于「Excel 能读」。这里按标签深度扫一遍：
          // 标签要配平、只能有一个根元素、根元素闭合后不能再有顶层内容。
          const malformed = []
          for (const name of pkg.names()) {
            if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue
            let text = null
            try {
              text = pkg.readText(name)
            } catch {
              continue
            }
            const problems = []
            if ((text.match(/<\?xml/g) ?? []).length > 1) problems.push('多条 XML 声明')
            const body = text.replace(/<\?xml[^>]*\?>/g, '').replace(/<!--[\s\S]*?-->/g, '')
            let depth = 0
            let rootClosed = false
            let extraRoots = 0
            let sawRoot = false
            for (const match of body.matchAll(/<[^>]+>/g)) {
              const tag = match[0]
              if (tag.startsWith('<?') || tag.startsWith('<!')) continue
              if (tag.startsWith('</')) {
                depth -= 1
                if (depth === 0) rootClosed = true
                if (depth < 0) {
                  problems.push('闭合标签多于开始标签')
                  depth = 0
                }
                continue
              }
              const selfClosing = tag.endsWith('/>')
              if (depth === 0) {
                if (rootClosed) extraRoots += 1
                else sawRoot = true
              }
              if (!selfClosing) depth += 1
            }
            if (!sawRoot) problems.push('没有根元素')
            if (extraRoots > 0) problems.push(`根元素之后还有 ${extraRoots} 个顶层元素`)
            if (depth !== 0) problems.push('标签未配平')
            if (problems.length > 0) malformed.push({ part: name, problems })
          }
          checks.push({
            name: 'XML 部件格式良好（标签配平、单一根元素）',
            ok: malformed.length === 0,
            detail:
              malformed.length === 0
                ? `${pkg.names().length} 个部件通过`
                : malformed.map((m) => `${m.part}：${m.problems.join('、')}`).join('；')
          })

          let wb = null
          try {
            wb = Workbook.open(buffer)
            checks.push({ name: '工作簿可重新打开', ok: true, detail: `${wb.sheetNames().length} 张工作表` })
          } catch (err) {
            checks.push({ name: '工作簿可重新打开', ok: false, detail: err.message })
          }

          if (wb) {
            for (const sheet of wb.info.sheets) {
              if (!pkg.has(sheet.path)) {
                checks.push({ name: `工作表「${sheet.name}」部件存在`, ok: false, detail: `缺少 ${sheet.path}` })
                continue
              }
              try {
                XmlDoc.parse(pkg.readText(sheet.path))
                checks.push({ name: `工作表「${sheet.name}」XML 合法`, ok: true })
              } catch (err) {
                checks.push({ name: `工作表「${sheet.name}」XML 合法`, ok: false, detail: err.message })
              }
            }
            // 关系有效性：每个工作表 XML 里引用的 rId 必须有对应关系
            const rels = new Map()
            if (pkg.has('xl/_rels/workbook.xml.rels')) {
              for (const rel of findAll(XmlDoc.parse(pkg.readText('xl/_rels/workbook.xml.rels')).root, 'Relationship')) {
                rels.set(attr(rel, 'Id'), attr(rel, 'Target'))
              }
            }
            const brokenRels = wb.info.sheets.filter((s) => !s.path || !rels.size || ![...rels.values()].some((t) => s.path.endsWith(t.replace(/^\/?xl\//, '').replace(/^\.\//, ''))))
            checks.push({
              name: '工作表关系有效',
              ok: brokenRels.length === 0,
              detail: brokenRels.length ? `可疑：${brokenRels.map((s) => s.name).join(', ')}` : '全部有效'
            })

            const formulaCount = countFormulas(pkg, wb)
            checks.push({ name: '公式保留', ok: true, detail: `检出 ${formulaCount} 个公式单元格` })
            if (formulaCount > 0) {
              warnings.push({ code: 'FORMULA_NOT_RECALCULATED', message: '工作簿含公式，插件不重算公式；缓存结果可能已过期，请在 Excel/WPS 中打开以重算。' })
            }
            if (wb.info.hasCharts) {
              const chartParts = pkg.names().filter((n) => n.startsWith('xl/charts/'))
              checks.push({ name: '图表关系完整', ok: chartParts.length > 0, detail: `${chartParts.length} 个图表部件` })
            }
            if (wb.info.hasMacro) {
              checks.push({ name: '宏部件保留', ok: pkg.has('xl/vbaProject.bin'), detail: 'vbaProject.bin 存在' })
            }
            if (wb.info.hasExternalLinks) {
              const links = pkg.names().filter((n) => n.startsWith('xl/externalLinks/'))
              checks.push({ name: '外部链接保留', ok: links.length > 0, detail: `${links.length} 个外部链接部件，未被更新` })
            }
            checks.push({ name: '工作表保护状态', ok: true, detail: wb.info.sheets.map((s) => `${s.name}:${s.state}`).join(', ') })
          }

          if (args.baseline_path) {
            const baseAbsolute = resolvePath(args.baseline_path, true)
            const { buffer: baseBuffer } = await readDocument(baseAbsolute, { maxBytes })
            const basePkg = ZipPackage.open(baseBuffer)
            const changed = []
            for (const name of basePkg.names()) {
              if (!pkg.has(name)) { changed.push({ part: name, change: 'removed' }); continue }
              if (!pkg.read(name).equals(basePkg.read(name))) changed.push({ part: name, change: 'modified' })
            }
            for (const name of pkg.names()) {
              if (!basePkg.has(name)) changed.push({ part: name, change: 'added' })
            }
            checks.push({ name: '与基线的部件差异', ok: true, detail: changed.length === 0 ? '无差异' : `${changed.length} 个部件变化` })
            return {
              ...ok({ requestId, warnings, data: { valid: checks.every((c) => c.ok), checks, changed_parts: changed } }),
              summary: `🔎 校验完成：${checks.filter((c) => c.ok).length}/${checks.length} 项通过${changed.length ? `，与基线相比 ${changed.length} 个部件变化（${changed.map((c) => c.part).join(', ')}）` : ''}`
            }
          }

          const valid = checks.every((c) => c.ok)
          return {
            ...ok({ requestId, warnings, data: { valid, checks } }),
            summary: `${valid ? '✅' : '❌'} 校验${valid ? '通过' : '失败'}：${checks.filter((c) => c.ok).length}/${checks.length} 项通过${warnings.length ? `｜${warnings.length} 条提示` : ''}`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '校验工作簿', kind: 'read', rawInput: args.path })
    })
  )

  // ───────────────────────── 任务控制 ─────────────────────────

  tools.push(
    defineTool({
      name: 'office_preview_operation',
      description:
        '预览一个批量修改计划：计算影响范围、命中数量与风险（是否覆盖原文件、是否破坏性、是否涉及宏或签名文件），返回 plan_id。确认无误后用 office_execute_operation 执行。所有破坏性或大范围操作都应先预览。',
      parameters: {
        path: { type: 'string', required: true, description: '目标文件路径。' },
        operation: {
          type: 'string',
          required: true,
          enum: ['write_cells', 'find_and_replace', 'delete_worksheet', 'insert_rows', 'delete_rows', 'insert_columns', 'delete_columns', 'merge_cells'],
          description: '要预览的操作类型。'
        },
        parameters: { type: 'json', required: true, description: '该操作的参数（与对应工具一致）。' }
      },
      output: envelopeOutput('操作预览'),
      async execute(args) {
        const requestId = newRequestId()
        try {
          const absolute = resolvePath(args.path, true)
          const { wb, type } = await openWorkbook(args.path)
          const warnings = securityWarnings(wb)
          const impacts = []
          let risky = false
          let destructive = false

          if (args.operation === 'write_cells') {
            const cells = args.parameters?.cells ?? []
            impacts.push(`将写入 ${cells.length} 个单元格到「${args.parameters?.sheet}」`)
            if (cells.length > 500) risky = true
          } else if (args.operation === 'find_and_replace') {
            const result = wb.findAndReplace({
              find: args.parameters?.find ?? '',
              replace: args.parameters?.replace ?? '',
              sheet: args.parameters?.sheet ?? null,
              matchCase: args.parameters?.match_case === true,
              wholeCell: args.parameters?.whole_cell === true,
              includeFormulas: args.parameters?.include_formulas === true
            })
            impacts.push(`命中 ${result.replacements} 处替换`)
            if (result.replacements > 100) risky = true
          } else if (args.operation === 'delete_worksheet') {
            destructive = true
            impacts.push(`将永久删除工作表「${args.parameters?.sheet}」及其全部内容`)
          } else if (args.operation.startsWith('delete_')) {
            destructive = true
            impacts.push(`将删除${args.operation.includes('rows') ? '行' : '列'}，并重排后续引用`)
          } else {
            impacts.push(`将执行 ${args.operation}`)
          }

          if (wb.info.hasMacro) impacts.push('文件含宏，修改后宏代码保留但可能影响宏的预期行为')
          if (wb.info.hasSignatures) impacts.push('文件含数字签名，任何修改都会使签名失效')
          if (type.ext === 'xlsm' || type.ext === 'docm' || type.ext === 'pptm') destructive = destructive || true

          const planId = `plan-${sha256(`${requestId}:${absolute}`).slice(0, 12)}`
          tasks.set(planId, {
            plan_id: planId,
            created_at: new Date().toISOString(),
            path: absolute,
            operation: args.operation,
            parameters: args.parameters,
            status: 'previewed',
            destructive,
            risky,
            warnings
          })

          return {
            ...ok({
              requestId,
              warnings,
              data: {
                plan_id: planId,
                operation: args.operation,
                file: absolute,
                impacts,
                destructive,
                risky,
                requires_confirmation: destructive || risky,
                warnings
              }
            }),
            summary: `🔍 预览 ${args.operation}：${impacts.join('；')}${destructive || risky ? '｜⚠️ 需要确认' : ''}｜plan_id=${planId}`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `预览 ${args.operation}`, kind: 'read', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_get_task_status',
      description: '查询一个预览/执行计划的状态。不传 plan_id 时列出本会话的全部计划。',
      parameters: {
        plan_id: { type: 'string', description: '计划标识，省略则列出全部。' }
      },
      output: envelopeOutput('任务状态'),
      async execute(args) {
        const requestId = newRequestId()
        if (!args.plan_id) {
          const all = [...tasks.values()]
          return {
            ...ok({ requestId, performance: null, data: { count: all.length, tasks: all } }),
            summary: `📋 共 ${all.length} 个计划：${all.map((t) => `${t.plan_id}(${t.status})`).join('、') || '无'}`
          }
        }
        const task = tasks.get(args.plan_id)
        if (!task) return failure(requestId, new OfficeError('FILE_NOT_FOUND', `未找到计划 ${args.plan_id}。`, { needsConfirmation: false }))
        return {
          ...ok({ requestId, performance: null, data: task }),
          summary: `📋 ${task.plan_id}：${task.status}｜操作 ${task.operation}｜文件 ${basename(task.path)}`
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '查询任务状态', kind: 'read', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_cancel_task',
      description: '取消一个尚未执行的预览计划。已经执行完成的任务无法取消，但可以用 office_rollback_operation 恢复。',
      parameters: {
        plan_id: { type: 'string', required: true, description: '计划标识。' }
      },
      output: envelopeOutput('取消任务'),
      async execute(args) {
        const requestId = newRequestId()
        const task = tasks.get(args.plan_id)
        if (!task) return failure(requestId, new OfficeError('FILE_NOT_FOUND', `未找到计划 ${args.plan_id}。`))
        if (task.status === 'executed') {
          return failure(requestId, new OfficeError('INVALID_REQUEST', '该计划已执行完成，无法取消；如需撤销请使用 office_rollback_operation。'))
        }
        task.status = 'cancelled'
        tasks.set(args.plan_id, task)
        return {
          ...ok({ requestId, performance: null, data: task }),
          summary: `🚫 已取消计划 ${args.plan_id}。`
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '取消任务', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_execute_operation',
      description:
        '执行一个已预览的计划。执行前会自动创建备份，执行后重新打开校验；校验失败则原文件保持不变。返回 backup_id，可用 office_rollback_operation 回滚。',
      parameters: {
        plan_id: { type: 'string', required: true, description: 'office_preview_operation 返回的计划标识。' },
        confirm: { type: 'boolean', description: '破坏性操作必须显式传 true 才会执行。' }
      },
      output: envelopeOutput('执行计划'),
      async execute(args) {
        const requestId = newRequestId()
        const task = tasks.get(args.plan_id)
        if (!task) return failure(requestId, new OfficeError('FILE_NOT_FOUND', `未找到计划 ${args.plan_id}。`))
        if (task.status === 'executed') return failure(requestId, new OfficeError('INVALID_REQUEST', '该计划已执行，不能重复执行。'))
        if (task.status === 'cancelled') return failure(requestId, new OfficeError('INVALID_REQUEST', '该计划已被取消。'))
        if ((task.destructive || task.risky) && args.confirm !== true) {
          return failure(requestId, new OfficeError('PERMISSION_DENIED', '该计划被判定为破坏性或大范围操作，需要显式传 confirm=true。', { needsConfirmation: true }))
        }
        try {
          const absolute = task.path
          const { buffer: backupBytes } = await readDocument(absolute, { maxBytes })
          const backupStore = await store()
          const backupId = `backup-${sha256(backupBytes).slice(0, 12)}`
          const backupPath = backupStore.outputPath(`${backupId}.bak`)
          if (!existsSync(backupPath)) {
            const { writeFile } = await import('node:fs/promises')
            await writeFile(backupPath, backupBytes)
          }
          task.backup_path = backupPath
          task.backup_id = backupId

          const p = task.parameters ?? {}
          const envelope = await mutateWorkbook({
            path: absolute,
            requestId,
            mutate: async ({ wb, changes }) => {
              if (task.operation === 'write_cells') {
                changes.push(...wb.writeCells(p.sheet, (p.cells ?? []).map((c) => ({ ...c, date1904: wb.info.date1904 }))))
              } else if (task.operation === 'find_and_replace') {
                const r = wb.findAndReplace({
                  find: p.find, replace: p.replace, sheet: p.sheet ?? null,
                  matchCase: p.match_case === true, wholeCell: p.whole_cell === true, includeFormulas: p.include_formulas === true
                })
                changes.push(...r.changes)
              } else if (task.operation === 'delete_worksheet') {
                wb.deleteWorksheet(p.sheet)
                changes.push({ type: 'delete_worksheet', sheet: p.sheet })
              } else if (task.operation === 'merge_cells') {
                wb.mergeCells(p.sheet, p.range)
                changes.push({ type: 'merge_cells', sheet: p.sheet, range: p.range })
              } else if (task.operation.startsWith('insert_') || task.operation.startsWith('delete_')) {
                const axis = task.operation.includes('rows') ? 'row' : 'column'
                const action = task.operation.startsWith('insert') ? 'insert' : 'delete'
                const r = wb.shiftRowsOrColumns(p.sheet, { axis, action, start: p.start, count: p.count ?? 1 })
                changes.push(...r.changes)
              } else {
                throw new OfficeError('UNSUPPORTED_FEATURE', `计划中的操作 ${task.operation} 暂不支持执行。`)
              }
            },
            summaryOf: (_s, changes) => `✅ 已执行 ${task.operation}，共 ${changes.length} 处变更。`
          })
          task.status = 'executed'
          task.executed_at = new Date().toISOString()
          tasks.set(args.plan_id, task)
          return {
            ...envelope,
            data: { ...envelope.data, plan_id: args.plan_id, backup_id: backupId },
            summary: `${envelope.summary}｜备份 ${backupId}`
          }
        } catch (err) {
          task.status = 'failed'
          task.error = String(err?.message ?? err)
          tasks.set(args.plan_id, task)
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '执行计划', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_rollback_operation',
      description: '把一个已执行计划回滚到执行前的状态：从备份文件恢复原内容。回滚前会校验备份完整性（哈希一致）。',
      parameters: {
        plan_id: { type: 'string', required: true, description: '要回滚的计划标识。' }
      },
      output: envelopeOutput('回滚'),
      async execute(args) {
        const requestId = newRequestId()
        const task = tasks.get(args.plan_id)
        if (!task) return failure(requestId, new OfficeError('FILE_NOT_FOUND', `未找到计划 ${args.plan_id}。`))
        if (!task.backup_path) return failure(requestId, new OfficeError('INVALID_REQUEST', '该计划没有可用备份，无法回滚。'))
        try {
          const { readFile, copyFile } = await import('node:fs/promises')
          const backupBytes = await readFile(task.backup_path)
          const expected = task.backup_id.replace(/^backup-/, '')
          if (sha256(backupBytes).slice(0, 12) !== expected) {
            throw new OfficeError('OUTPUT_VALIDATION_FAILED', '备份文件哈希不符，已中止回滚以避免写入损坏内容。')
          }
          const temp = await (await store()).temp('rollback')
          const transaction = await Transaction.begin(task.path, temp.path)
          try {
            await transaction.commit(backupBytes, {
              overwrite: true,
              verify: (out) => {
                Workbook.open(out)
              }
            })
          } finally {
            await transaction.rollback()
            await temp.dispose()
          }
          task.status = 'rolled_back'
          tasks.set(args.plan_id, task)
          return {
            ...ok({
              requestId,
              outputFile: { name: basename(task.path), path: task.path, size: backupBytes.length, sha256: sha256(backupBytes) },
              performance: null,
              data: { plan_id: args.plan_id, restored_from: task.backup_path }
            }),
            summary: `↩️ 已将 ${basename(task.path)} 回滚到执行前状态（备份 ${task.backup_id}）。`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: '回滚操作', kind: 'write', rawInput: args })
    })
  )

  tools.push(
    defineTool({
      name: 'office_delete_file',
      description: '删除一个办公文件。破坏性操作，必须显式传 confirm=true。删除前会先做类型校验，避免误删非目标文件。',
      parameters: {
        path: { type: 'string', required: true, description: '要删除的文件路径。' },
        confirm: { type: 'boolean', required: true, description: '必须为 true 才执行删除。' }
      },
      output: envelopeOutput('删除文件'),
      async execute(args) {
        const requestId = newRequestId()
        if (args.confirm !== true) {
          return failure(requestId, new OfficeError('PERMISSION_DENIED', '删除文件是破坏性操作，需要显式传 confirm=true。', { needsConfirmation: true }))
        }
        try {
          const absolute = resolvePath(args.path, true)
          await readDocument(absolute, { maxBytes })
          const { rm } = await import('node:fs/promises')
          await rm(absolute, { force: false })
          return {
            ...ok({ requestId, performance: null, data: { deleted: absolute } }),
            summary: `🗑️ 已删除 ${basename(absolute)}。`
          }
        } catch (err) {
          return failure(requestId, err)
        }
      },
      presentCall: (args) => ({ card: 'generic', title: `删除 ${args.path}`, kind: 'write', rawInput: args })
    })
  )

  return tools
}

/**
 * 把一张工作表渲染为 CSV 字节（UTF-8 带 BOM，Excel 打开中文不乱码）。
 *
 * 公式单元格导出其缓存值（与 Excel 的「另存为 CSV」行为一致），
 * 空单元格导出为空字段，行尾空列会被裁掉。
 *
 * @param {Workbook} wb - 工作簿。
 * @param {string} sheetName - 工作表名。
 * @returns {Buffer} CSV 字节。
 */
function workbookSheetToCsv(wb, sheetName) {
  const content = wb.readSheet(sheetName, { maxCells: 2_000_000 })
  const index = new Map(content.cells.map((c) => [`${c.col}:${c.row}`, c]))
  const lines = []
  for (let r = 1; r <= content.row_count; r += 1) {
    const fields = []
    for (let c = 1; c <= content.column_count; c += 1) {
      const cell = index.get(`${c}:${r}`)
      fields.push(cell ? formatCsvField(cell.value) : '')
    }
    while (fields.length > 0 && fields[fields.length - 1] === '') fields.pop()
    lines.push(fields.join(','))
  }
  while (lines.length > 0 && lines[lines.length - 1] === '') lines.pop()
  return Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(lines.join('\r\n'), 'utf8')])
}

/**
 * 格式化一个 CSV 字段，按 RFC 4180 加引号转义。
 * @param {unknown} value - 单元格值。
 * @returns {string} CSV 字段文本。
 */
function formatCsvField(value) {
  if (value === null || value === undefined) return ''
  let text
  if (typeof value === 'number') text = String(value)
  else if (typeof value === 'boolean') text = value ? 'TRUE' : 'FALSE'
  else text = String(value)
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text
}

/**
 * 统计工作簿中的公式单元格数量。
 * @param {ZipPackage} pkg - OPC 包。
 * @param {Workbook} wb - 工作簿。
 * @returns {number} 公式数量。
 */
function countFormulas(pkg, wb) {
  let count = 0
  for (const sheet of wb.info.sheets) {
    if (!pkg.has(sheet.path)) continue
    const doc = XmlDoc.parse(pkg.readText(sheet.path))
    count += findAll(doc.root, 'f').length
  }
  return count
}

/**
 * 构造一个最小合法的新工作簿。
 * @param {string} sheetName - 首张工作表名。
 * @returns {Buffer} .xlsx 字节。
 */
export function buildEmptyWorkbook(sheetName = 'Sheet1') {
  const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { maxEntries: 64, maxEntryBytes: 1 << 20, maxTotalBytes: 1 << 24, maxRatio: 200 })
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
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${escapeAttr(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`
  )
  pkg.write(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`
  )
  pkg.write(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData></sheetData></worksheet>`
  )
  pkg.write(
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`
  )
  pkg.write(
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>`
  )
  return pkg.toBuffer()
}

/**
 * 转义 XML 属性值。
 * @param {string} text - 文本。
 * @returns {string} 转义结果。
 */
function escapeAttr(text) {
  return String(text).replace(/[&<>"]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&quot;'))
}






