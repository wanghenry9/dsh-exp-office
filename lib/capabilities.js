/**
 * 能力清单：插件向 Harness 声明自身版本、协议版本与支持的能力。
 *
 * 对齐开发要求 §十：优先使用能力检测而非版本号判断；新增字段默认可选；
 * 不破坏已有工具名称与返回结构。
 *
 * @module dsh-exp-office/capabilities
 */

/** 插件版本，与 package.json / dsh.plugin.json 保持一致。 */
export const PLUGIN_VERSION = '0.1.0'

/** 插件标识。 */
export const PLUGIN_NAME = 'dsh-exp-office'

/** 支持的协议版本（docs §十 示例）。 */
export const PROTOCOL_VERSIONS = Object.freeze(['v1'])

/** 声明的 Harness 兼容范围。 */
export const SUPPORTED_HARNESS_VERSIONS = '>=0.1.0-rc.6'

/** 已实现的能力清单。未实现的能力不在此列出，避免过度承诺。 */
export const CAPABILITIES = Object.freeze([
  'file.detect_type',
  'file.inspect',
  'file.list',
  'file.create',
  'file.delete',
  'file.transactional_write',
  'file.backup_rollback',
  'file.lock',
  'xlsx.read',
  'xlsx.read_range',
  'xlsx.write',
  'xlsx.write_formula',
  'xlsx.worksheet.add',
  'xlsx.worksheet.delete',
  'xlsx.worksheet.rename',
  'xlsx.find_replace',
  'xlsx.merge_cells',
  'xlsx.table.create',
  'xlsx.chart.create',
  'xlsx.shift_rows',
  'xlsx.shift_columns',
  'xlsx.style.font',
  'xlsx.style.fill',
  'xlsx.style.number_format',
  'xlsx.style.alignment',
  'xlsx.export.copy',
  'xlsx.export.csv',
  'xlsx.validate',
  'docx.read',
  'docx.structure',
  'docx.write.paragraphs',
  'docx.style.paragraph',
  'docx.style.character',
  'docx.page_break',
  'docx.find_replace',
  'docx.headers.write',
  'docx.page_number_field',
  'docx.table_cells.write',
  'docx.tables.rows_write',
  'docx.tables.merge_cells',
  'docx.tables.style',
  'docx.tables.create',
  'docx.images.write',
  'docx.images.resize',
  'docx.images.wrap',
  'docx.images.delete',
  'docx.headers.create',
  'docx.page_layout',
  'docx.margins',
  'docx.page_number',
  'docx.bookmarks.write',
  'docx.validate',
  'docx.export',
  'pptx.create',
  'pptx.read',
  'pptx.structure',
  'pptx.write.text',
  'pptx.slide_management',
  'pptx.layouts.write',
  'pptx.theme.write',
  'pptx.images.write',
  'pptx.tables.write',
  'pptx.charts.write',
  'pptx.textbox.write',
  'pptx.validate',
  'pptx.export.pdf',
  'pdf.read',
  'pdf.structure',
  'pdf.text.extract',
  'pdf.text.search',
  'pdf.validate',
  'pdf.pages.rotate',
  'pdf.pages.delete',
  'pdf.pages.reorder',
  'pdf.overlay.watermark',
  'pdf.overlay.page_number',
  'pdf.annotate',
  'pdf.form.fill',
  'pdf.metadata.write',
  'pdf.rewrite',
  'pdf.split',
  'pdf.merge',
  'pdf.images.extract',
  'pdf.incremental_update',
  'convert.to_pdf',
  'convert.font_report',
  'task.preview',
  'task.execute',
  'task.rollback',
  'task.status',
  'task.cancel'
])

/** 当前阶段明确不支持的能力，用于向 Agent 给出可预期的边界（docs §二十三）。 */
export const NOT_IMPLEMENTED = Object.freeze([
  'xlsx.chart.advanced',
  'xlsx.pivot.modify',
  'xlsx.macro.execute',
  'xlsx.formula.recalculate',
  'pptx.validate.render',
  'pptx.animation.edit',
  'pdf.text.edit',
  'pdf.overlay.cjk',
  'pdf.create',
  'pdf.ocr',
  'pdf.image.rasterize',
  'pdf.to_office',
  'pdf.visual_regression',
  'office.local_automation'
])

/**
 * 构造启动能力清单（docs §十 的启动返回结构）。
 * @returns {object} 能力清单对象。
 */
export function toCapabilityManifest() {
  return {
    plugin_name: PLUGIN_NAME,
    plugin_version: PLUGIN_VERSION,
    protocol_versions: [...PROTOCOL_VERSIONS],
    supported_harness_versions: SUPPORTED_HARNESS_VERSIONS,
    capabilities: [...CAPABILITIES],
    not_implemented: [...NOT_IMPLEMENTED],
    runtime: {
      node: process.version,
      format: 'ooxml-minimal-edit',
      dependencies: 'none'
    }
  }
}
