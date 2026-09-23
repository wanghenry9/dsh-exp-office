/**
 * 统一错误码体系与结构化请求/响应封装。
 *
 * 依据开发要求：
 *   - §4  所有接口必须返回稳定的结构化结果，不允许只返回普通文本或底层异常堆栈。
 *   - §13 每个错误至少包含：错误码、人类可读描述、是否可重试、可能的解决方案、
 *         是否产生部分输出、是否需要回滚、是否需要人工确认。
 *   - §13 任何文件处理异常都不能直接吞掉，也不能只记录日志而向 Agent 返回成功。
 *
 * @module dsh-exp-office/errors
 */

/** 统一错误码定义表（docs §13）。code → 默认语义。 */
const ERROR_CODES = {
  INVALID_REQUEST: { retryable: false, solution: '检查工具参数是否符合 schema，修正后重试。', confirmation: false },
  UNSUPPORTED_FILE_TYPE: { retryable: false, solution: '仅支持 .xlsx/.docx/.pptx/.pdf 及其已声明的变体。', confirmation: false },
  CORRUPTED_DOCUMENT: { retryable: false, solution: '文件结构损坏，请提供可正常打开的原始文件。', confirmation: false },
  PASSWORD_REQUIRED: { retryable: false, solution: '文件已加密，需要先提供密码或解密后的副本。', confirmation: true },
  PERMISSION_DENIED: { retryable: false, solution: '当前权限策略不允许该操作，请调整授权范围。', confirmation: true },
  FILE_NOT_FOUND: { retryable: false, solution: '确认文件路径正确且文件存在。', confirmation: false },
  FILE_LOCKED: { retryable: true, solution: '文件正被其它任务占用，稍后重试。', confirmation: false },
  UNSUPPORTED_FEATURE: { retryable: false, solution: '该文档特性当前不支持，已保留原内容并给出警告。', confirmation: false },
  LAYOUT_CHANGED: { retryable: false, solution: '操作可能导致排版变化，请人工复核输出文件。', confirmation: true },
  FONT_NOT_FOUND: { retryable: false, solution: '目标字体缺失，可能发生字体替换，请人工复核。', confirmation: false },
  FORMULA_NOT_RECALCULATED: { retryable: true, solution: '公式未重算，需要本地 Office/WPS 联动或读取缓存值。', confirmation: false },
  OFFICE_NOT_INSTALLED: { retryable: false, solution: '未检测到本地 Microsoft Office，改用独立文件处理模式。', confirmation: false },
  WPS_NOT_INSTALLED: { retryable: false, solution: '未检测到本地 WPS Office，改用独立文件处理模式。', confirmation: false },
  AUTOMATION_DISABLED: {
    retryable: false,
    solution: '本地 Office/WPS 联动默认关闭：需要插件配置 allowLocalAutomation=true 才允许启动本机 Office 进程。',
    confirmation: true
  },
  AUTOMATION_FAILED: {
    retryable: true,
    solution: '本地引擎执行失败（已清理它启动的进程，原文件未被改动）。可查看审计日志，或改用独立文件处理模式。',
    confirmation: false
  },
  TIMEOUT: { retryable: true, solution: '任务超时，可缩小处理范围或增大 timeout_ms 后重试。', confirmation: false },
  MEMORY_LIMIT: { retryable: false, solution: '超出内存上限，请改用区域级/分块处理。', confirmation: false },
  OUTPUT_VALIDATION_FAILED: { retryable: false, solution: '输出文件未通过重新打开校验，原文件已保留。', confirmation: true },
  EXTERNAL_RESOURCE_BLOCKED: { retryable: false, solution: '外部链接/模板/图片访问被安全策略拦截。', confirmation: false },
  MACRO_DETECTED: { retryable: false, solution: '检测到宏，默认不执行宏；如需保留请显式确认。', confirmation: true },
  EMBEDDED_OBJECT_DETECTED: { retryable: false, solution: '检测到嵌入对象，已保留但不解析。', confirmation: false },
  INTERNAL_ERROR: { retryable: true, solution: '插件内部错误，请附带 request_id 上报。', confirmation: false }
};

/** 需要人工确认后才能执行的操作（docs §3.6）。 */
export const CONFIRM_REQUIRED_ACTIONS = Object.freeze([
  'overwrite_original',
  'delete_sheet',
  'delete_slide',
  'delete_paragraph',
  'delete_pdf_pages',
  'bulk_modify',
  'bulk_replace',
  'modify_signature',
  'modify_password',
  'modify_permission',
  'lossy_convert',
  'modify_macro_document'
]);

/**
 * 插件统一异常。抛出后由 {@link toErrorEnvelope} 转换为结构化错误响应，
 * 绝不允许把底层堆栈直接暴露给 Agent。
 */
export class OfficeError extends Error {
  /**
   * @param {string} code - 必须是 {@link ERROR_CODES} 中的键。
   * @param {string} message - 面向人的中文描述。
   * @param {object} [details] - 附加上下文（不得包含文档正文）。
   * @param {object} [options] - 覆盖默认语义。
   */
  constructor(code, message, details = {}, options = {}) {
    super(message);
    const spec = ERROR_CODES[code] ?? ERROR_CODES.INTERNAL_ERROR;
    this.name = 'OfficeError';
    this.code = ERROR_CODES[code] ? code : 'INTERNAL_ERROR';
    this.retryable = options.retryable ?? spec.retryable;
    this.solution = options.solution ?? spec.solution;
    this.needsConfirmation = options.needsConfirmation ?? spec.confirmation;
    this.partialOutput = options.partialOutput ?? false;
    this.needsRollback = options.needsRollback ?? false;
    this.details = details;
  }

  /** 转换为响应体中的 `error` 字段（docs §4.3）。 */
  toJSON() {
    return {
      code: this.code,
      message: this.message,
      details: this.details,
      retryable: this.retryable,
      solution: this.solution,
      partial_output: this.partialOutput,
      needs_rollback: this.needsRollback,
      needs_confirmation: this.needsConfirmation
    };
  }
}

/**
 * 把任意抛出物收敛为 {@link OfficeError}，保证不泄漏底层堆栈。
 * @param {unknown} err - 捕获到的异常。
 * @returns {OfficeError} 结构化异常。
 */
export function asOfficeError(err) {
  if (err instanceof OfficeError) return err;
  if (err && typeof err === 'object' && err.code === 'ENOENT') {
    return new OfficeError('FILE_NOT_FOUND', '文件不存在。', { syscall: err.syscall });
  }
  if (err && typeof err === 'object' && (err.code === 'EACCES' || err.code === 'EPERM')) {
    return new OfficeError('PERMISSION_DENIED', '操作系统拒绝了文件访问。', { syscall: err.syscall });
  }
  if (err && typeof err === 'object' && err.code === 'EBUSY') {
    return new OfficeError('FILE_LOCKED', '文件被其它进程占用。');
  }
  const message = err instanceof Error ? err.message : String(err);
  return new OfficeError('INTERNAL_ERROR', message);
}

/**
 * 成功响应构造器（docs §4.2）。
 * @param {object} args - 响应字段。
 * @returns {object} 结构化成功响应。
 */
export function ok({ requestId, documentId = null, outputFile = null, warnings = [], changes = [], performance = null, data = null }) {
  return {
    success: true,
    request_id: requestId,
    document_id: documentId,
    output_file: outputFile,
    warnings,
    changes,
    performance: performance ?? { duration_ms: 0, memory_bytes: 0 },
    data,
    error: null
  };
}

/**
 * 失败响应构造器（docs §4.3）。
 * @param {string} requestId - 关联标识。
 * @param {unknown} err - 原始异常，会被收敛。
 * @returns {object} 结构化失败响应。
 */
export function fail(requestId, err) {
  return {
    success: false,
    request_id: requestId,
    document_id: null,
    output_file: null,
    warnings: [],
    changes: [],
    performance: null,
    data: null,
    error: asOfficeError(err).toJSON()
  };
}

/**
 * 生成一次调用的关联标识。使用内置 crypto，不引入 uuid 依赖。
 * @returns {string} RFC4122 v4 形式的标识。
 */
export function newRequestId() {
  return globalThis.crypto.randomUUID();
}

/**
 * 按响应级别裁剪结果（docs §14.3：支持 summary / detail 两种响应级别）。
 * @param {object} envelope - 完整响应。
 * @param {'summary'|'detail'} level - 目标级别。
 * @returns {object} 裁剪后的响应。
 */
export function projectLevel(envelope, level) {
  if (level === 'detail') return envelope;
  if (!envelope.data || typeof envelope.data !== 'object') return envelope;
  const { rows, cells, paragraphs, slides, pages, ...rest } = envelope.data;
  return {
    ...envelope,
    data: {
      ...rest,
      ...(Array.isArray(rows) ? { row_count: rows.length } : {}),
      ...(Array.isArray(cells) ? { cell_count: cells.length } : {}),
      ...(Array.isArray(paragraphs) ? { paragraph_count: paragraphs.length } : {}),
      ...(Array.isArray(slides) ? { slide_count: slides.length } : {}),
      ...(Array.isArray(pages) ? { page_count: pages.length } : {})
    }
  };
}
