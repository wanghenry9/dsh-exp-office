/**
 * PPTX 适配器（阶段 4：读取与结构识别）。
 *
 * 与 XLSX / DOCX 共用同一套底座：`ooxml.js` 的 ZIP 容器与最小修改 XML 引擎，
 * 因此本文件只负责 PresentationML 语义。这已经是第三个格式适配器，
 * 前两个的经验在这里体现为：不重复容器逻辑、不猜测结构、对不支持的特性明确报告。
 *
 * 当前阶段**只读**。写入（文本/图片/表格/幻灯片增删）会在后续窗口补齐。
 *
 * 对齐开发要求 §七（PPTX 功能规划与开发注意事项）。
 *
 * @module dsh-exp-office/pptx
 */

import {
  ZipPackage,
  XmlDoc,
  findAll,
  find,
  attr,
  nodeText,
  escapeXmlText,
  escapeXmlAttr,
  ensureContentTypeOverride,
  removeContentTypeOverride,
  ensureContentTypeDefault,
  addRelationshipTo,
  removeRelationshipFrom,
  EMU_PER_PX,
  imageContentType,
  readImageSize,
  resolveImageExtent,
  DEFAULT_ZIP_LIMITS
} from './ooxml.js'
import { buildChartSpaceXml, buildLiteralSeriesXml, indexToCol } from './xlsx.js'
import { OfficeError } from './errors.js'

const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
/** 图表命名空间：`p:graphicFrame` 里引用图表部件时 `a:graphicData` 的 uri 就是它。 */
const NS_CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const CT_PRESENTATION = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
const CT_SLIDE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
const CT_SLIDE_MASTER = 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
const CT_SLIDE_LAYOUT = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
const CT_NOTES_SLIDE = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'
const CT_THEME = 'application/vnd.openxmlformats-officedocument.theme+xml'
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'

/** 占位符类型的中文名。 */
const PLACEHOLDER_LABELS = Object.freeze({
  title: '标题',
  ctrTitle: '居中标题',
  subTitle: '副标题',
  body: '正文',
  obj: '内容',
  dt: '日期',
  ftr: '页脚',
  sldNum: '幻灯片编号',
  pic: '图片占位符',
  chart: '图表占位符',
  tbl: '表格占位符'
})

/** 媒体类型判定（按扩展名，媒体部件不解析字节）。 */
const MEDIA_KINDS = Object.freeze({
  png: 'image',
  jpg: 'image',
  jpeg: 'image',
  gif: 'image',
  bmp: 'image',
  tif: 'image',
  tiff: 'image',
  emf: 'image',
  wmf: 'image',
  svg: 'image',
  mp4: 'video',
  m4v: 'video',
  mov: 'video',
  avi: 'video',
  wmv: 'video',
  m4a: 'audio',
  mp3: 'audio',
  wav: 'audio',
  wma: 'audio'
})

/** 一个已打开的 PPTX 演示文稿。 */
export class PptxPresentation {
  #pkg
  #docCache = new Map()

  /**
   * @param {ZipPackage} pkg - 底层 OPC 包。
   * @param {object} info - 结构信息。
   */
  constructor(pkg, info) {
    this.#pkg = pkg
    this.info = info
  }

  /**
   * 打开一个 .pptx。
   * @param {Buffer} buffer - 文件字节。
   * @param {Partial<typeof DEFAULT_ZIP_LIMITS>} [limits] - ZIP 安全上限。
   * @returns {PptxPresentation} 演示文稿对象。
   */
  static open(buffer, limits) {
    if (buffer.length >= 8 && buffer.readUInt32LE(0) === 0xe011cfd0) {
      throw new OfficeError('PASSWORD_REQUIRED', '文件是加密的 OOXML（OLE 复合文档容器），需要密码或解密副本。')
    }
    const pkg = ZipPackage.open(buffer, limits)
    if (!pkg.has('ppt/presentation.xml')) {
      throw new OfficeError('CORRUPTED_DOCUMENT', '缺少 ppt/presentation.xml，不是有效的 .pptx 文件。')
    }

    const presentationDoc = XmlDoc.parse(pkg.readText('ppt/presentation.xml'))
    const rels = readRels(pkg, 'ppt/_rels/presentation.xml.rels')
    const names = pkg.names()

    const slides = []
    const sldIdLst = find(presentationDoc.root, 'sldIdLst')
    for (const node of (sldIdLst?.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'sldId')) {
      // `<p:sldId id="256" r:id="rId2"/>` 同时带 id 与 r:id 两个属性。
      // 按本地名取 'id' 会拿到幻灯片编号 256 而不是关系 ID，必须精确取 r:id。
      const rid = relationshipId(node)
      const target = rels.get(rid)
      if (!target) continue
      slides.push({
        index: slides.length,
        rid,
        part: normalizePart('ppt', target),
        slideId: Number(attr(node, 'id') ?? 0)
      })
    }

    const sldSz = find(presentationDoc.root, 'sldSz')
    const media = names.filter((n) => n.startsWith('ppt/media/'))
    const info = {
      presentationDoc,
      rels,
      slides,
      widthEmu: Number(attr(sldSz ?? { attrs: new Map() }, 'cx') ?? 0) || null,
      heightEmu: Number(attr(sldSz ?? { attrs: new Map() }, 'cy') ?? 0) || null,
      masters: names.filter((n) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(n)).sort(),
      layouts: names.filter((n) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n)).sort(),
      themes: names.filter((n) => /^ppt\/theme\/theme\d+\.xml$/.test(n)).sort(),
      notesSlides: names.filter((n) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n)).sort(),
      charts: names.filter((n) => /^ppt\/charts\/.*\.xml$/.test(n)).sort(),
      diagrams: names.filter((n) => n.startsWith('ppt/diagrams/')),
      embeddings: names.filter((n) => n.startsWith('ppt/embeddings/')),
      media,
      mediaKinds: {
        image: media.filter((n) => MEDIA_KINDS[extOf(n)] === 'image').length,
        video: media.filter((n) => MEDIA_KINDS[extOf(n)] === 'video').length,
        audio: media.filter((n) => MEDIA_KINDS[extOf(n)] === 'audio').length,
        other: media.filter((n) => MEDIA_KINDS[extOf(n)] === undefined).length
      },
      hasMacro: names.some((n) => n.includes('vbaProject')),
      hasSignatures: names.some((n) => n.includes('_xmlsignatures/'))
    }
    return new PptxPresentation(pkg, info)
  }

  /**
   * 取得（并缓存）某个部件的 XML 文档。
   * @param {string} part - 部件路径。
   * @returns {XmlDoc} 文档对象。
   */
  #doc(part) {
    let doc = this.#docCache.get(part)
    if (!doc) {
      doc = XmlDoc.parse(this.#pkg.readText(part))
      this.#docCache.set(part, doc)
    }
    return doc
  }

  /**
   * 读取演示文稿元数据。
   * @returns {object} 元数据。
   */
  metadata() {
    const result = {
      title: null,
      subject: null,
      creator: null,
      last_modified_by: null,
      created: null,
      modified: null,
      revision: null
    }
    if (this.#pkg.has('docProps/core.xml')) {
      const doc = XmlDoc.parse(this.#pkg.readText('docProps/core.xml'))
      const root = doc.root.children.find((c) => c.type === 'element')
      const pick = (name) => {
        const node = find(root, name)
        return node ? nodeText(doc, node) : null
      }
      result.title = pick('title')
      result.subject = pick('subject')
      result.creator = pick('creator')
      result.last_modified_by = pick('lastModifiedBy')
      result.created = pick('created')
      result.modified = pick('modified')
      result.revision = pick('revision')
    }
    if (this.#pkg.has('docProps/app.xml')) {
      const doc = XmlDoc.parse(this.#pkg.readText('docProps/app.xml'))
      const root = doc.root.children.find((c) => c.type === 'element')
      const pick = (name) => {
        const node = find(root, name)
        return node ? nodeText(doc, node) : null
      }
      result.application = pick('Application')
      result.slides_reported = Number(pick('Slides') ?? 0) || null
      result.words = Number(pick('Words') ?? 0) || null
      const titles = find(root, 'TitlesOfParts')
      result.slide_titles = titles ? findAll(titles, 'lpstr').map((n) => nodeText(doc, n)) : null
    }
    return result
  }

  /**
   * 读取一张幻灯片的形状与文本。
   *
   * 只遍历 `spTree` 的直接子形状；组合形状（`p:grpSp`）递归展开但仍标记来源，
   * 避免把组合内的文本当成顶层形状。
   *
   * @param {number} index - 幻灯片下标（0 基）。
   * @returns {object} 幻灯片内容。
   */
  readSlide(index) {
    const slide = this.info.slides[index]
    if (!slide) {
      throw new OfficeError('FILE_NOT_FOUND', `不存在第 ${index} 张幻灯片（共 ${this.info.slides.length} 张）。`, {
        slide_count: this.info.slides.length
      })
    }
    const doc = this.#doc(slide.part)
    const spTree = find(doc.root, 'spTree')
    if (!spTree) throw new OfficeError('CORRUPTED_DOCUMENT', `${slide.part} 缺少 spTree。`)

    // 与写入方法共用同一个枚举器，保证「读到的第 N 个形状」就是「写入时的第 N 个」
    const shapes = []
    let title = null
    for (const { node, type, depth, name } of this.#shapeNodes(spTree)) {
      if (type === 'shape') {
        const info = shapeInfo(doc, node, depth)
        shapes.push(info)
        if (info.placeholder === 'title' || info.placeholder === 'ctrTitle') title = title ?? info.text
      } else if (type === 'table') {
        // 表格内容要读出来，否则「写入后读回验证」无从下手
        const table = tableInfo(doc, node)
        shapes.push({ type, name, depth, text: table.rows.map((r) => r.join('\t')).join('\n'), table })
      } else {
        shapes.push({ type, name, text: '', depth })
      }
    }

    // 备注：只取 type="body" 的占位符。
    // 备注页上还有 sldImg 与 sldNum 两个占位符，把它们的文本一并收进来
    // 会让「讲稿」里混入幻灯片编号。
    const rels = readRels(this.#pkg, slideRelPath(slide.part))
    let notes = null
    for (const target of rels.values()) {
      if (!/notesSlide\d+\.xml$/.test(target)) continue
      const notesPart = normalizePart('ppt/slides', target)
      if (!this.#pkg.has(notesPart)) continue
      const notesDoc = this.#doc(notesPart)
      const bodyShapes = findAll(notesDoc.root, 'sp').filter((sp) => {
        const ph = find(sp, 'ph')
        return ph !== undefined && attr(ph, 'type') === 'body'
      })
      const text = bodyShapes
        .flatMap((sp) => findAll(find(sp, 'txBody') ?? sp, 'p').map((p) => paragraphText(notesDoc, p)))
        .filter((t) => t !== '')
        .join('\n')
      notes = text
    }

    const rootElement = doc.root.children.find((c) => c.type === 'element')
    return {
      index,
      part: slide.part,
      // 隐藏标记写在幻灯片部件根元素的 show 属性上，不在 presentation.xml 的 sldId 上
      hidden: attr(rootElement ?? { attrs: new Map() }, 'show') === '0',
      title,
      shape_count: shapes.length,
      shapes,
      text: shapes
        .filter((s) => s.text)
        .map((s) => s.text)
        .join('\n'),
      notes: notes && notes !== '' ? notes : null,
      has_animation: hasAnimation(this.#pkg, slide.part),
      has_transition: hasTransition(this.#doc(slide.part)),
      layout: findLayout(this.#pkg, slide.part)
    }
  }

  /**
   * 读取全部幻灯片。
   * @param {object} [options] - 选项。
   * @param {boolean} [options.includeShapes] - 是否包含每个形状的明细，默认 true。
   * @returns {object} `{count, slides}`。
   */
  slides({ includeShapes = true } = {}) {
    const out = this.info.slides.map((_, i) => {
      const slide = this.readSlide(i)
      if (includeShapes) return slide
      const { shapes, ...rest } = slide
      return { ...rest, shape_count: shapes.length }
    })
    return { count: out.length, slides: out }
  }

  /**
   * 提取全部文本。
   * @param {number} [maxChars] - 字符上限。
   * @returns {object} `{text, truncated, length}`。
   */
  text(maxChars = 200000) {
    const blocks = this.info.slides.map((_, i) => {
      const slide = this.readSlide(i)
      const header = `--- 第 ${i + 1} 张幻灯片${slide.title ? `：${slide.title}` : ''} ---`
      return `${header}\n${slide.text}`
    })
    const text = blocks.join('\n\n')
    return { text: text.slice(0, maxChars), truncated: text.length > maxChars, length: text.length }
  }

  /**
   * 结构识别与风险提示。
   * @returns {object} 结构统计。
   */
  structure() {
    const slides = this.info.slides.map((_, i) => this.readSlide(i))
    const animated = slides.filter((s) => s.has_animation).length
    const transitioned = slides.filter((s) => s.has_transition).length
    return {
      slide_count: this.info.slides.length,
      hidden_slides: slides.filter((s) => s.hidden).map((s) => s.index),
      slide_size: {
        width_emu: this.info.widthEmu,
        height_emu: this.info.heightEmu,
        width_inch: this.info.widthEmu ? Math.round((this.info.widthEmu / 914400) * 100) / 100 : null,
        height_inch: this.info.heightEmu ? Math.round((this.info.heightEmu / 914400) * 100) / 100 : null,
        aspect: this.info.widthEmu && this.info.heightEmu ? aspectName(this.info.widthEmu, this.info.heightEmu) : null
      },
      masters: this.info.masters.length,
      layouts: this.info.layouts.length,
      themes: this.info.themes.length,
      notes_slides: this.info.notesSlides.length,
      charts: this.info.charts.length,
      smartart_parts: this.info.diagrams.length,
      embedded_objects: this.info.embeddings.length,
      media: this.info.media.length,
      media_kinds: this.info.mediaKinds,
      animated_slides: animated,
      slides_with_transition: transitioned,
      has_speaker_notes: slides.some((s) => s.notes !== null),
      shapes_without_text: slides.reduce((n, s) => n + s.shapes.filter((x) => x.text === '' && x.type === 'shape').length, 0),
      has_macro: this.info.hasMacro,
      has_digital_signature: this.info.hasSignatures
    }
  }

  /**
   * 校验演示文稿结构。
   * @returns {object} `{valid, checks, not_checked}`。
   */
  validate() {
    const checks = []
    const pkg = this.#pkg
    const names = new Set(pkg.names())

    checks.push({ name: 'ZIP 包结构', ok: true, detail: `${names.size} 个部件` })
    const required = ['[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml']
    const missing = required.filter((p) => !names.has(p))
    checks.push({ name: '必需部件齐全', ok: missing.length === 0, detail: missing.length ? `缺少 ${missing.join(', ')}` : '全部存在' })

    let brokenSlide = null
    for (const slide of this.info.slides) {
      if (!names.has(slide.part)) {
        brokenSlide = `缺少 ${slide.part}`
        break
      }
      try {
        XmlDoc.parse(pkg.readText(slide.part))
      } catch (err) {
        brokenSlide = `${slide.part}: ${err.message}`
        break
      }
    }
    checks.push({ name: '幻灯片部件可解析', ok: brokenSlide === null, detail: brokenSlide ?? `${this.info.slides.length} 张全部可解析` })

    const rels = this.info.rels
    const dangling = []
    for (const slide of this.info.slides) {
      if (!rels.get(slide.rid)) dangling.push(slide.rid)
    }
    checks.push({ name: '幻灯片关系有效', ok: dangling.length === 0, detail: dangling.length ? `悬空：${dangling.join(', ')}` : `${this.info.slides.length} 条关系有效` })

    // 每张幻灯片的 r:embed / 图表 r:id 是否都有对应关系
    const missingEmbeds = []
    for (const slide of this.info.slides) {
      if (!names.has(slide.part)) continue
      const doc = this.#doc(slide.part)
      const slideRels = readRels(pkg, slideRelPath(slide.part))
      for (const blip of findAll(doc.root, 'blip')) {
        const embed = attr(blip, 'embed')
        if (!embed) continue
        const target = slideRels.get(embed)
        if (!target) missingEmbeds.push(`${slide.part}:${embed}`)
        else if (!names.has(normalizePart('ppt/slides', target))) missingEmbeds.push(`${slide.part}:${embed}→${target}`)
      }
      // 图表引用：`<c:chart r:id>` 指向图表部件，目标不存在时 PowerPoint 同样打不开
      for (const chartRef of findAll(doc.root, 'chart')) {
        const rid = attr(chartRef, 'id')
        if (!rid) continue
        const target = slideRels.get(rid)
        if (!target) missingEmbeds.push(`${slide.part}:${rid}（图表）`)
        else if (!names.has(normalizePart('ppt/slides', target))) missingEmbeds.push(`${slide.part}:${rid}→${target}`)
      }
    }
    checks.push({
      name: '图片/图表引用可解析',
      ok: missingEmbeds.length === 0,
      detail: missingEmbeds.length ? `悬空：${missingEmbeds.join(', ')}` : '全部可解析'
    })

    // 图表部件自身的 externalData 关系是否指到真实存在的嵌入工作簿
    const brokenChartLinks = []
    for (const chartPart of names) {
      if (!/^ppt\/charts\/chart\d+\.xml$/.test(chartPart)) continue
      const chartDoc = XmlDoc.parse(pkg.readText(chartPart))
      const external = find(chartDoc.root, 'externalData')
      const rid = external ? attr(external, 'id') : undefined
      if (!rid) continue
      const chartRels = readRels(pkg, `ppt/charts/_rels/${chartPart.split('/').pop()}.rels`)
      const target = chartRels.get(rid)
      if (!target) brokenChartLinks.push(`${chartPart}:${rid}`)
      else if (!names.has(normalizePart('ppt/charts', target))) brokenChartLinks.push(`${chartPart}:${rid}→${target}`)
    }
    checks.push({
      name: '图表嵌入工作簿可解析',
      ok: brokenChartLinks.length === 0,
      detail: brokenChartLinks.length ? `悬空：${brokenChartLinks.join(', ')}` : '全部可解析'
    })

    checks.push({
      name: '母版与主题保留',
      ok: this.info.masters.length > 0,
      detail: `${this.info.masters.length} 母版 / ${this.info.layouts.length} 版式 / ${this.info.themes.length} 主题`
    })

    const structure = this.structure()
    const warnings = []
    if (structure.animated_slides > 0) warnings.push(`${structure.animated_slides} 张幻灯片含动画，插件不修改动画`)
    if (structure.smartart_parts > 0) warnings.push(`${structure.smartart_parts} 个 SmartArt 部件，插件不解析其内容`)
    if (structure.media_kinds.video > 0 || structure.media_kinds.audio > 0) {
      warnings.push(`含 ${structure.media_kinds.video} 个视频 / ${structure.media_kinds.audio} 个音频，插件保留但不处理`)
    }

    return {
      valid: checks.every((c) => c.ok),
      checks,
      warnings,
      not_checked: PPTX_NOT_CHECKED,
      slide_size: structure.slide_size
    }
  }

  /**
   * 按固定顺序枚举一张幻灯片的形状节点，并给出**细化后的类型**。
   *
   * `readSlide` 与写入方法**必须共用这一个枚举器**：如果读和写各写一套遍历、
   * 或各自做一次类型细化（例如把 graphicFrame 细分出 table / chart），
   * 「读到的第 2 个形状」与「写入的第 2 个形状」就会指向不同对象或报出不同类型。
   * 这类错位在结构简单的样本上不会暴露，因此细化只在这里做一次。
   *
   * @param {object} spTree - `<p:spTree>` 节点。
   * @returns {object[]} `{node, type, depth, name}` 列表。
   */
  #shapeNodes(spTree) {
    const out = []
    const walk = (node, depth) => {
      for (const child of node.children ?? []) {
        if (child.type !== 'element') continue
        const name = localName(child)
        if (name === 'sp') out.push({ node: child, type: 'shape', depth, name: shapeName(child) })
        else if (name === 'pic') out.push({ node: child, type: 'picture', depth, name: shapeName(child) })
        else if (name === 'cxnSp') out.push({ node: child, type: 'connector', depth, name: shapeName(child) })
        else if (name === 'graphicFrame') out.push({ node: child, type: graphicFrameKind(child), depth, name: shapeName(child) })
        else if (name === 'grpSp') {
          out.push({ node: child, type: 'group', depth, name: shapeName(child) })
          walk(child, depth + 1)
        }
      }
    }
    walk(spTree, 0)
    return out
  }

  /**
   * 取得（并缓存）一张幻灯片文档与它的形状节点列表。
   * @param {number} index - 幻灯片下标。
   * @returns {{slide: object, doc: XmlDoc, nodes: object[]}} 幻灯片上下文。
   */
  #slideContext(index) {
    const slide = this.info.slides[index]
    if (!slide) {
      throw new OfficeError('FILE_NOT_FOUND', `不存在第 ${index} 张幻灯片（共 ${this.info.slides.length} 张）。`, {
        slide_count: this.info.slides.length
      })
    }
    const doc = this.#doc(slide.part)
    const spTree = find(doc.root, 'spTree')
    if (!spTree) throw new OfficeError('CORRUPTED_DOCUMENT', `${slide.part} 缺少 spTree。`)
    return { slide, doc, nodes: this.#shapeNodes(spTree) }
  }

  /**
   * 改写一张幻灯片里某个形状的文本。
   *
   * 每个 `\n` 生成一个新的 `<a:p>` 段落，并复用原第一段的 `<a:pPr>`
   * （保住项目符号层级与缩进）。只改目标形状的 `<p:txBody>`，
   * 其余形状、版式、母版、主题与动画一律不动。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 幻灯片下标（0 基）。
   * @param {number} args.shape - 形状下标（0 基，与 `readSlide().shapes` 同序）。
   * @param {string} args.text - 新文本，`\n` 分段。
   * @param {boolean} [args.allowMarkupLoss] - 明知会丢失域/超链接仍继续。
   * @returns {object} 变更信息。
   */
  updateSlideText({ index, shape, text, allowMarkupLoss = false }) {
    const { doc, nodes, slide } = this.#slideContext(index)
    const target = nodes[shape]
    if (!target) {
      throw new OfficeError('FILE_NOT_FOUND', `第 ${index} 张幻灯片没有第 ${shape} 个形状（共 ${nodes.length} 个）。`, {
        shape_count: nodes.length
      })
    }
    // 只有 <p:sp> 才有自己的文本体。
    // 不能用递归查找：表格（graphicFrame）的单元格里也有 <a:txBody>，
    // 递归找会把表格当成「可写文本的形状」，然后在错误的节点上打补丁。
    const txBody =
      target.type === 'shape'
        ? (target.node.children ?? []).find((n) => n.type === 'element' && localName(n) === 'txBody')
        : undefined
    if (!txBody) {
      throw new OfficeError('UNSUPPORTED_FEATURE', `第 ${shape} 个形状是 ${target.type}，没有自己的文本体，无法写入文本。`, {
        shape_type: target.type,
        hint: target.type === 'table' ? '表格请改用对应表格工具（尚未实现）；当前版本只支持改写形状文本。' : undefined
      })
    }

    // 保护性检查：域（页码/日期）与超链接会被整段替换抹掉
    const protectedNames = []
    for (const tag of ['fld', 'hlinkClick']) {
      if (find(txBody, tag)) protectedNames.push(tag === 'fld' ? '域（页码/日期）' : '超链接')
    }
    if (protectedNames.length > 0 && !allowMarkupLoss) {
      throw new OfficeError('UNSUPPORTED_FEATURE', `该形状含${protectedNames.join('、')}，改写文本会使其失效。`, { shape, protected: protectedNames }, {
        solution: '如确实要改写，请显式传 allow_markup_loss=true 接受该损失；或改用其它形状。',
        needsConfirmation: true
      })
    }

    const paragraphs = (txBody.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'p')
    const previous = paragraphs.map((p) => paragraphText(doc, p)).join('\n')
    // 复用第一段的段落属性（项目符号层级、缩进、字号）
    const templatePPr = paragraphs[0] ? find(paragraphs[0], 'pPr') : null
    const pPrXml = templatePPr ? doc.raw(templatePPr) : ''
    const xml = String(text ?? '')
      .split('\n')
      .map((line) => `<a:p>${pPrXml}${textRunsXml(line)}</a:p>`)
      .join('')

    if (paragraphs.length > 0) {
      doc.patch(paragraphs[0].start, paragraphs[paragraphs.length - 1].end, xml)
    } else {
      doc.appendChild(txBody, xml)
    }

    this.#pkg.write(slide.part, doc.toString())
    doc.rescan()
    return {
      type: 'update_slide_text',
      slide: index,
      shape,
      from: preview(previous),
      to: preview(String(text ?? '')),
      paragraphs: String(text ?? '').split('\n').length,
      reused_paragraph_properties: pPrXml !== ''
    }
  }

  /**
   * 调整幻灯片顺序。
   * @param {object} args - 参数。
   * @param {number} args.from - 原下标（0 基）。
   * @param {number} args.to - 目标下标（0 基）。
   * @returns {object} 变更信息。
   */
  moveSlide({ from, to }) {
    const count = this.info.slides.length
    if (!Number.isInteger(from) || from < 0 || from >= count) {
      throw new OfficeError('INVALID_REQUEST', `原下标 ${from} 超出范围（共 ${count} 张）。`, { slide_count: count })
    }
    if (!Number.isInteger(to) || to < 0 || to >= count) {
      throw new OfficeError('INVALID_REQUEST', `目标下标 ${to} 超出范围（共 ${count} 张）。`, { slide_count: count })
    }
    if (from === to) return { type: 'move_slide', from, to, changed: false }

    const sldIdLst = find(this.info.presentationDoc.root, 'sldIdLst')
    if (!sldIdLst) throw new OfficeError('CORRUPTED_DOCUMENT', 'presentation.xml 缺少 sldIdLst。')
    const nodes = (sldIdLst.children ?? []).filter((n) => n.type === 'element' && localName(n) === 'sldId')
    if (nodes.length !== count) {
      throw new OfficeError('CORRUPTED_DOCUMENT', `sldIdLst 中有 ${nodes.length} 个条目，与解析出的 ${count} 张幻灯片不一致。`)
    }

    // 重排 = 把这一个 <p:sldId> 节点整体搬走（删除 + 在目标位置插入）
    const moving = nodes[from]
    const xml = this.info.presentationDoc.raw(moving)
    const anchor = nodes[to]
    this.info.presentationDoc.remove(moving)
    if (to < from) this.info.presentationDoc.insertBefore(anchor, xml)
    else this.info.presentationDoc.insertAfter(anchor, xml)
    this.#commitPresentation()

    const [entry] = this.info.slides.splice(from, 1)
    this.info.slides.splice(to, 0, entry)
    this.info.slides.forEach((s, i) => {
      s.index = i
    })
    return { type: 'move_slide', from, to, changed: true, order: this.info.slides.map((s) => s.part) }
  }

  /**
   * 删除一张幻灯片。
   *
   * 会一并删除该幻灯片部件、它的关系部件与内容类型声明，
   * 不留孤儿部件 —— 只从 `sldIdLst` 摘掉引用会让文件里留下
   * 永远不会被引用、却仍占体积的幻灯片。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 幻灯片下标（0 基）。
   * @returns {object} 变更信息。
   */
  deleteSlide({ index }) {
    const count = this.info.slides.length
    if (count <= 1) {
      throw new OfficeError('INVALID_REQUEST', '演示文稿只剩一张幻灯片，删除后将不再是有效演示文稿。')
    }
    const slide = this.info.slides[index]
    if (!slide) {
      throw new OfficeError('INVALID_REQUEST', `幻灯片下标 ${index} 超出范围（共 ${count} 张）。`, { slide_count: count })
    }

    const sldIdLst = find(this.info.presentationDoc.root, 'sldIdLst')
    const node = (sldIdLst?.children ?? []).find((n) => n.type === 'element' && localName(n) === 'sldId' && relationshipId(n) === slide.rid)
    if (node) this.info.presentationDoc.remove(node)
    this.#commitPresentation()

    removeRelationshipFrom(this.#pkg, 'ppt/_rels/presentation.xml.rels', slide.rid)
    removeContentTypeOverride(this.#pkg, `/${slide.part}`)
    this.#pkg.delete(slide.part)
    this.#pkg.delete(slideRelPath(slide.part))
    this.#docCache.delete(slide.part)

    this.info.slides.splice(index, 1)
    this.info.slides.forEach((s, i) => {
      s.index = i
    })
    return { type: 'delete_slide', index, removed_part: slide.part, slide_count: this.info.slides.length }
  }

  /**
   * 新增一张空白幻灯片。
   *
   * 参考既有幻灯片来继承版式：新页面引用与参考页相同的 slideLayout，
   * 这样标题/正文占位符的样式与母版保持一致，不需要自己造版式。
   *
   * @param {object} args - 参数。
   * @param {number} [args.after] - 插入到第几张之后（0 基）；-1 表示插到最前；省略则追加到末尾。
   * @param {number} [args.layoutOf] - 参照哪张幻灯片取版式，默认第 0 张。
   * @returns {object} 变更信息。
   */
  addSlide({ after = null, layoutOf = 0 } = {}) {
    const reference = this.info.slides[layoutOf]
    if (!reference) {
      throw new OfficeError('INVALID_REQUEST', `参照幻灯片下标 ${layoutOf} 不存在（共 ${this.info.slides.length} 张）。`, {
        slide_count: this.info.slides.length
      })
    }
    const referenceLayout = findLayout(this.#pkg, reference.part)
    if (!referenceLayout) {
      throw new OfficeError('CORRUPTED_DOCUMENT', `参照幻灯片 ${reference.part} 没有版式关系，无法继承版式。`)
    }
    // 关系目标是相对 `ppt/slides/` 的路径；版式在 `ppt/slideLayouts/`，
    // 因此必须写成 `../slideLayouts/...`。少这一层 `../` 会让 Word 找不到版式，
    // 而且文件本身仍能打开、只是版式静默丢失。
    const layoutTarget = `../${referenceLayout.replace(/^ppt\//, '')}`

    // 新的部件名与幻灯片编号都要避开已占用的值
    let n = 1
    while (this.#pkg.has(`ppt/slides/slide${n}.xml`)) n += 1
    const part = `ppt/slides/slide${n}.xml`
    this.#pkg.write(part, EMPTY_SLIDE_XML)
    this.#pkg.write(
      slideRelPath(part),
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS_R}/slideLayout" Target="${escapeXmlAttr(layoutTarget)}"/></Relationships>`
    )
    ensureContentTypeOverride(this.#pkg, `/${part}`, CT_SLIDE)

    const rid = addRelationshipTo(this.#pkg, 'ppt/_rels/presentation.xml.rels', `${NS_R}/slide`, part.replace(/^ppt\//, ''))
    const usedIds = this.info.slides.map((s) => s.slideId)
    let slideId = 256
    while (usedIds.includes(slideId)) slideId += 1

    const sldIdLst = find(this.info.presentationDoc.root, 'sldIdLst')
    if (!sldIdLst) throw new OfficeError('CORRUPTED_DOCUMENT', 'presentation.xml 缺少 sldIdLst。')
    const xml = `<p:sldId id="${slideId}" r:id="${rid}"/>`
    const nodes = (sldIdLst.children ?? []).filter((x) => x.type === 'element' && localName(x) === 'sldId')
    if (after === null || after === undefined) {
      this.info.presentationDoc.appendChild(sldIdLst, xml)
    } else if (after < 0) {
      if (nodes[0]) this.info.presentationDoc.insertBefore(nodes[0], xml)
      else this.info.presentationDoc.appendChild(sldIdLst, xml)
    } else {
      const anchor = nodes[after]
      if (anchor) this.info.presentationDoc.insertAfter(anchor, xml)
      else this.info.presentationDoc.appendChild(sldIdLst, xml)
    }
    this.#commitPresentation()

    const position = after === null || after === undefined ? this.info.slides.length : Math.min(after + 1, this.info.slides.length)
    const entry = { index: position, rid, part, slideId }
    this.info.slides.splice(position, 0, entry)
    this.info.slides.forEach((s, i) => {
      s.index = i
    })
    return { type: 'add_slide', index: position, part, relationship_id: rid, slide_id: slideId, layout: referenceLayout, slide_count: this.info.slides.length }
  }

  /**
   * 复制一张幻灯片（插在源页之后）。
   *
   * 幻灯片正文与它的关系部件**整体复制**：关系里的目标（版式、图片、图表）都是相对
   * `ppt/slides/` 的路径，复制后依然有效，因此不需要搬运媒体部件。
   * **演讲者备注不复制**（`notesSlide` 关系会被去掉）—— 两个幻灯片共用一个备注部件
   * 属于无效结构，PowerPoint 会要求修复；要备注请复制后单独设置。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 源幻灯片下标（0 基）。
   * @returns {object} 变更信息。
   */
  duplicateSlide({ index }) {
    const source = this.info.slides[index]
    if (!source) {
      throw new OfficeError('FILE_NOT_FOUND', `不存在第 ${index} 张幻灯片（共 ${this.info.slides.length} 张）。`, {
        slide_count: this.info.slides.length
      })
    }
    let n = 1
    while (this.#pkg.has(`ppt/slides/slide${n}.xml`)) n += 1
    const part = `ppt/slides/slide${n}.xml`
    this.#pkg.write(part, this.#pkg.readText(source.part))
    ensureContentTypeOverride(this.#pkg, `/${part}`, CT_SLIDE)

    // 复制关系部件，但去掉 notesSlide（共用备注部件会让 PowerPoint 报修复），
    // 图表关系则**连图表部件一起深拷贝**：两页共用一个 chart 部件时 PowerPoint 直接拒绝打开
    // （PowerPoint 自己复制幻灯片时也会新建 chartN.xml，不共用）。
    const sourceRels = slideRelPath(source.part)
    let notesDropped = false
    let chartsCloned = 0
    if (this.#pkg.has(sourceRels)) {
      const relsDoc = XmlDoc.parse(this.#pkg.readText(sourceRels))
      for (const rel of findAll(relsDoc.root, 'Relationship')) {
        const type = attr(rel, 'Type') ?? ''
        if (type.endsWith('/notesSlide')) {
          relsDoc.remove(rel)
          notesDropped = true
        } else if (type.endsWith('/chart')) {
          const cloned = this.#cloneChartPart(attr(rel, 'Target') ?? '')
          relsDoc.patch(rel.start, rel.end, `<Relationship Id="${attr(rel, 'Id')}" Type="${type}" Target="${cloned}"/>`)
          chartsCloned += 1
        }
      }
      this.#pkg.write(slideRelPath(part), relsDoc.toString())
    }

    const rid = addRelationshipTo(this.#pkg, 'ppt/_rels/presentation.xml.rels', `${NS_R}/slide`, part.replace(/^ppt\//, ''))
    const usedIds = this.info.slides.map((s) => s.slideId)
    let slideId = 256
    while (usedIds.includes(slideId)) slideId += 1

    const sldIdLst = find(this.info.presentationDoc.root, 'sldIdLst')
    if (!sldIdLst) throw new OfficeError('CORRUPTED_DOCUMENT', 'presentation.xml 缺少 sldIdLst。')
    const nodes = (sldIdLst.children ?? []).filter((x) => x.type === 'element' && localName(x) === 'sldId')
    const xml = `<p:sldId id="${slideId}" r:id="${rid}"/>`
    const anchor = nodes[index]
    if (anchor) this.info.presentationDoc.insertAfter(anchor, xml)
    else this.info.presentationDoc.appendChild(sldIdLst, xml)
    this.#commitPresentation()

    const position = index + 1
    this.info.slides.splice(position, 0, { index: position, rid, part, slideId })
    this.info.slides.forEach((s, i) => {
      s.index = i
    })
    return {
      type: 'duplicate_slide',
      source: index,
      index: position,
      part,
      slide_id: slideId,
      relationship_id: rid,
      slide_count: this.info.slides.length,
      notes_dropped: notesDropped
    }
  }

  /**
   * 往一张幻灯片里插入文本框。
   *
   * 形状要写齐 `p:nvSpPr`（非可视属性 + 形状 id/名称）、`p:spPr`（位置尺寸 + 几何）
   * 与 `p:txBody`（正文）三段；形状 id 必须**全页唯一**，否则 PowerPoint 报修复。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 幻灯片下标（0 基）。
   * @param {string} args.text - 文本（`\n` 分段）。
   * @param {number} [args.leftPx] - 左边距（像素），默认 80。
   * @param {number} [args.topPx] - 上边距（像素），默认 80。
   * @param {number} [args.widthPx] - 宽（像素），默认 400。
   * @param {number} [args.heightPx] - 高（像素），默认 100。
   * @param {number} [args.fontSizePt] - 字号（磅），默认 18。
   * @param {boolean} [args.bold] - 是否加粗。
   * @param {string} [args.align] - `left` | `center` | `right`。
   * @returns {object} 变更信息。
   */
  addTextBox({ index, text, leftPx = 80, topPx = 80, widthPx = 400, heightPx = 100, fontSizePt = 18, bold = false, align = 'left' }) {
    const { doc, slide } = this.#slideContext(index)
    const spTree = find(doc.root, 'spTree')
    if (!spTree) throw new OfficeError('CORRUPTED_DOCUMENT', `${slide.part} 缺少 spTree。`)
    const alignMap = { left: 'l', center: 'ctr', right: 'r' }
    if (!alignMap[align]) throw new OfficeError('INVALID_REQUEST', `align 只支持 left / center / right，实际 ${align}。`)

    const maxId = findAll(spTree, 'cNvPr').reduce((max, node) => Math.max(max, Number(attr(node, 'id') ?? 0)), 1)
    const shapeId = maxId + 1
    const cx = Math.round(widthPx * EMU_PER_PX)
    const cy = Math.round(heightPx * EMU_PER_PX)
    const offX = Math.round(leftPx * EMU_PER_PX)
    const offY = Math.round(topPx * EMU_PER_PX)
    const size = Math.round(fontSizePt * 100)
    const paragraphs = String(text ?? '')
      .split('\n')
      .map(
        (line) =>
          `<a:p><a:pPr algn="${alignMap[align]}"/><a:r><a:rPr lang="zh-CN" sz="${size}"${bold ? ' b="1"' : ''} dirty="0"/>` +
          `<a:t>${escapeXmlText(line)}</a:t></a:r></a:p>`
      )
      .join('')
    const xml =
      `<p:sp><p:nvSpPr><p:cNvPr id="${shapeId}" name="文本框 ${shapeId}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${offX}" y="${offY}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>` +
      `<p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`
    doc.appendChild(spTree, xml)
    this.#pkg.write(slide.part, doc.toString())
    doc.rescan()
    return {
      type: 'add_text_box',
      index,
      shape_id: shapeId,
      text,
      position_px: { left: leftPx, top: topPx, width: widthPx, height: heightPx },
      font_size_pt: fontSizePt,
      bold,
      align
    }
  }

  /**
   * 往一张幻灯片里插入一张图片。
   *
   * 需要新增四处，缺一个 PowerPoint 都会报演示文稿损坏：
   *   1. `ppt/media/` 下的媒体部件
   *   2. `[Content_Types].xml` 里该扩展名的 Default 声明
   *   3. 该幻灯片的 `_rels` 里的 image 关系
   *   4. `<p:spTree>` 里引用该关系的 `<p:pic>`
   *
   * 尺寸未指定时用图片原始像素（96 DPI），并且**按幻灯片居中**放置 ——
   * 默认落在左上角会盖住标题占位符，那是几乎没人想要的默认值。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 幻灯片下标（0 基）。
   * @param {Buffer} args.data - 图片字节。
   * @param {string} args.extension - 图片扩展名（已按真实字节判定）。
   * @param {number} [args.leftPx] - 左边距（像素）；省略则水平居中。
   * @param {number} [args.topPx] - 上边距（像素）；省略则垂直居中。
   * @param {number} [args.widthPx] - 显示宽度（像素）；省略则用原始宽度。
   * @param {number} [args.heightPx] - 显示高度（像素）；省略则按比例。
   * @param {string} [args.altText] - 替代文本。
   * @returns {object} 变更信息。
   */
  insertSlideImage({ index, data, extension, leftPx, topPx, widthPx, heightPx, altText }) {
    const { doc, slide } = this.#slideContext(index)
    const spTree = find(doc.root, 'spTree')
    if (!spTree) throw new OfficeError('CORRUPTED_DOCUMENT', `${slide.part} 缺少 spTree。`)

    const slideWidthEmu = this.info.widthEmu ?? 12192000
    const slideHeightEmu = this.info.heightEmu ?? 6858000

    let n = 1
    while (this.#pkg.has(`ppt/media/image${n}.${extension}`)) n += 1
    const mediaPart = `ppt/media/image${n}.${extension}`
    this.#pkg.write(mediaPart, data)
    ensureContentTypeDefault(this.#pkg, extension, imageContentType(extension))
    const rid = addRelationshipTo(this.#pkg, slideRelPath(slide.part), REL_IMAGE, `../media/image${n}.${extension}`)

    const natural = readImageSize(data, extension)
    const { cx, cy } = resolveImageExtent({ natural, widthPx, heightPx })
    const offX = leftPx === undefined ? Math.max(0, Math.round((slideWidthEmu - cx) / 2)) : Math.round(leftPx * EMU_PER_PX)
    const offY = topPx === undefined ? Math.max(0, Math.round((slideHeightEmu - cy) / 2)) : Math.round(topPx * EMU_PER_PX)
    const id = this.#nextShapeId(doc)

    doc.appendChild(spTree, buildPicXml({ rid, id, offX, offY, cx, cy, altText }))
    this.#pkg.write(slide.part, doc.toString())
    doc.rescan()

    return {
      type: 'insert_slide_image',
      slide: index,
      media_part: mediaPart,
      relationship_id: rid,
      natural_size: natural,
      display_size_px: { width: Math.round(cx / EMU_PER_PX), height: Math.round(cy / EMU_PER_PX) },
      offset_emu: { x: offX, y: offY },
      bytes: data.length
    }
  }

  /**
   * 取一个未被占用的形状 id。
   *
   * `p:cNvPr/@id` 在同一张幻灯片内必须唯一，PowerPoint 对重复 id 会报错，
   * 因此扫描既有形状取最大值 +1。
   *
   * @param {XmlDoc} doc - 幻灯片文档。
   * @returns {number} 可用的 id。
   */
  #nextShapeId(doc) {
    let max = 1
    for (const node of findAll(doc.root, 'cNvPr')) {
      const id = Number(attr(node, 'id') ?? 0)
      if (Number.isFinite(id) && id > max) max = id
    }
    return max + 1
  }

  /**
   * 往一张幻灯片里插入一个表格。
   *
   * 表格在 OOXML 里是 `graphicFrame` + `a:tbl`，不需要新增部件或关系
   * （除非引用表格样式部件，见下），因此比插图简单。
   *
   * 默认套用 PowerPoint 内置的「中档样式 2 - 强调 1」样式 ID，
   * 否则表格会是完全无边框的裸文本 —— 那不像用户期望的「插入一个表格」。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 幻灯片下标（0 基）。
   * @param {string[][]} args.rows - 二维文本，第一维是行。
   * @param {number} [args.leftPx] - 左边距（像素）；省略则水平居中。
   * @param {number} [args.topPx] - 上边距（像素）；省略则垂直居中。
   * @param {number} [args.widthPx] - 总宽度（像素）；省略按列数估算。
   * @param {number} [args.heightPx] - 总高度（像素）；省略按行数估算。
   * @param {boolean} [args.firstRowHeader] - 是否把首行当表头，默认 true。
   * @returns {object} 变更信息。
   */
  addSlideTable({ index, rows, leftPx, topPx, widthPx, heightPx, firstRowHeader = true }) {
    const { doc, slide } = this.#slideContext(index)
    const spTree = find(doc.root, 'spTree')
    if (!spTree) throw new OfficeError('CORRUPTED_DOCUMENT', `${slide.part} 缺少 spTree。`)

    if (!Array.isArray(rows) || rows.length === 0) {
      throw new OfficeError('INVALID_REQUEST', 'rows 必须是非空的二维数组。')
    }
    for (const row of rows) {
      if (!Array.isArray(row)) throw new OfficeError('INVALID_REQUEST', 'rows 的每一项都必须是数组。')
    }
    const columnCount = Math.max(...rows.map((r) => r.length))
    if (columnCount === 0) throw new OfficeError('INVALID_REQUEST', '表格至少需要一列。')
    const rowCount = rows.length

    const slideWidthEmu = this.info.widthEmu ?? 12192000
    const slideHeightEmu = this.info.heightEmu ?? 6858000
    const cx = Math.round((widthPx ?? Math.min(720, columnCount * 160)) * EMU_PER_PX)
    const cy = Math.round((heightPx ?? Math.min(480, rowCount * 40)) * EMU_PER_PX)
    const offX = leftPx === undefined ? Math.max(0, Math.round((slideWidthEmu - cx) / 2)) : Math.round(leftPx * EMU_PER_PX)
    const offY = topPx === undefined ? Math.max(0, Math.round((slideHeightEmu - cy) / 2)) : Math.round(topPx * EMU_PER_PX)

    const id = this.#nextShapeId(doc)
    doc.appendChild(spTree, buildTableGraphicFrameXml({ id, offX, offY, cx, cy, rows, columnCount, firstRowHeader }))
    this.#pkg.write(slide.part, doc.toString())
    doc.rescan()

    return {
      type: 'add_slide_table',
      slide: index,
      shape_id: id,
      rows: rowCount,
      columns: columnCount,
      offset_emu: { x: offX, y: offY },
      size_emu: { cx, cy },
      first_row_header: firstRowHeader === true
    }
  }

  /**
   * 列出全部版式（含名称与类型），供 `setSlideLayout` 选择。
   *
   * 名称取自 `<p:cSld name="…">`（PowerPoint 界面里显示的就是这个名字），
   * 类型取自 `<p:sldLayout type="…">`（如 `title` / `blank` / `obj`）。
   *
   * @returns {object[]} 版式列表。
   */
  layouts() {
    return this.info.layouts.map((part, index) => {
      const doc = XmlDoc.parse(this.#pkg.readText(part))
      // 注意：`doc.root` 是 `#document` 容器节点，根**元素**是它的第一个 element 子节点。
      // 直接 `attr(doc.root, …)` 永远拿不到属性（会静默返回 null）。
      const root = doc.root.children.find((c) => c.type === 'element')
      const cSld = find(doc.root, 'cSld')
      const rawName = attr(cSld ?? { attrs: new Map() }, 'name') ?? null
      return {
        index,
        part,
        name: cleanOfficeName(rawName),
        raw_name: rawName,
        type: attr(root ?? { attrs: new Map() }, 'type') ?? null,
        placeholder_count: findAll(doc.root, 'ph').length,
        // 该版式当前被哪些幻灯片使用（0 基下标）
        used_by: this.info.slides.filter((s) => findLayout(this.#pkg, s.part) === part).map((s) => s.index)
      }
    })
  }

  /**
   * 切换一张幻灯片使用的版式。
   *
   * 幻灯片的版式关联**只由关系决定**：`ppt/slides/_rels/slideN.xml.rels` 里那条
   * `…/slideLayout` 关系的 `Target` 指向哪个版式部件，PowerPoint 就按哪个版式渲染。
   * 所以这里改的是关系的目标，而不是往幻灯片 XML 里塞引用元素。
   *
   * **不会**挪动或删除已有形状：占位符的形状下标与文本留在原处；若新版式里没有
   * 对应占位符，它们会以「自由形状」的样子继续显示（PowerPoint 自己的行为也是如此）。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 幻灯片下标（0 基）。
   * @param {number} [args.layout] - 目标版式下标（0 基，见 `layouts()`）。
   * @param {string} [args.name] - 目标版式名（与 `layout` 二选一，重名时取第一个）。
   * @returns {object} 变更信息。
   */
  setSlideLayout({ index, layout, name }) {
    const slide = this.info.slides[index]
    if (!slide) {
      throw new OfficeError('FILE_NOT_FOUND', `不存在第 ${index} 张幻灯片（共 ${this.info.slides.length} 张）。`, {
        slide_count: this.info.slides.length
      })
    }
    if (layout === undefined && name === undefined) {
      throw new OfficeError('INVALID_REQUEST', '必须提供 layout（下标）或 name（版式名）之一。')
    }
    let target
    if (layout !== undefined) {
      target = this.info.layouts[layout]
      if (!target) {
        throw new OfficeError('INVALID_REQUEST', `版式下标 ${layout} 不存在（共 ${this.info.layouts.length} 个）。`, {
          layout_count: this.info.layouts.length
        })
      }
    } else {
      const wanted = cleanOfficeName(name)?.toLowerCase() ?? ''
      const hit = this.layouts().find((item) => item.name?.toLowerCase() === wanted)
      if (!hit) {
        throw new OfficeError('INVALID_REQUEST', `找不到名为「${name}」的版式。`, {
          layout_names: this.layouts().map((l) => l.name)
        })
      }
      target = hit.part
    }

    const relsPart = slideRelPath(slide.part)
    if (!this.#pkg.has(relsPart)) throw new OfficeError('CORRUPTED_DOCUMENT', `${slide.part} 缺少关系部件。`)
    const previous = findLayout(this.#pkg, slide.part)
    const doc = XmlDoc.parse(this.#pkg.readText(relsPart))
    const rel = findAll(doc.root, 'Relationship').find((r) => (attr(r, 'Type') ?? '').endsWith('/slideLayout'))
    if (!rel) throw new OfficeError('CORRUPTED_DOCUMENT', `${relsPart} 里没有 slideLayout 关系。`)
    const relId = attr(rel, 'Id')
    // 目标是相对 `ppt/slides/` 的路径：版式在 `ppt/slideLayouts/`，因此必须带 `../`
    doc.patch(
      rel.start,
      rel.end,
      `<Relationship Id="${relId}" Type="${NS_R}/slideLayout" Target="../slideLayouts/${target.split('/').pop()}"/>`
    )
    this.#pkg.write(relsPart, doc.toString())

    const layoutName = this.layouts().find((item) => item.part === target)?.name ?? null
    return {
      type: 'set_slide_layout',
      slide: index,
      part: slide.part,
      relationship_id: relId,
      from: previous,
      to: target,
      layout_name: layoutName,
      slide_count: this.info.slides.length
    }
  }

  /**
   * 读回每张幻灯片当前使用的版式（写入后用它验证）。
   * @returns {object[]} 每页的版式信息。
   */
  slideLayouts() {
    const byPart = new Map(this.layouts().map((item) => [item.part, item]))
    return this.info.slides.map((slide) => {
      const part = findLayout(this.#pkg, slide.part)
      return {
        index: slide.index,
        part: slide.part,
        layout_part: part,
        layout_index: part ? (byPart.get(part)?.index ?? null) : null,
        layout_name: part ? (byPart.get(part)?.name ?? null) : null
      }
    })
  }

  /**
   * 深拷贝一个图表部件（含它自己的关系与嵌入工作簿），返回新的幻灯片相对目标路径。
   *
   * 为什么要拷贝而不是共用：两页引用同一个 `chartN.xml` 时 PowerPoint 会**拒绝打开**整个文件
   * （它自己复制含图表的幻灯片时也会新建一份图表部件）。嵌入工作簿同理复制一份，
   * 免得两页的「编辑数据」互相串。
   *
   * @param {string} target - 幻灯片关系里的原始目标（如 `../charts/chart1.xml`）。
   * @returns {string} 新的相对目标（如 `../charts/chart2.xml`）。
   */
  #cloneChartPart(target) {
    const sourcePart = normalizePart('ppt/slides', target)
    if (!this.#pkg.has(sourcePart)) throw new OfficeError('CORRUPTED_DOCUMENT', `图表部件 ${sourcePart} 不存在。`)
    const chartIndex = this.#pkg.names().filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n)).length + 1
    const newPart = `ppt/charts/chart${chartIndex}.xml`
    this.#pkg.write(newPart, this.#pkg.readText(sourcePart))
    ensureContentTypeOverride(this.#pkg, `/${newPart}`, CT_CHART)

    const sourceRels = `ppt/charts/_rels/${sourcePart.split('/').pop()}.rels`
    if (this.#pkg.has(sourceRels)) {
      const relsDoc = XmlDoc.parse(this.#pkg.readText(sourceRels))
      for (const rel of findAll(relsDoc.root, 'Relationship')) {
        const relTarget = attr(rel, 'Target') ?? ''
        const relType = attr(rel, 'Type') ?? ''
        const embedded = normalizePart('ppt/charts', relTarget)
        if (!/^ppt\/embeddings\//.test(embedded) || !this.#pkg.has(embedded)) continue
        const embedIndex = this.#pkg.names().filter((n) => /^ppt\/embeddings\/.+\.xlsx$/.test(n)).length + 1
        const extension = embedded.split('.').pop()
        const newEmbed = `ppt/embeddings/Microsoft_Excel_Worksheet${embedIndex}.${extension}`
        this.#pkg.write(newEmbed, this.#pkg.read(embedded))
        relsDoc.patch(
          rel.start,
          rel.end,
          `<Relationship Id="${attr(rel, 'Id')}" Type="${relType}" Target="../embeddings/${newEmbed.split('/').pop()}"/>`
        )
      }
      this.#pkg.write(`ppt/charts/_rels/${newPart.split('/').pop()}.rels`, relsDoc.toString())
    }
    return `../charts/${newPart.split('/').pop()}`
  }

  /**
   * 往一张幻灯片插入原生图表（DrawingML 图表部件）。
   *
   * 需要写齐几处，缺一处 PowerPoint 就报「需要修复」：
   *   1. `ppt/charts/chartN.xml` 图表本体（与 XLSX 共用同一套 `c:chartSpace` 构造器）；
   *   2. `ppt/charts/_rels/chartN.xml.rels` 里的 `externalData` 关系（见下）；
   *   3. 幻灯片的 `_rels` 里一条 `…/chart` 关系；
   *   4. `<p:spTree>` 里引用该关系的 `<p:graphicFrame>`（`a:graphicData` 的 uri 必须是
   *      `…/drawingml/2006/chart`，写错 PowerPoint 会当成未知图形而丢弃）；
   *   5. `[Content_Types].xml` 里图表部件的 Override。
   *
   * 数据用**字面量缓存**（`c:strLit` / `c:numLit`）而不是单元格引用：PPTX 里没有必然存在的
   * 工作表可引用，字面量让阅读器不必重算就能作图。同时会**嵌入一个由本插件自己生成的 xlsx
   * 工作簿**并接上 `c:externalData` —— 没有它 PowerPoint 虽然能画图，但「编辑数据」是灰的。
   *
   * @param {object} args - 参数。
   * @param {number} args.index - 幻灯片下标（0 基）。
   * @param {string} [args.type] - `column`（默认）/ `bar` / `line` / `pie`。
   * @param {string} [args.title] - 图表标题。
   * @param {string[]} [args.categories] - 分类标签。
   * @param {object[]} args.series - `[{name, values: number[]}]`。
   * @param {number} [args.leftPx] - 左边距（像素）；省略则水平居中。
   * @param {number} [args.topPx] - 上边距（像素）；省略则垂直居中。
   * @param {number} [args.widthPx] - 宽（像素），默认 480。
   * @param {number} [args.heightPx] - 高（像素），默认 320。
   * @returns {object} 变更信息。
   */
  addSlideChart({ index, type = 'column', title = null, categories = [], series, leftPx, topPx, widthPx = 480, heightPx = 320 }) {
    const CHART_TYPES = { column: 'bar', bar: 'bar', line: 'line', pie: 'pie' }
    if (!CHART_TYPES[type]) {
      throw new OfficeError('INVALID_REQUEST', `不支持的图表类型：${type}（可用 column / bar / line / pie）。`)
    }
    if (!Array.isArray(series) || series.length === 0) {
      throw new OfficeError('INVALID_REQUEST', 'series 不能为空，且每个系列需要 {name, values}。')
    }
    if (type === 'pie' && series.length > 1) {
      throw new OfficeError('INVALID_REQUEST', '饼图只支持一个数据系列。')
    }
    for (const item of series) {
      if (!Array.isArray(item?.values) || item.values.length === 0) {
        throw new OfficeError('INVALID_REQUEST', '每个系列都需要非空的 values 数组。')
      }
      if (item.values.some((v) => !Number.isFinite(v))) {
        throw new OfficeError('INVALID_REQUEST', 'values 必须是数字数组。')
      }
    }
    const labels = Array.isArray(categories) && categories.length > 0 ? categories.map(String) : series[0].values.map((_, i) => String(i + 1))
    const { doc, slide } = this.#slideContext(index)
    const spTree = find(doc.root, 'spTree')
    if (!spTree) throw new OfficeError('CORRUPTED_DOCUMENT', `${slide.part} 缺少 spTree。`)

    // 1) 图表部件
    const chartIndex = this.#pkg.names().filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n)).length + 1
    const chartPart = `ppt/charts/chart${chartIndex}.xml`
    const lastRow = labels.length + 1
    const seriesXml = series
      .map((item, i) => {
        const col = indexToCol(i + 2)
        return buildLiteralSeriesXml({
          index: i,
          name: item.name ?? null,
          nameRef: `Sheet1!$${col}$1`,
          categories: labels,
          categoryRef: `Sheet1!$A$2:$A$${lastRow}`,
          values: item.values,
          valuesRef: `Sheet1!$${col}$2:$${col}$${lastRow}`
        })
      })
      .join('')
    const chartXml = buildChartSpaceXml({
      kind: CHART_TYPES[type],
      type,
      seriesXml,
      title,
      seriesCount: series.length
    })
    // externalData 的位置卡得很死：CT_ChartSpace 的顺序是
    //   date1904? lang? roundedCorners? style? clrMapOvr? pivotSource? protection? **chart** spPr? txPr?
    //   **externalData**? printSettings? userShapes? extLst?
    // 也就是必须排在 `c:chart` **与 `c:spPr` 之后**、`</c:chartSpace>` 之前。
    // 放到 chart 之前或 spPr 之前，PowerPoint 直接「无法打开该文件」（不是提示修复，是拒绝打开）。
    const withExternal = chartXml.replace(
      '</c:chartSpace>',
      '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>'
    )
    this.#pkg.write(chartPart, withExternal)
    ensureContentTypeOverride(this.#pkg, `/${chartPart}`, CT_CHART)

    // 2) 嵌入工作簿（让 PowerPoint 的「编辑数据」可用）：用本插件自己的 XLSX 写入器生成
    const workbook = buildChartWorkbook({ labels, series })
    const embedPart = `ppt/embeddings/Microsoft_Excel_Worksheet${chartIndex}.xlsx`
    this.#pkg.write(embedPart, workbook)
    ensureContentTypeDefault(this.#pkg, 'xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    const chartRid = addRelationshipTo(
      this.#pkg,
      `ppt/charts/_rels/chart${chartIndex}.xml.rels`,
      `${NS_R}/package`,
      `../embeddings/Microsoft_Excel_Worksheet${chartIndex}.xlsx`
    )
    // 关系 id 固定写 rId1：上面 externalData 里引用的就是它
    if (chartRid !== 'rId1') {
      const relsPath = `ppt/charts/_rels/chart${chartIndex}.xml.rels`
      this.#pkg.write(relsPath, this.#pkg.readText(relsPath).replace(`Id="${chartRid}"`, 'Id="rId1"'))
    }

    // 3) 幻灯片关系 → 图表部件
    const slideRid = addRelationshipTo(this.#pkg, slideRelPath(slide.part), `${NS_R}/chart`, `../charts/chart${chartIndex}.xml`)

    // 4) graphicFrame
    const slideSize = this.#slideSize()
    const cx = Math.round(widthPx * EMU_PER_PX)
    const cy = Math.round(heightPx * EMU_PER_PX)
    const id = this.#nextShapeId(doc)
    const offX = leftPx === undefined ? Math.max(0, Math.round((slideSize.cx - cx) / 2)) : Math.round(leftPx * EMU_PER_PX)
    const offY = topPx === undefined ? Math.max(0, Math.round((slideSize.cy - cy) / 2)) : Math.round(topPx * EMU_PER_PX)
    doc.appendChild(spTree, buildChartFrameXml({ id, offX, offY, cx, cy, rid: slideRid, name: `图表 ${id}` }))
    this.#pkg.write(slide.part, doc.toString())
    doc.rescan()

    return {
      type: 'add_slide_chart',
      slide: index,
      chart_part: chartPart,
      embedded_workbook: embedPart,
      chart_type: type,
      title,
      series_count: series.length,
      categories: labels.length,
      shape_id: id,
      relationship_id: slideRid,
      offset_emu: { x: offX, y: offY },
      size_emu: { cx, cy },
      note: '图表数据用字面量缓存写入，并另存一份嵌入工作簿供 PowerPoint「编辑数据」使用。'
    }
  }

  /**
   * 取幻灯片尺寸（EMU）。
   * @returns {{cx: number, cy: number}} 宽高。
   */
  #slideSize() {
    const sldSz = find(this.info.presentationDoc.root, 'sldSz')
    const cx = Number(attr(sldSz ?? { attrs: new Map() }, 'cx') ?? 0) || 12192000
    const cy = Number(attr(sldSz ?? { attrs: new Map() }, 'cy') ?? 0) || 6858000
    return { cx, cy }
  }

  /**
   * 列出全部主题（名称 + 被哪个母版引用）。
   *
   * 主题名取自 `<a:theme name="…">` —— 这是 PowerPoint「设计」里显示的名字。
   *
   * @returns {object[]} 主题列表。
   */
  themes() {
    return this.info.themes.map((part, index) => {
      const doc = XmlDoc.parse(this.#pkg.readText(part))
      // 同上：根元素要从 `#document` 容器的 element 子节点取，否则 `name` 永远是 null
      const root = doc.root.children.find((c) => c.type === 'element')
      const rawName = attr(root ?? { attrs: new Map() }, 'name') ?? null
      return {
        index,
        part,
        name: cleanOfficeName(rawName),
        raw_name: rawName,
        used_by_masters: this.info.masters
          .map((m, i) => ({ m, i }))
          .filter(({ m }) => findTheme(this.#pkg, m) === part)
          .map(({ i }) => i)
      }
    })
  }

  /**
   * 切换母版使用的主题。
   *
   * 主题关联**只由关系决定**：`ppt/slideMasters/_rels/slideMasterN.xml.rels` 里那条
   * `…/theme` 关系的 `Target` 指向哪个主题部件，PowerPoint 就按哪套配色/字体渲染该母版下的
   * 全部版式与幻灯片。所以这里改的是关系目标，不动任何版式、母版或幻灯片 XML。
   *
   * **验证手段**：PowerPoint COM 能读出 `SlideMaster.Design.Name` 与
   * `Theme.ThemeColorScheme.Colors(n).RGB`，所以「换主题确实生效」可以被客观验证 ——
   * 实测把源主题的 accent1 改成红色后切换，PowerPoint 读回的 `ACCENT1_RGB` 从默认值变成 `255`（红）。
   *
   * @param {object} args - 参数。
   * @param {number} [args.master] - 母版下标（0 基），默认 0。
   * @param {number} [args.theme] - 目标主题下标（0 基，见 `themes()`）。
   * @param {string} [args.name] - 目标主题名（与 `theme` 二选一，重名时取第一个）。
   * @returns {object} 变更信息。
   */
  setTheme({ master = 0, theme, name }) {
    const masterPart = this.info.masters[master]
    if (!masterPart) {
      throw new OfficeError('FILE_NOT_FOUND', `不存在第 ${master} 个母版（共 ${this.info.masters.length} 个）。`, {
        master_count: this.info.masters.length
      })
    }
    if (theme === undefined && name === undefined) {
      throw new OfficeError('INVALID_REQUEST', '必须提供 theme（下标）或 name（主题名）之一。')
    }
    let source
    if (theme !== undefined) {
      source = this.info.themes[theme]
      if (!source) {
        throw new OfficeError('INVALID_REQUEST', `主题下标 ${theme} 不存在（共 ${this.info.themes.length} 个）。`, {
          theme_count: this.info.themes.length
        })
      }
    } else {
      const hit = this.themes().find((item) => item.name?.toLowerCase() === (cleanOfficeName(name)?.toLowerCase() ?? ''))
      if (!hit) {
        throw new OfficeError('INVALID_REQUEST', `找不到名为「${name}」的主题。`, {
          theme_names: this.themes().map((t) => t.name)
        })
      }
      source = hit.part
    }

    const relsPart = slideRelPath(masterPart)
    if (!this.#pkg.has(relsPart)) throw new OfficeError('CORRUPTED_DOCUMENT', `${masterPart} 缺少关系部件。`)
    const previous = findTheme(this.#pkg, masterPart)
    if (previous === source) {
      // 返回值形状与正常路径保持一致（缺字段会让调用方踩 `undefined`，这一点吃过亏）
      return {
        type: 'set_theme',
        master,
        part: masterPart,
        from: previous,
        to: source,
        source_theme: source,
        theme_name: this.themes().find((item) => item.part === source)?.name ?? null,
        reused_part: false,
        changed: false,
        note: '母版已经用的是这套主题，未做改动。'
      }
    }

    // 主题部件不能直接「指过去」：`themeN.xml` 往往已被别的部件引用（典型是备注母版），
    // 直接把幻灯片母版指过去会造出「双重引用 + 另一个主题变孤儿」，
    // 实测 **PowerPoint 会直接拒绝打开整个文件**（对照组：只换主题内容则正常）。
    // PowerPoint 自己的做法是**新建一个主题部件**，这里照做：
    //   - 当前主题只被这个母版引用 → 就地改它的内容（来回切换不会无限堆积部件）；
    //   - 还被别处引用（如 presentation.xml.rels / 备注母版）→ 新建 `themeN+1.xml` 再挂上去。
    const referrers = findReferrers(this.#pkg, previous)
    let target
    let reused = false
    if (referrers.length <= 1) {
      target = previous
      this.#pkg.write(target, this.#pkg.readText(source))
      reused = true
    } else {
      const index = this.#pkg.names().filter((n) => /^ppt\/theme\/theme\d+\.xml$/.test(n)).length + 1
      target = `ppt/theme/theme${index}.xml`
      this.#pkg.write(target, this.#pkg.readText(source))
      ensureContentTypeOverride(this.#pkg, `/${target}`, CT_THEME)
    }

    const doc = XmlDoc.parse(this.#pkg.readText(relsPart))
    const rel = findAll(doc.root, 'Relationship').find((r) => (attr(r, 'Type') ?? '').endsWith('/theme'))
    if (!rel) throw new OfficeError('CORRUPTED_DOCUMENT', `${relsPart} 里没有 theme 关系。`)
    const relId = attr(rel, 'Id')
    // 目标是相对 `ppt/slideMasters/` 的路径：主题在 `ppt/theme/`，因此必须带 `../`
    doc.patch(rel.start, rel.end, `<Relationship Id="${relId}" Type="${NS_R}/theme" Target="../theme/${target.split('/').pop()}"/>`)
    this.#pkg.write(relsPart, doc.toString())

    return {
      type: 'set_theme',
      master,
      part: masterPart,
      relationship_id: relId,
      from: previous,
      to: target,
      source_theme: source,
      theme_name: this.themes().find((item) => item.part === source)?.name ?? null,
      reused_part: reused,
      changed: true,
      note: '克隆主题部件后再挂载（直接指向别处已在用的主题部件会让 PowerPoint 拒绝打开）；配色/字体的实际效果请在 PowerPoint 里确认。'
    }
  }

  /**
   * 读回每个母版当前使用的主题（写入后用它验证）。
   * @returns {object[]} 每个母版的主题信息。
   */
  masterThemes() {
    const byPart = new Map(this.themes().map((item) => [item.part, item]))
    return this.info.masters.map((part, index) => {
      const themePart = findTheme(this.#pkg, part)
      return {
        index,
        part,
        theme_part: themePart,
        theme_index: themePart ? (byPart.get(themePart)?.index ?? null) : null,
        theme_name: themePart ? (byPart.get(themePart)?.name ?? null) : null
      }
    })
  }

  /**
   * 写回 presentation.xml。
   * @returns {void}
   */
  #commitPresentation() {
    this.#pkg.write('ppt/presentation.xml', this.info.presentationDoc.toString())
    this.info.presentationDoc.rescan()
  }

  /**
   * 序列化演示文稿。
   * @returns {Buffer} .pptx 字节。
   */
  save() {
    return this.#pkg.toBuffer()
  }

  /** @returns {ZipPackage} 底层 OPC 包。 */
  get pkg() {
    return this.#pkg
  }
}

/** PPTX 中必须渲染才能判定、本阶段不做的检查项。 */
const PPTX_NOT_CHECKED = Object.freeze([
  { name: '文本溢出与裁剪', reason: '需要按字体度量排版后才能判定' },
  { name: '元素重叠', reason: '需要渲染后才能判定' },
  { name: '字体替换', reason: '需要目标机器字体环境；可用 office_convert_document 的 missing_fonts 判定' },
  { name: '空白幻灯片', reason: '需要渲染后才能判定（结构上只能报告无文本的形状数）' }
])

/**
 * 提取表格（`<a:tbl>`）的行列文本。
 *
 * 行列只数直接子元素：嵌套表格的 `<a:tr>` 不会被误当成外层表格的行。
 *
 * @param {XmlDoc} doc - 文档对象。
 * @param {object} node - graphicFrame 节点。
 * @returns {{rows: string[][], row_count: number, column_count: number}} 表格内容。
 */
function tableInfo(doc, node) {
  const tbl = find(node, 'tbl')
  if (!tbl) return { rows: [], row_count: 0, column_count: 0 }
  const rows = (tbl.children ?? [])
    .filter((n) => n.type === 'element' && localName(n) === 'tr')
    .map((tr) =>
      (tr.children ?? [])
        .filter((n) => n.type === 'element' && localName(n) === 'tc')
        .map((tc) =>
          findAll(tc, 'p')
            .map((p) => paragraphText(doc, p))
            .join('\n')
        )
    )
  return {
    rows,
    row_count: rows.length,
    column_count: rows.reduce((m, r) => Math.max(m, r.length), 0)
  }
}

/**
 * 取形状的显示名（`<p:cNvPr name="...">`）。
 * @param {object} node - 形状节点。
 * @returns {string|null} 名称。
 */
function shapeName(node) {
  const cNvPr = find(node, 'cNvPr')
  return cNvPr ? (attr(cNvPr, 'name') ?? null) : null
}

/**
 * 由 graphicData 的 uri 细分 graphicFrame 的实际类型。
 * @param {object} node - `<p:graphicFrame>` 节点。
 * @returns {'table'|'chart'|'smartart'|'graphicFrame'} 类型。
 */
function graphicFrameKind(node) {
  const uri = attr(find(node, 'graphicData') ?? { attrs: new Map() }, 'uri') ?? ''
  if (uri.includes('/table')) return 'table'
  if (uri.includes('/chart')) return 'chart'
  if (uri.includes('/diagram')) return 'smartart'
  return 'graphicFrame'
}

/**
 * 生成一段 DrawingML 文本的 run 序列。
 * @param {string} text - 纯文本。
 * @returns {string} run 序列 XML；空文本返回空串。
 */
function textRunsXml(text) {
  const value = String(text ?? '')
  if (value === '') return ''
  return `<a:r><a:rPr lang="zh-CN" dirty="0"/><a:t>${escapeXmlText(value)}</a:t></a:r>`
}

/**
 * 生成用于变更记录的短预览（不把正文写进日志）。
 * @param {string} text - 原文。
 * @param {number} [max] - 最大长度。
 * @returns {string} 预览文本。
 */
function preview(text, max = 30) {
  const value = String(text ?? '').replace(/\n/g, '⏎')
  return value.length > max ? `${value.slice(0, max)}…` : value
}

/**
 * 精确读取元素的 `r:id` 属性。
 *
 * OOXML 里同一个元素常同时带无前缀的 `id` 和命名空间化的 `r:id`，
 * 二者语义完全不同。按本地名匹配会取错，必须按完整属性名取。
 *
 * @param {object} node - 元素节点。
 * @returns {string|undefined} 关系 ID。
 */
function relationshipId(node) {
  return node.attrs?.get('r:id')?.value ?? node.attrs?.get('rel:id')?.value
}

/**
 * 取节点本地名。
 * @param {object} node - 元素节点。
 * @returns {string} 本地名。
 */
function localName(node) {
  return node.name.includes(':') ? node.name.slice(node.name.indexOf(':') + 1) : node.name
}

/**
 * 取路径扩展名（小写）。
 * @param {string} path - 路径。
 * @returns {string} 扩展名。
 */
function extOf(path) {
  const i = path.lastIndexOf('.')
  return i === -1 ? '' : path.slice(i + 1).toLowerCase()
}

/**
 * 提取一个 `<a:p>` 段落的文本。
 * @param {XmlDoc} doc - 文档对象。
 * @param {object} node - 段落或任意容器节点。
 * @returns {string} 文本。
 */
function paragraphText(doc, node) {
  let out = ''
  const walk = (current) => {
    for (const child of current.children ?? []) {
      if (child.type !== 'element') continue
      const name = localName(child)
      if (name === 't') out += nodeText(doc, child)
      else if (name === 'br') out += '\n'
      else if (name === 'tab') out += '\t'
      else if (name === 'pPr' || name === 'rPr' || name === 'endParaRPr' || name === 'defRPr') continue
      else walk(child)
    }
  }
  walk(node)
  return out
}

/**
 * 汇总一个形状的信息。
 * @param {XmlDoc} doc - 文档对象。
 * @param {object} sp - `<p:sp>` 节点。
 * @param {number} depth - 组合嵌套深度。
 * @returns {object} 形状信息。
 */
function shapeInfo(doc, sp, depth) {
  const cNvPr = find(sp, 'cNvPr')
  const ph = find(sp, 'ph')
  const txBody = find(sp, 'txBody')
  const paragraphs = txBody ? findAll(txBody, 'p').map((p) => paragraphText(doc, p)) : []
  const phType = ph ? (attr(ph, 'type') ?? 'body') : null
  return {
    type: 'shape',
    name: cNvPr ? (attr(cNvPr, 'name') ?? null) : null,
    placeholder: phType,
    placeholder_label: phType ? (PLACEHOLDER_LABELS[phType] ?? phType) : null,
    text: paragraphs.join('\n'),
    paragraph_count: paragraphs.length,
    depth
  }
}

/**
 * PowerPoint 内置表格样式「中档样式 2 - 强调 1」。
 *
 * 不指定 tableStyleId 时表格会渲染成完全无边框的裸文本，
 * 与用户说「插入一个表格」时的预期不符。
 */
const DEFAULT_TABLE_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'

/**
 * 构造一个表格的 `<p:graphicFrame>`。
 * @param {object} args - `{id, offX, offY, cx, cy, rows, columnCount, firstRowHeader}`。
 * @returns {string} graphicFrame XML。
 */
function buildTableGraphicFrameXml({ id, offX, offY, cx, cy, rows, columnCount, firstRowHeader }) {
  const colWidth = Math.floor(cx / columnCount)
  const rowHeight = Math.floor(cy / rows.length)
  const grid = Array.from({ length: columnCount }, () => `<a:gridCol w="${colWidth}"/>`).join('')
  const body = rows
    .map((row) => {
      const cells = Array.from({ length: columnCount }, (_, c) => {
        const text = row[c] === undefined || row[c] === null ? '' : String(row[c])
        const runs = text
          .split('\n')
          .map((line) => `<a:p><a:r><a:rPr lang="zh-CN" dirty="0"/><a:t>${escapeXmlText(line)}</a:t></a:r></a:p>`)
          .join('')
        return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${runs}</a:txBody><a:tcPr/></a:tc>`
      }).join('')
      return `<a:tr h="${rowHeight}">${cells}</a:tr>`
    })
    .join('')

  const tblPr = `<a:tblPr${firstRowHeader ? ' firstRow="1"' : ''} bandRow="1"><a:tableStyleId>${DEFAULT_TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>`

  return (
    '<p:graphicFrame>' +
    `<p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${offX}" y="${offY}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
    `<a:tbl>${tblPr}<a:tblGrid>${grid}</a:tblGrid>${body}</a:tbl>` +
    '</a:graphicData></a:graphic></p:graphicFrame>'
  )
}

/** 关系类型：图片。 */
const REL_IMAGE = `${NS_R}/image`

/**
 * 构造一个引用图片的 `<p:pic>`。
 * @param {object} args - `{rid, id, offX, offY, cx, cy, altText}`。
 * @returns {string} `<p:pic>` XML。
 */
function buildPicXml({ rid, id, offX, offY, cx, cy, altText }) {
  const descr = altText ? ` descr="${escapeXmlAttr(altText)}"` : ''
  const name = `Picture ${id}`
  return (
    '<p:pic>' +
    `<p:nvPicPr><p:cNvPr id="${id}" name="${name}"${descr}/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${offX}" y="${offY}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    '</p:pic>'
  )
}

/** 新建幻灯片的最小合法 XML（空的形状树 + 继承母版配色）。 */
const EMPTY_SLIDE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sld xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`

/* 内容类型与关系操作已下沉到 ooxml.js（ensureContentTypeOverride /
 * removeContentTypeOverride / addRelationshipTo / removeRelationshipFrom），
 * 与 XLSX / DOCX 适配器共用同一份实现。 */

/**
 * 读取关系文件。
 * @param {ZipPackage} pkg - OPC 包。
 * @param {string} part - 关系部件路径。
 * @returns {Map<string, string>} `rId → Target`。
 */
function readRels(pkg, part) {
  const map = new Map()
  if (!pkg.has(part)) return map
  const doc = XmlDoc.parse(pkg.readText(part))
  for (const rel of findAll(doc.root, 'Relationship')) {
    const id = attr(rel, 'Id')
    const target = attr(rel, 'Target')
    if (id && target) map.set(id, target)
  }
  return map
}

/**
 * 由幻灯片部件路径推出其关系部件路径。
 * @param {string} slidePart - 形如 `ppt/slides/slide1.xml`。
 * @returns {string} 形如 `ppt/slides/_rels/slide1.xml.rels`。
 */
function slideRelPath(slidePart) {
  const dir = slidePart.slice(0, slidePart.lastIndexOf('/'))
  const file = slidePart.slice(slidePart.lastIndexOf('/') + 1)
  return `${dir}/_rels/${file}.rels`
}

/**
 * 把关系目标解析为包内绝对路径。
 * @param {string} baseDir - 宿主部件所在目录。
 * @param {string} target - 关系目标。
 * @returns {string} 归一化路径。
 */
function normalizePart(baseDir, target) {
  if (target.startsWith('/')) return target.slice(1)
  const stack = []
  for (const seg of `${baseDir}/${target}`.split('/')) {
    if (seg === '' || seg === '.') continue
    if (seg === '..') stack.pop()
    else stack.push(seg)
  }
  return stack.join('/')
}

/**
 * 判断幻灯片是否含真实动画。
 *
 * `<p:timing>` 节点即使没有动画也会存在（带空 tnLst），因此必须看里面
 * 是否真有 `p:anim*` / `p:par` 节点，否则会把所有幻灯片都误报为有动画。
 *
 * @param {ZipPackage} pkg - OPC 包。
 * @param {string} part - 幻灯片部件路径。
 * @returns {boolean} 是否含动画。
 */
function hasAnimation(pkg, part) {
  if (!pkg.has(part)) return false
  const raw = pkg.readText(part)
  if (!/<p:timing>/.test(raw)) return false
  return /<p:(anim|animClr|animEffect|animMotion|animRot|animScale|cmd)\b/.test(raw)
}

/**
 * 判断幻灯片是否含切换效果。
 * @param {XmlDoc} doc - 幻灯片文档。
 * @returns {boolean} 是否含切换。
 */
function hasTransition(doc) {
  return find(doc.root, 'transition') !== undefined
}

/**
 * 构造引用图表部件的 `<p:graphicFrame>`。
 *
 * `a:graphicData` 的 `uri` 必须是 `…/drawingml/2006/chart` —— 写错时 PowerPoint
 * 不报错也不显示，图形会静默消失（这种「静默丢弃」最难查）。
 *
 * @param {object} args - 参数。
 * @param {number} args.id - 形状 id。
 * @param {number} args.offX - 左边距（EMU）。
 * @param {number} args.offY - 上边距（EMU）。
 * @param {number} args.cx - 宽（EMU）。
 * @param {number} args.cy - 高（EMU）。
 * @param {string} args.rid - 幻灯片关系 id。
 * @param {string} args.name - 形状名。
 * @returns {string} `<p:graphicFrame>` XML。
 */
function buildChartFrameXml({ id, offX, offY, cx, cy, rid, name }) {
  return (
    '<p:graphicFrame>' +
    `<p:nvGraphicFramePr><p:cNvPr id="${id}" name="${escapeXmlAttr(name)}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${offX}" y="${offY}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>` +
    '<a:graphic>' +
    `<a:graphicData uri="${NS_CHART}"><c:chart xmlns:c="${NS_CHART}" xmlns:r="${NS_R}" r:id="${rid}"/></a:graphicData>` +
    '</a:graphic></p:graphicFrame>'
  )
}

/**
 * 生成嵌入给 PowerPoint「编辑数据」用的 xlsx 工作簿。
 *
 * 用**本插件自己的 XLSX 写入器**生成，而不是塞一份静态字节：数据与图表一致，
 * 且不需要在仓库里放二进制模板。第一列是分类标签，之后每列一个系列。
 *
 * @param {object} args - 参数。
 * @param {string[]} args.labels - 分类标签。
 * @param {object[]} args.series - `[{name, values}]`。
 * @returns {Buffer} .xlsx 字节。
 */
function buildChartWorkbook({ labels, series }) {
  const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { ...DEFAULT_ZIP_LIMITS })
  const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
  const colName = (index) => {
    let n = index
    let name = ''
    while (n > 0) {
      const rest = (n - 1) % 26
      name = String.fromCharCode(65 + rest) + name
      n = Math.floor((n - 1) / 26)
    }
    return name
  }
  const cellXml = (ref, value) =>
    typeof value === 'number'
      ? `<c r="${ref}"><v>${value}</v></c>`
      : `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXmlText(String(value))}</t></is></c>`

  const rows = [[null, ...series.map((s, i) => s.name ?? `系列${i + 1}`)]]
  labels.forEach((label, rowIndex) => {
    rows.push([label, ...series.map((s) => s.values[rowIndex] ?? null)])
  })
  const sheetData = rows
    .map((row, r) => {
      const cells = row
        .map((value, c) => (value === null || value === undefined ? '' : cellXml(colName(c + 1) + (r + 1), value)))
        .join('')
      return `<row r="${r + 1}">${cells}</row>`
    })
    .join('')

  pkg.write(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>'
  )
  pkg.write(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`
  )
  pkg.write(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">` +
      '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
  )
  pkg.write(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${NS_R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`
  )
  pkg.write(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="${NS_MAIN}"><sheetData>${sheetData}</sheetData></worksheet>`
  )
  return pkg.toBuffer()
}

/**
 * 找出所有引用了某个部件的 `.rels` 文件（用来判断「这个主题部件还有别人在用吗」）。
 *
 * 主题部件不能随便被第二个宿主引用：实测把幻灯片母版直接指向备注母版在用的主题，
 * PowerPoint 会拒绝打开整个文件；反过来，把只被一个母版引用的主题就地改内容则是安全的。
 *
 * @param {ZipPackage} pkg - OPC 包。
 * @param {string} targetPart - 目标部件路径。
 * @returns {string[]} 引用它的 `.rels` 部件路径列表。
 */
function findReferrers(pkg, targetPart) {
  const out = []
  for (const part of pkg.names()) {
    if (!part.endsWith('.rels')) continue
    const baseDir = part.includes('/_rels/') ? part.slice(0, part.indexOf('/_rels/')) : ''
    for (const target of readRels(pkg, part).values()) {
      if (normalizePart(baseDir, target) === targetPart) {
        out.push(part)
        break
      }
    }
  }
  return out
}

/**
 * 规整 Office 里的名称（版式名 / 主题名）。
 *
 * PowerPoint 简体中文版会在主题名后写入**零宽空格**（实测 `Office 主题` 后面跟两个 `U+200B`），
 * 版式名里也可能有。不清理的话「按名字匹配」永远匹配不上，而错误信息里两次打印看起来一模一样
 * —— 这类不可见字符造成的失败极难排查，所以在入口统一清掉，并把原始值一并返回备查。
 *
 * @param {string|null|undefined} value - 原始名称。
 * @returns {string|null} 清掉零宽字符并 trim 后的名称；输入为空时返回 null。
 */
function cleanOfficeName(value) {
  if (typeof value !== 'string') return null
  const cleaned = value.replace(/[\u200B-\u200D\uFEFF]/g, '').trim()
  return cleaned === '' ? null : cleaned
}

/**
 * 由母版关系推出它使用的主题部件路径。
 * @param {ZipPackage} pkg - OPC 包。
 * @param {string} masterPart - 母版部件路径。
 * @returns {string|null} 主题部件路径。
 */
function findTheme(pkg, masterPart) {
  const rels = readRels(pkg, slideRelPath(masterPart))
  for (const target of rels.values()) {
    if (/theme\/theme\d+\.xml$/.test(target)) return normalizePart('ppt/slideMasters', target)
  }
  return null
}

/**
 * 由幻灯片关系推出它使用的版式名。
 * @param {ZipPackage} pkg - OPC 包。
 * @param {string} slidePart - 幻灯片部件路径。
 * @returns {string|null} 版式部件路径。
 */
function findLayout(pkg, slidePart) {
  const rels = readRels(pkg, slideRelPath(slidePart))
  for (const target of rels.values()) {
    if (/slideLayout\d+\.xml$/.test(target)) return normalizePart('ppt/slides', target)
  }
  return null
}

/**
 * 由幻灯片尺寸判定宽高比名称。
 * @param {number} cx - 宽（EMU）。
 * @param {number} cy - 高（EMU）。
 * @returns {string} 比例名。
 */
function aspectName(cx, cy) {
  const ratio = cx / cy
  if (Math.abs(ratio - 16 / 9) < 0.02) return '16:9'
  if (Math.abs(ratio - 4 / 3) < 0.02) return '4:3'
  if (Math.abs(ratio - 16 / 10) < 0.02) return '16:10'
  return `${Math.round(ratio * 100) / 100}:1`
}

export { CT_PRESENTATION, CT_SLIDE, CT_SLIDE_MASTER, CT_SLIDE_LAYOUT, CT_NOTES_SLIDE, CT_THEME, NS_P, NS_A }
