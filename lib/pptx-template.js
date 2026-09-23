/**
 * 从零生成一份最小但**完整合法**的 .pptx（空白演示文稿模板）。
 *
 * 为什么要手写而不是塞一个二进制模板进仓库：模板二进制体积大、来源与许可都不清晰，
 * 而且改一个字号就要重新做一份。这里把必需的部件全部程序化生成，配色/字体也是我们自己的，
 * 不复制任何 Office 文件里的内容。
 *
 * 一份能被 PowerPoint 正常打开的演示文稿至少要写齐这些部件（少一个都会报「需要修复」或拒绝打开）：
 *   1. `[Content_Types].xml` —— 每个部件的 Override 与扩展名 Default
 *   2. `_rels/.rels` —— 指向 presentation.xml（以及可选的 docProps）
 *   3. `ppt/presentation.xml` —— sldMasterIdLst / sldIdLst / sldSz / notesSz
 *   4. `ppt/_rels/presentation.xml.rels` —— 母版、幻灯片、主题的关系
 *   5. `ppt/slideMasters/slideMaster1.xml` —— **必须带 `p:clrMap`**，否则 PowerPoint 报修复
 *   6. `ppt/slideMasters/_rels/slideMaster1.xml.rels` —— 版式 + 主题
 *   7. `ppt/slideLayouts/slideLayout1.xml` + 它的 rels（指回母版）
 *   8. `ppt/theme/theme1.xml` —— clrScheme / fontScheme / **fmtScheme（fill+ln+effect+bgFill 各 3 条）**
 *   9. `ppt/slides/slide1.xml` + 它的 rels（指向版式）
 *  10. `docProps/core.xml` 与 `docProps/app.xml`（可选，但文件属性会好看很多）
 *
 * @module lib/pptx-template
 */
import { ZipPackage, ensureContentTypeOverride, DEFAULT_ZIP_LIMITS } from './ooxml.js'
import { OfficeError } from './errors.js'

const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const CT_PRESENTATION = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
const CT_SLIDE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
const CT_SLIDE_MASTER = 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
const CT_SLIDE_LAYOUT = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
const CT_THEME = 'application/vnd.openxmlformats-officedocument.theme+xml'
const CT_CORE = 'application/vnd.openxmlformats-package.core-properties+xml'
const CT_APP = 'application/vnd.openxmlformats-officedocument.extended-properties+xml'

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

/** 16:9 幻灯片尺寸（EMU）：13.333in × 7.5in，与 PowerPoint 的「宽屏」一致。 */
const SLIDE_W = 12192000
const SLIDE_H = 6858000

/** 版式：`title` = 标题 + 副标题；`blank` = 纯空白。 */
const LAYOUTS = Object.freeze({
  title: { name: '标题幻灯片', type: 'title', placeholders: ['ctrTitle', 'subTitle'] },
  blank: { name: '空白', type: 'blank', placeholders: [] }
})

/**
 * 生成一个形状（占位符）的 XML。
 *
 * 占位符必须写 `<p:ph type="…" idx="…"/>`，且 `p:spPr` 里要有 xfrm 与 prstGeom，
 * 否则 PowerPoint 会把它当成坏形状（打开时提示修复）。
 *
 * @param {object} args - 参数。
 * @param {number} args.id - 形状 id（同一页内唯一）。
 * @param {string} args.name - 形状名。
 * @param {string} args.placeholder - 占位符类型。
 * @param {number} [args.idx] - 占位符索引（副标题等需要）。
 * @param {number} args.x - 左边距（EMU）。
 * @param {number} args.y - 上边距（EMU）。
 * @param {number} args.cx - 宽（EMU）。
 * @param {number} args.cy - 高（EMU）。
 * @param {string} args.text - 文本（可为空）。
 * @param {number} args.fontSize - 字号（百分之一磅）。
 * @returns {string} `<p:sp>` XML。
 */
function placeholderXml({ id, name, placeholder, idx = null, x, y, cx, cy, text, fontSize }) {
  const ph = idx === null ? `<p:ph type="${placeholder}"/>` : `<p:ph type="${placeholder}" idx="${idx}"/>`
  const paragraphs = (text === '' ? [''] : String(text).split('\n'))
    .map((line) => `<a:p><a:r><a:rPr lang="zh-CN" sz="${fontSize}" dirty="0"/><a:t>${escapeXml(line)}</a:t></a:r></a:p>`)
    .join('')
  return (
    '<p:sp>' +
    `<p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
    `<p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`
  )
}

/**
 * 空的形状树（母版/版式/幻灯片都要有）。
 * @param {string} content - 形状 XML。
 * @returns {string} `<p:spTree>` XML。
 */
function spTreeXml(content) {
  return (
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    `${content}</p:spTree>`
  )
}

/**
 * 主题 XML（配色、字体、格式方案）。
 *
 * `fmtScheme` 的四组样式**每组必须正好 3 条**（fillStyleLst / lnStyleLst / effectStyleLst / bgFillStyleLst），
 * 少一条 PowerPoint 就报内容有问题 —— 这是手写主题最容易踩的坑。
 *
 * 配色与字体是**我们自己定的**（不复制 Office 主题文件的内容），名字里带插件名以便识别。
 *
 * @returns {string} `ppt/theme/theme1.xml` 内容。
 */
function themeXml() {
  const solid = (color) => `<a:solidFill><a:schemeClr val="${color}"/></a:solidFill>`
  // 渐变停靠点的颜色必须**直接**写在 `<a:gs>` 下（`<a:gs pos="0"><a:schemeClr …/></a:gs>`）；
  // 包一层 `<a:solidFill>` 看起来像但**不是合法的 CT_GradientStop 颜色选择**，
  // PowerPoint 会因此拒绝打开整个文件（实测：把 fmtScheme 换成 PowerPoint 自己写的就正常）。
  const grad = (a, b) =>
    `<a:gradFill rotWithShape="1"><a:gsLst>` +
    `<a:gs pos="0"><a:schemeClr val="${a}"/></a:gs><a:gs pos="100000"><a:schemeClr val="${b}"/></a:gs>` +
    `</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`
  const line = (w) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">${solid('phClr')}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`
  const effect = `<a:effectStyle><a:effectLst/></a:effectStyle>`
  return (
    XML_HEAD +
    `<a:theme xmlns:a="${NS_A}" name="office-plugin">` +
    '<a:themeElements>' +
    '<a:clrScheme name="office-plugin">' +
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
    '<a:dk2><a:srgbClr val="1F2A44"/></a:dk2>' +
    '<a:lt2><a:srgbClr val="EEF2F8"/></a:lt2>' +
    '<a:accent1><a:srgbClr val="2F6FEB"/></a:accent1>' +
    '<a:accent2><a:srgbClr val="7A5AF8"/></a:accent2>' +
    '<a:accent3><a:srgbClr val="12B886"/></a:accent3>' +
    '<a:accent4><a:srgbClr val="F59F00"/></a:accent4>' +
    '<a:accent5><a:srgbClr val="E8590C"/></a:accent5>' +
    '<a:accent6><a:srgbClr val="D6336C"/></a:accent6>' +
    '<a:hlink><a:srgbClr val="1C7ED6"/></a:hlink>' +
    '<a:folHlink><a:srgbClr val="7048E8"/></a:folHlink>' +
    '</a:clrScheme>' +
    '<a:fontScheme name="office-plugin">' +
    '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface="微软雅黑"/><a:cs typeface=""/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="微软雅黑"/><a:cs typeface=""/></a:minorFont>' +
    '</a:fontScheme>' +
    '<a:fmtScheme name="office-plugin">' +
    `<a:fillStyleLst>${solid('phClr')}${grad('phClr', 'tx1')}${grad('lt1', 'phClr')}</a:fillStyleLst>` +
    `<a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
    `<a:effectStyleLst>${effect}${effect}${effect}</a:effectStyleLst>` +
    `<a:bgFillStyleLst>${solid('phClr')}${solid('phClr')}${grad('phClr', 'dk1')}</a:bgFillStyleLst>` +
    '</a:fmtScheme>' +
    '</a:themeElements>' +
    '<a:objectDefaults/><a:extraClrSchemeLst/>' +
    '</a:theme>'
  )
}

/**
 * 母版 XML。
 *
 * `p:clrMap` 是**必需**元素（bg1/tx1/bg2/tx2/accent1-6/hlink/folHlink 十项缺一不可），
 * 缺了 PowerPoint 直接报「内容有问题」。
 *
 * @returns {string} `ppt/slideMasters/slideMaster1.xml` 内容。
 */
function slideMasterXml() {
  const body =
    placeholderXml({
      id: 2,
      name: '标题占位符',
      placeholder: 'title',
      x: 838200,
      y: 365125,
      cx: 10515600,
      cy: 1325563,
      text: '单击此处编辑母版标题样式',
      fontSize: 4400
    }) +
    placeholderXml({
      id: 3,
      name: '正文占位符',
      placeholder: 'body',
      idx: 1,
      x: 838200,
      y: 1825625,
      cx: 10515600,
      cy: 4351338,
      text: '',
      fontSize: 2800
    })
  return (
    XML_HEAD +
    `<p:sldMaster xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}">` +
    `<p:cSld>${spTreeXml(body)}</p:cSld>` +
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" ' +
    'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
    '<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>' +
    '</p:sldMaster>'
  )
}

/**
 * 版式 XML。
 * @param {object} layout - `LAYOUTS` 里的一项。
 * @returns {string} `<p:sldLayout>` XML。
 */
function slideLayoutXml(layout) {
  let body = ''
  if (layout.type === 'title') {
    body =
      placeholderXml({
        id: 2,
        name: '标题占位符',
        placeholder: 'ctrTitle',
        x: 838200,
        y: 1709738,
        cx: 10515600,
        cy: 1909763,
        text: '单击此处编辑母版标题样式',
        fontSize: 4400
      }) +
      placeholderXml({
        id: 3,
        name: '副标题占位符',
        placeholder: 'subTitle',
        idx: 1,
        x: 1750063,
        y: 3884613,
        cx: 9291325,
        cy: 1298575,
        text: '单击此处编辑母版副标题样式',
        fontSize: 2000
      })
  }
  return (
    XML_HEAD +
    `<p:sldLayout xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}" type="${layout.type}">` +
    `<p:cSld name="${escapeXml(layout.name)}">${spTreeXml(body)}</p:cSld>` +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
    '</p:sldLayout>'
  )
}

/**
 * 幻灯片 XML。
 * @param {object} args - 参数。
 * @param {object} args.layout - 版式定义。
 * @param {string} args.title - 标题文本。
 * @param {string} [args.subtitle] - 副标题文本。
 * @returns {string} `<p:sld>` XML。
 */
function slideXml({ layout, title, subtitle = '' }) {
  let body = ''
  if (layout.type === 'title') {
    body =
      placeholderXml({
        id: 2,
        name: '标题 1',
        placeholder: 'ctrTitle',
        x: 838200,
        y: 1709738,
        cx: 10515600,
        cy: 1909763,
        text: title,
        fontSize: 4400
      }) +
      placeholderXml({
        id: 3,
        name: '副标题 2',
        placeholder: 'subTitle',
        idx: 1,
        x: 1750063,
        y: 3884613,
        cx: 9291325,
        cy: 1298575,
        text: subtitle,
        fontSize: 2000
      })
  }
  return (
    XML_HEAD +
    `<p:sld xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}">` +
    `<p:cSld>${spTreeXml(body)}</p:cSld>` +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
    '</p:sld>'
  )
}

/**
 * 生成一份最小但完整的演示文稿。
 *
 * @param {object} [args] - 参数。
 * @param {string} [args.title] - 首页标题，默认「新建演示文稿」。
 * @param {string} [args.subtitle] - 首页副标题。
 * @param {string} [args.author] - 文档属性里的作者。
 * @param {'title'|'blank'} [args.layout] - 首页版式，默认 `title`。
 * @returns {{bytes: Buffer, slides: number, layout: string}} 生成结果。
 */
export function buildBlankPptx({ title = '新建演示文稿', subtitle = '', author = 'dsh-exp-office', layout = 'title' } = {}) {
  const spec = LAYOUTS[layout]
  if (!spec) throw new OfficeError('INVALID_REQUEST', `layout 只支持 ${Object.keys(LAYOUTS).join(' / ')}，实际「${layout}」。`)

  const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { ...DEFAULT_ZIP_LIMITS })

  // 1) 内容类型
  pkg.write(
    '[Content_Types].xml',
    XML_HEAD +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/ppt/presentation.xml" ContentType="${CT_PRESENTATION}"/>` +
      `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${CT_SLIDE_MASTER}"/>` +
      `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${CT_SLIDE_LAYOUT}"/>` +
      `<Override PartName="/ppt/slides/slide1.xml" ContentType="${CT_SLIDE}"/>` +
      `<Override PartName="/ppt/theme/theme1.xml" ContentType="${CT_THEME}"/>` +
      `<Override PartName="/docProps/core.xml" ContentType="${CT_CORE}"/>` +
      `<Override PartName="/docProps/app.xml" ContentType="${CT_APP}"/>` +
      '</Types>'
  )

  // 2) 包级关系
  pkg.write(
    '_rels/.rels',
    XML_HEAD +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="ppt/presentation.xml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
      `<Relationship Id="rId3" Type="${NS_R}/extended-properties" Target="docProps/app.xml"/>` +
      '</Relationships>'
  )

  // 3) presentation.xml
  pkg.write(
    'ppt/presentation.xml',
    XML_HEAD +
      `<p:presentation xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}" saveSubsetFonts="1">` +
      '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
      '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>' +
      `<p:sldSz cx="${SLIDE_W}" cy="${SLIDE_H}" type="screen16x9"/>` +
      '<p:notesSz cx="6858000" cy="9144000"/>' +
      '<p:defaultTextStyle><a:defPPr><a:defRPr lang="zh-CN"/></a:defPPr></p:defaultTextStyle>' +
      '</p:presentation>'
  )
  pkg.write(
    'ppt/_rels/presentation.xml.rels',
    XML_HEAD +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${NS_R}/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
      `<Relationship Id="rId2" Type="${NS_R}/slide" Target="slides/slide1.xml"/>` +
      `<Relationship Id="rId3" Type="${NS_R}/theme" Target="theme/theme1.xml"/>` +
      '</Relationships>'
  )

  // 4) 母版 + 版式 + 主题
  pkg.write('ppt/slideMasters/slideMaster1.xml', slideMasterXml())
  pkg.write(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    XML_HEAD +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${NS_R}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
      `<Relationship Id="rId2" Type="${NS_R}/theme" Target="../theme/theme1.xml"/>` +
      '</Relationships>'
  )
  pkg.write('ppt/slideLayouts/slideLayout1.xml', slideLayoutXml(spec))
  pkg.write(
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    XML_HEAD +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${NS_R}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
      '</Relationships>'
  )
  pkg.write('ppt/theme/theme1.xml', themeXml())

  // 5) 首页
  pkg.write('ppt/slides/slide1.xml', slideXml({ layout: spec, title, subtitle }))
  pkg.write(
    'ppt/slides/_rels/slide1.xml.rels',
    XML_HEAD +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${NS_R}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
      '</Relationships>'
  )

  // 6) 文档属性
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z')
  pkg.write(
    'docProps/core.xml',
    XML_HEAD +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
      'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dc:title>${escapeXml(title)}</dc:title><dc:creator>${escapeXml(author)}</dc:creator>` +
      `<cp:lastModifiedBy>${escapeXml(author)}</cp:lastModifiedBy>` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
      `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
      '</cp:coreProperties>'
  )
  pkg.write(
    'docProps/app.xml',
    XML_HEAD +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      '<Application>dsh-exp-office</Application><Slides>1</Slides>' +
      '<PresentationFormat>宽屏</PresentationFormat><Company></Company>' +
      '</Properties>'
  )

  return { bytes: pkg.toBuffer(), slides: 1, layout: spec.type }
}

/**
 * 转义 XML 文本。
 * @param {string} text - 原文。
 * @returns {string} 转义结果。
 */
function escapeXml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

export { LAYOUTS as PPTX_TEMPLATE_LAYOUTS, CT_PRESENTATION, CT_SLIDE, CT_SLIDE_MASTER, CT_SLIDE_LAYOUT, CT_THEME }
