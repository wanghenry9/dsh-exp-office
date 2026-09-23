# dsh-exp-office

面向 [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness) 的 Office/WPS 办公文件插件。
为 Agent 提供统一、稳定、安全的办公文件处理能力。

**当前版本 0.1.0｜阶段 0–2 已交付（骨架 + XLSX 全链路）｜零运行时依赖**

---

## 这是什么

一个 Cordis 插件，把 30 个 `office_*` 工具注册进 `ctx.tools`。注册后 schema 自动进入系统提示词，
宿主平面的所有 agent 都能读写 Excel 工作簿与 Word 文档。

核心设计是**字节级最小修改**：`.xlsx` 被当作 OPC（ZIP + XML）包直接操作，
只有目标单元格所在的 XML 区间被重写，其余字节原样输出。
因此修改一个单元格不会破坏图表、图片、宏、条件格式、数据验证、透视表或外部链接。

## 快速开始

```bash
# 在工作区中开发与自测
npm test                      # 全部套件（719 项断言，约 45 秒）
npm run test:engine           # 引擎：ZIP / XML 最小修改 / 安全防护
npm run test:xlsx             # XLSX 适配器（含样式与导出）
npm run test:docx             # DOCX 适配器（读取 / 写入 / 校验）
npm run test:pptx             # PPTX 适配器（读写/幻灯片管理/图片与表格）
npm run test:pdf              # PDF 适配器（读取 / 文本提取 / 页面级写操作）
npm run test:fidelity         # 保真度：真实 Excel 多特性样本
npm run test:word-fidelity    # 保真度：真实 Word 复杂样本（批注/修订/目录）
npm run test:plugin           # 插件集成：注册 / 端到端 / 事务 / 任务五件套
npm run test:scan             # XLSX 区域读取：字节扫描早停 / 上限 / 与 DOM 参照逐格等价
npm run test:automation       # 本地联动的 Node 侧（桩脚本，不启动 Office）
npm run verify:automation     # 本地联动真机验证（真实 Excel/WPS/Word/PowerPoint）


npm run test:perf             # 性能基准（小型 + 中型）
npm run test:perf:large       # 性能基准（追加 50 万单元格大文件；样本由生成器现造）
npm run fixture:xlsx-large    # 造大样本（--rows/--cols/--wide/--random/--out）
npm run verify:xlsx-large     # 真实 Excel 打开大样本（配合 test/read-cells-check.mjs 逐格对照）
npm run verify:boot           # profile 启动自检：副本漂移 + 真实 Cordis 装载（不开服务）
```

安装到 DSH profile 后，`cordis.patch.yml` 会自动挂载：

```yaml
- insert:
    - id: dsh-exp-office
      name: 'dsh-exp-office'
```

## 工具清单（90 个）

| 分类 | 工具 | 说明 |
|---|---|---|
| 文件 | `office_get_file_info` | 真实类型/大小/哈希/安全标志，不依赖扩展名 |
| 文件 | `office_list_files` | 盘点工作区内的办公文件 |
| 文件 | `office_create_document` | 新建 .xlsx，已存在时拒绝覆盖 |
| 文件 | `office_delete_file` | 删除文件（需 `confirm=true`） |
| 读取 | `office_read_workbook` | 工作表清单、可见性、安全标志 |
| 读取 | `office_read_sheet` | 稀疏单元格列表，支持分页 |
| 读取 | `office_read_range` | A1 区域读二维数组（推荐，只解析目标区域） |
| 写入 | `office_write_cells` | 批量写文本/数字/布尔/日期/公式 |
| 写入 | `office_set_cell_style` | 字体/底色/数字格式/对齐（语义化，自动查重样式定义） |
| 写入 | `office_find_and_replace` | 查找替换，默认跳过公式，支持 `dry_run` |
| 写入 | `office_merge_cells` | 合并区域 |
| 写入 | `office_create_table` | 建 Excel 表格对象（ListObject，含筛选与样式） |
| 写入 | `office_create_chart` | 建图表（柱状/条形/折线/饼图，含分类与数值缓存） |
| 写入 | `office_insert_rows` / `office_delete_rows` | 插入/删除行 |
| 写入 | `office_insert_columns` / `office_delete_columns` | 插入/删除列 |
| 工作表 | `office_add_worksheet` | 新增工作表 |
| 工作表 | `office_rename_worksheet` | 重命名工作表 |
| 工作表 | `office_delete_worksheet` | 删除工作表（需 `confirm=true`） |
| 导出 | `office_export_workbook` | 导出 xlsx 副本或 UTF-8 带 BOM 的 CSV |
| DOCX | `office_read_docx` | 读取 Word：元数据/段落/表格/页眉页脚/批注/修订/结构识别 |
| DOCX | `office_insert_paragraph` | 插入段落（可指定样式与对齐） |
| DOCX | `office_set_paragraph_style` | 套用段落样式（`w:pStyle`，缺定义自动补） |
| DOCX | `office_set_character_style` | 给段落内所有 run 套用字符样式（`w:rStyle`） |
| DOCX | `office_insert_page_break` | 插入分页符（独立段落里的 `w:br type=page`） |
| DOCX | `office_insert_bookmark` | 插入书签（`bookmarkStart`/`bookmarkEnd` 成对，id 全文档唯一） |
| DOCX | `office_update_paragraph` | 改写段落文本，保留样式与编号 |
| DOCX | `office_delete_paragraph` | 删除段落（需 `confirm=true`） |
| DOCX | `office_find_and_replace_docx` | Word 查找替换，支持 `dry_run` 与跨 run 警告 |
| DOCX | `office_set_docx_header` | 设置页眉文本，可选附加自动页码域 |
| DOCX | `office_set_docx_footer` | 设置页脚文本，可选附加自动页码域 |
| DOCX | `office_set_docx_page_layout` | 纸张（A4/A3/A5/Letter/Legal/自定义）与方向 |
| DOCX | `office_set_docx_margins` | 页边距（厘米，逐项改写；新建时写齐七项） |
| DOCX | `office_insert_docx_page_number` | 插入 `PAGE` 页码域（页脚/页眉，可加前后缀与对齐） |
| DOCX | `office_update_docx_table_cell` | 改写表格单元格文本，保留单元格格式 |
| DOCX | `office_create_docx_table` | 插入新表格（tblGrid 列宽 + 表头加粗 + 样式） |
| DOCX | `office_set_docx_table_style` | 套用表格样式（tblStyle + tblLook，缺定义自动补） |
| DOCX | `office_insert_docx_table_row` | 插入表格行（继承参照行的行/单元格格式） |
| DOCX | `office_delete_docx_table_row` | 删除表格行（保护性检查；仅剩一行时拒绝） |
| DOCX | `office_merge_docx_table_cells` | 合并矩形区域（横向 gridSpan + 纵向 vMerge） |
| DOCX | `office_insert_docx_image` | 插入图片（类型按字节判定、自动补齐部件与关系） |
| DOCX | `office_resize_docx_image` | 调整图片尺寸（`wp:extent` 与 `a:ext` 同步改） |
| DOCX | `office_delete_docx_image` | 删除图片（空段落一并删除；媒体与关系保留） |
| DOCX | `office_set_docx_image_wrap` | 图片环绕：行内 ↔ 四周型/上下型/浮于文字上方/衬于文字下方 |
| PPTX | `office_read_pptx` | 读取演示文稿：结构/幻灯片/备注/隐藏页/动画检测 |
| PPTX | `office_update_slide_text` | 改写形状文本，`\n` 分段并保留项目符号层级 |
| PPTX | `office_create_presentation` | 从零生成演示文稿（自造主题/母版/版式，不依赖模板文件） |
| PPTX | `office_add_slide` | 新增空白页，继承参照页版式 |
| PPTX | `office_delete_slide` | 删除页并清理孤儿部件（需 `confirm=true`） |
| PPTX | `office_reorder_slides` | 调整幻灯片顺序 |
| PPTX | `office_insert_slide_image` | 插入图片，默认居中，自动补齐关系 |
| PPTX | `office_add_slide_table` | 插入表格，默认套用内置样式并标记表头 |
| PPTX | `office_add_slide_chart` | 插入原生图表（柱状/条形/折线/饼图），含嵌入工作簿 |
| PPTX | `office_duplicate_slide` | 复制幻灯片到源页之后（丢弃备注页关系，避免 PowerPoint 报修复） |
| PPTX | `office_add_text_box` | 插入文本框（形状 id 全页唯一，像素→EMU 换算） |
| PPTX | `office_set_slide_layout` | 切换幻灯片版式（改 slideLayout 关系目标，不动形状） |
| PPTX | `office_set_theme` | 切换母版主题（克隆新主题部件后挂载，配色可被 COM 验证） |
| PPTX | `office_export_presentation` | 导出 PDF（内置引擎；说明动画/切换/媒体丢失并核对页数） |
| PPTX | `office_validate_pptx` | 演示文稿结构校验；明确列出未检查项 |
| PDF | `office_read_pdf` | 读取元数据/页数/尺寸/文本/结构（含加密与签名检测） |
| PDF | `office_fill_pdf_form` | 填写 AcroForm 表单（文本/复选/单选/下拉），只改 /V 与 /AS + NeedAppearances |
| PDF | `office_search_pdf` | 全文搜索：页码 + 页内字符偏移 + 行号 + 上下文，按字面量匹配 |
| PDF | `office_validate_pdf` | PDF 结构校验；列出未检查项 |
| PDF | `office_rotate_pdf_pages` | 旋转页面（单页或全部，可累加），增量更新写回 |
| PDF | `office_delete_pdf_page` | 删除页面（从页面树摘引用）；仅剩一页时拒绝 |
| PDF | `office_reorder_pdf_pages` | 重排页面（`order` 全排列或 `from`+`to` 单页移动） |
| PDF | `office_add_pdf_annotation` | 加便签批注（`/Subtype /Text`，**内容支持中文**） |
| PDF | `office_add_pdf_watermark` | 叠加式加水印文字（英文/数字，斜向或水平） |
| PDF | `office_add_pdf_page_numbers` | 叠加式加页码（可设格式、起始编号、页眉/页脚） |
| PDF | `office_set_pdf_metadata` | 写入标题/作者/主题/关键词/创建程序/生成程序 |
| PDF | `office_optimize_pdf` | 对象图重写：回收增量更新的字节；可原地或另存 |
| PDF | `office_split_pdf` | 按页拆出新文件；未选页在字节里彻底消失 |
| PDF | `office_merge_pdfs` | 合并多份 PDF（跨文件对象图搬运，各自页尺寸/字体互不干扰） |
| PDF | `office_extract_pdf_images` | 提取页面图片（JPEG/JP2 原样导出，Flate 流还原预测器后包成 PNG） |
| 校验 | `office_validate_workbook` | 结构/关系/公式/宏/图表校验，可与基线比对 |
| 校验 | `office_validate_docx` | Word 结构校验；明确列出未检查项与字体清单 |
| 校验 | `office_compare_docx` | 两份 Word 的段落级差异（新增/删除/修改 + 文字 vs 样式）、表格逐格变化、样式集合与元数据差异；只读，不写修订标记 |
| 本地联动 | `office_detect_engines` | 检测本机 Office/WPS 组件与版本（默认只读注册表；`probe_com=true` 会实测 COM 可用性，需授权） |
| 本地联动 | `office_recalculate` | 用**真实引擎**重算（Excel 公式缓存 / Word 域）后另存新文件；源文件永不改动（需授权） |
| 本地联动 | `office_rerender` | 用**真实引擎**整份重写（等价于用 Office 重新保存）后另存新文件；源文件永不改动（需授权） |
| PDF | `office_create_pdf` | **从零生成** PDF（自写页面树/内容流/字体资源/xref，含 `/ToUnicode` 因此文本可提取）；拉丁文本用标准 14 字体，含中文时**自动嵌入子集化字体** |
| PDF | `office_extract_pdf_tables` | 按文本位置推断表格（行按 y 聚类、列按 x 聚类，不依赖框线），返回二维单元格文本 |
| 转换 | `office_convert_document` | docx/xlsx/pptx → PDF（宿主内置 LibreOffice 引擎） |
| 任务 | `office_preview_operation` | 预览计划，返回 `plan_id` 与风险评级 |
| 任务 | `office_execute_operation` | 执行计划（自动备份 + 校验） |
| 任务 | `office_rollback_operation` | 按备份回滚 |
| 任务 | `office_get_task_status` | 查询计划状态 |
| 任务 | `office_cancel_task` | 取消未执行的计划 |

## 保真度保证

| 保证 | 实现方式 | 验证 |
|---|---|---|
| 无关部件逐字节不变 | 未修改的 ZIP 条目复用原压缩字节 | ✅ 测试断言部件级字节相等 |
| 未触碰的 XML 逐字节不变 | 解析器记录源码偏移，编辑只产生区间补丁 | ✅ 测试断言补丁外区域完全一致 |
| 公式本身与缓存结果分离 | 读时分别返回 `formula` / `value`，默认不重算 | ✅ Excel/WPS 实测公式保留 |
| 保存后必须能重新打开 | 事务提交前强制重开校验，失败即回滚 | ✅ 每次写入都过校验闸门 |
| 失败不破坏原文件 | 暂存文件 + 原子替换，失败路径必清理 | ✅ 测试断言失败后原文件字节不变 |
| 宏按原字节保留 | `vbaProject.bin` 是普通 ZIP 条目，走同一套透传逻辑 | ✅ 改名/编辑/多轮编辑/增删工作表后均字节不变 |
| 真实类型检测 | 按魔数与包内标志部件判定，不信扩展名 | ✅ 把 .docx 改名成 .xlsx、把 .xlsm 改名成 .xlsx 均被识破 |

> **宏验证的诚实边界**：本机没有真实宏文件样本，且用 Excel 写 VBA 需要开启「信任对 VBA 工程对象模型的访问」——那属于改动你 Office 的安全设置，我没有擅自开启。因此宏测试用的是**合成** `vbaProject.bin`：它证明了「检测 + 字节保留」，**未证明** Excel 能运行其中的宏。要补齐这一项，需要一个真实 .xlsm 样本，或你授权调整该安全设置。

**已用真实软件验证**（不是只跑单元测试）：

在**真实 Microsoft Excel 创作**的工作簿（含图表、条件格式、合并单元格、跨表公式、数字格式、双工作表）上修改一个普通单元格后，逐项复核：

| 复核项 | 结果 |
|---|---|
| 图表部件 `xl/charts/chart1.xml` | 逐字节不变 |
| 绘图与图表关系部件 `xl/drawings/*` | 逐字节不变 |
| 主题 / 样式 / 文档属性 / `[Content_Types].xml` | 逐字节不变 |
| 发生变化的部件 | **仅目标工作表** `xl/worksheets/sheet2.xml` |
| Excel 重新打开后的图表数量 | 1（标题完好） |
| Excel 重新打开后的条件格式规则 | 1（保留） |
| Excel 重新打开后的合并单元格 | 保留 |
| Excel 重新打开后的跨表公式 | 重算为 99,800.00 |
| WPS 表格重新打开 | 工作表名、数值、公式全部无损 |

对应测试：`node test/fidelity.test.mjs`（17 项）、`npm run verify:excel`、`npm run verify:wps`。

## 安全设计

- **路径**：拒绝路径穿越、UNC/设备路径、空字节；所有访问限定在工作区根目录内。
- **类型**：按真实字节识别（PDF magic、OLE 容器、ZIP 内标志部件），改扩展名无法绕过。
- **压缩炸弹**：限制条目数、单条解压大小、总解压量与压缩比。
- **XML 攻击**：拒绝 DOCTYPE/DTD/外部实体；限制字节数、节点数、嵌套深度。
- **宏与外部资源**：检测并报告，但从不执行宏、不更新外部链接、不访问网络。
- **破坏性操作**：删除工作表/文件必须显式 `confirm=true`；预览会标记风险等级。
- **并发**：独占文件锁 + 残留锁超时回收。
- **资源**：随机临时目录，成功与失败路径都清理；锁文件不残留。
- **日志**：不记录文档正文，仅记录操作类型、范围与哈希。

## 支持范围

| 格式 | 读取 | 写入 | 校验 |
|---|---|---|---|
| `.xlsx` | ✅ | ✅ 单元格/样式/行列/合并/**表格对象**/**图表** | ✅ |
| `.xlsm` / `.xltx` | ✅（宏按原字节保留，不执行） | ✅ | ✅ |
| `.docx` | ✅ 元数据/段落/表格/页眉页脚/结构识别 | ✅ 段落插入·改写·删除、查找替换、页眉页脚、表格单元格、图片 | ✅ |
| `.pptx` | ✅ 尺寸/母版/版式/主题/幻灯片文本/备注/隐藏页/动画检测 | ✅ 文本改写、幻灯片增删排序、图片与表格插入 | ✅ |
| `.pdf` | ✅ 元数据/页数/尺寸/文本提取/结构识别 | ✅ 页面旋转·删除·重排 + 水印·页码（增量更新写回） | ✅ |

**任意 Office 格式 → PDF**：`office_convert_document`（见下方「转 PDF」）。

### XLSX 表格对象（ListObject）

`office_create_table` 把一片区域变成 Excel 表格：带表头、自动筛选与套用样式。要写齐**四处**，
少任何一处 Excel 都会判定文件需要修复：

1. `xl/tables/tableN.xml` 本体（列名、区域、样式）；
2. 工作表里的 `<tableParts><tablePart r:id="…"/></tableParts>`（schema 顺序要求在 `extLst` 之前）；
3. 工作表关系部件里指向表格部件的 relationship；
4. `[Content_Types].xml` 里表格部件的 Override。

另外两条硬约束：**列名必须非空且互不重复**，且应与表头单元格一致 —— 表头为空或重复时，
插件会把生成的列名（列1、列2…）**写回表头单元格**（Excel 自己建表也是这个行为），
而不是留下「表头空着、表格却说叫列1」的别扭状态；同一工作表内与该区域重叠的表格会被拒绝。

**真实软件验证**（`npm run verify:table`）：

| 软件 | 结果 |
|---|---|
| Microsoft Excel 16 | `LISTOBJECTS_销售数据: 1`，`name=销售表 range=A1:D5 header=A1:D1 totals=False style=TableStyleMedium9`，`OPEN_OK` |
| WPS 表格 12 | `LISTOBJECTS: 1`，`name=销售表 range=A1:D5`，`WPS_OPEN_OK` |

> **踩坑记录**：第一版把 `<tableParts>` 挂到了**文档节点**上（`doc.appendChild(doc.root, …)` 而不是
> 挂到 `<worksheet>` 元素），结果写出一个**有两个根元素**的 XML。本插件自己的解析器是宽松的，
> 照样能读、测试全绿；**Excel 直接判文件损坏**。后来给 `office_validate_workbook` 加了一条
> 「标签配平、单一根元素」的格式良好检查，并补了回归测试，这类问题不再只靠真机验证兜底。

### XLSX 图表

`office_create_chart` 建柱状/条形/折线/饼图。同样要写齐**五处**：图表部件、绘图部件
（`twoCellAnchor` + `graphicFrame` 引用图表）、工作表的 `<drawing>`、工作表关系 → 绘图、
绘图关系 → 图表，以及图表与绘图两个内容类型 Override。工作表已有绘图（图片或别的图表）时
**复用同一个绘图部件并追加锚点** —— 一张工作表只能有一个 `<drawing>` 引用。

- 分类与数值会同时写入**缓存值**（`c:cat` 的 strCache、`c:val` 的 numCache）：规范里可以省略，
  但省略后没有任何阅读器能立刻画出图，得先重算；数据本来就在手里，没理由不给
- 系列名可以给字面文本（`name`）或取自单元格（`name_ref`，会写成引用公式）
- 支持跨表数据源（`values: '销售数据!D2:D4'`），引用公式按 OOXML 规则加表名与绝对引用
- 锚点按工作表的**真实列宽/行高**换算（读 `sheetFormatPr` 与 `<cols>`），不是按默认值瞎估
- 轴 id 与元素顺序严格按 schema 排列（Excel 对图表 XML 的顺序极敏感，错一处就要求修复）

**真实软件验证**（`npm run verify:chart`）：

| 软件 | 结果 |
|---|---|
| Microsoft Excel 16 | `CHARTCOUNT: 1`，`type=51`（柱状簇）、`title='各产品金额'`、`series=1`，`OPEN_OK` |
| WPS 表格 12 | `WPS_OPEN_OK`（无修复提示） |

### PDF 能力（阶段 5）

`office_read_pdf` 四种视图：`summary` / `structure` / `pages` / `text`。

- **元数据**：标题、作者、主题、关键词、生成器、创建/修改日期
- **页面**：数量、每页尺寸（pt）、旋转、内容流数量，按 `/Kids` 顺序归一化
- **文本提取**：解析内容流的 `Tj` / `TJ` / `'` / `"` 操作符；通过字体的 **ToUnicode CMap** 解码，**中文不会乱码**（不做这一步 CJK 会提取成乱码）
- **结构识别**：是否加密（含算法与密钥长度）、数字签名、表单字段、批注、图片、字体清单、嵌入文件

**便签批注** —— `office_add_pdf_annotation`：

- **批注文字支持中文**：便签由**阅读器自己排版渲染**，不需要在 PDF 里嵌入字体 ——
  与页面上的水印/叠加文字（只能 ASCII）是两套机制。内容按 **UTF-16BE + BOM** 写入 PDF 字符串，
  这是 PDF 表示非 ASCII 文本的标准做法
- 写入同样是**增量更新**：新建批注对象，把引用追加到该页 `/Annots`，页面内容流与其它对象一个字节不动
- `/Annots` 是**间接引用**（指向一个数组对象）时会解引用后追加，并把那个数组对象一并写回 ——
  直接覆盖会让原有批注全部消失（这条有专门的回归用例与合成样本）
- 写 `F=4`（Print 标志）：不写的话批注打印时会被丢掉；位置用 PDF 用户空间坐标（**原点在左下角**，pt），
  省略时放页面左上角内侧
- 读回视图：`office_read_pdf(detail="annotations")` 给出每条的 page / subtype / contents / author / rect / flags / open

> **验证边界（如实说明）**：批注的**视觉呈现**（便签图标在阅读器里的位置与外观）本机无法验证 ——
> 没有 Acrobat，Word 的 PDF 重排也不显示批注。已验证的是：插件读取器能逐字段读回（含中文）、
> 结构校验通过、**Word 自带的 PDF 解析器仍能正常解析**（`REFLOW_PAGES 2 / WORDS 43`，文本完好），
> 以及增量更新不触碰原有字节。图标实际长什么样，请你在阅读器里看一眼。

**AcroForm 表单填写** —— `office_fill_pdf_form`：

- 覆盖三类终端字段：文本（`Tx`）、按钮（`Btn`：复选框 / 单选组）、选择（`Ch`：列表 / 下拉）；
  另有 `read_pdf(detail="form")` 读回字段名、类型、当前值与可选项
- **开状态名要读出来，不能猜**：复选框/单选的 `/AP /N` 子键可能是 `Yes`、`1`、`开`……
  工具按实际读到的名字写 `/AS`，非法值直接报错并把允许值列出来
- **单选组不能整组置开**：每个 widget 有自己的开状态名，只有与目标值同名的那一个置为开、其余置 `Off`
  （否则整组都会显示选中）
- 只改 `/V`（按钮另改 `/AS`）并要求阅读器按 `/NeedAppearances` 重绘外观 ——
  **不生成外观流**：那需要把文字按字体度量渲染成 XObject，属于排版引擎的活；宁可让阅读器重绘，也不画一个错的
- **两遍式写入**：先全量校验、再开始写。中途失败不会留下「填了一半」的状态
  （回归用例断言失败后 `dirty=false` 且字节与源文件一致）
- 中文值走 UTF-16BE + BOM（PDF 字符串表示非 ASCII 的标准做法），读回逐字符一致
- 只读字段（`/Ff` 第 1 位）默认拒绝，需 `force=true`；数字签名与按钮字段明确拒绝

> **验证边界**：字段**值的读回**、`/AS` 外观状态、`/NeedAppearances`、增量更新不触碰原有字节、
> 以及 **Word 自带 PDF 解析器仍能正常读取**（`REFLOW_PAGES 1 / WORDS 2`）都已验证；
> 但**填写后的视觉呈现**（文字是否落在框里、位置对不对）本机无法验证 —— 没有 Acrobat，
> 请你在阅读器里打开 `test/fixtures/form-filled.pdf` 看一眼。

**全文搜索** —— `office_search_pdf`：

- 返回每处命中的**页码 + 页内字符偏移 + 行号 + 前后上下文**，以及 `matches_by_page` 与总命中数
- 查询按**字面量**处理（正则元字符一律转义）：查 `(a)`、`[`、`1.5`、`.` 都只是找这些字符本身，
  既不会误匹配、也不会因 `[` 抛出正则语法错误
- 大小写不敏感走正则 `i` 标志，**不用 `toLowerCase()` 比较**：个别字符（如 `İ`）小写化后长度会变，
  拿长度当偏移映射会把位置算错；命中返回的是**原文大小写**
- `maxResults` 截断时 `truncated=true` 且 `total_matches` 仍是真实总数；`page` 可限定单页
- **位置是页内字符偏移，不是页面坐标**：给 x/y 要按文本矩阵与字体宽度算版面，那是排版引擎的活，
  插件不假装能做到（`not_checked` 里也没有「版面还原」这种含糊说法）
- **独立交叉验证**：`pwsh -File test/pdf-open-check.ps1 -Path test/fixtures/sample-multipage.pdf -Match XLSX`
  让 **Word 自带的 PDF 解析器**（PDF Reflow，与我们的实现完全无关）数同一个关键词，
  `MATCH_COUNT: 2` 与 `office_search_pdf` 的 `total_matches: 2`（第 2、3 页各一处）**逐项吻合**

**解析策略上的明确取舍**：标准做法是先读 `startxref` → 交叉引用表 → 按引用遍历。但现代 PDF 普遍使用**交叉引用流**与**对象流**，完整实现两者代码量与出错面都很大。本实现对整个文件做**对象扫描**并从对象流补充成员：对「读取」更宽容（文件尾部截断或增量更新链断裂时仍能读出内容），代价是失去交叉引用一致性校验能力——因此 `validate()` 会把「交叉引用表一致性」如实列进 `not_checked`，而不是假装校验过了。

**写回用增量更新（incremental update）** —— `office_rotate_pdf_pages` / `office_delete_pdf_page` / `office_reorder_pdf_pages`：

- **原有字节一个都不动**：只在文件尾部追加被修改对象的新版本、新的交叉引用段，以及带 `/Prev` 指回旧 xref 的 trailer。这与 OOXML 侧「只改目标字节区间」是同一思路——没被触碰的内容不可能被改坏
- **页面树才是文档的页**：读取以页面树的 `/Kids` 为准取页面的集合与顺序。若只按 `/Type /Page` 扫描，`delete_page` 摘掉引用后，那个对象仍留在文件里（增量更新不回收字节），重新打开会看起来「删了没生效」
- **旋转**：写 `/Rotate`（0/90/180/270 归一并支持累加），不动页面内容流
- **删除**：从 `/Kids` 摘引用并更新 `/Count`；仅剩一页时拒绝（删完就不是有效文档了）
- **重排**：只改 `/Kids` 顺序，页面内容、资源、批注一律不动
- **加密 PDF 拒绝改写**（不带口令改写只会产出损坏文件）；含数字签名时明确提示「修改会使签名失效」
- 写回前会**重新打开并校验**输出，失败即回滚，绝不留半个文件

**叠加式写入（水印 / 页码）** —— `office_add_pdf_watermark` / `office_add_pdf_page_numbers`：

- **原内容流一个字节都不改**：新建一个只画这一行字的内容流，追加到页面 `/Contents` 之后。
  PDF 规定 `/Contents` 数组里的多个流按顺序画在同一张画布上，叠加是天然语义 ——
  这正是开发要求 §八.4 把叠加式列为稳定性最高写法的原因
- **字体用阅读器内置的 Helvetica（base-14）**，不嵌入字体，所以拉丁水印新增体积只有几百字节到 1 KB
- **中文水印 / 页码自动嵌入字体子集**：含非 WinAnsi 字符时，用 `pdf-font.js` 在叠加流的资源里挂一份
  **只带用到的字形的 TrueType 子集**（Type0 + CIDFontType2 + FontFile2 + CIDToGIDMap + ToUnicode）。
  实测 3 页中文水印让 80 KB 的样本变成 92 KB（子集未压缩 38.9 KB），**原字节作为前缀一个都没变**
- **资源是合并而不是覆盖**：页面级 `/Resources` 会整体覆盖从页面树继承来的资源，
  只写一个含水印字体的资源字典会让**原有文字丢掉字体**（我们实测到中文退化成控制字符）。
  实现上复制「有效资源」再在其 `/Font` 之上加一个 `Wm` / `WmCJK`
- **字体与资源字典必须内联**：写成 `/Resources → 对象 → /Font → 对象 → /Wm → 对象` 的多级间接引用，
  规范合法、本插件自己的解析器也读得出来，但 **Word 的 PDF 解析器会静默丢弃这些文字** ——
  这个坑只有第三方阅读器能暴露（见下）
- **嵌入字体的对象号必须一次预留**：Type0 → CIDFont → FontDescriptor → FontFile2 / CIDToGIDMap / ToUnicode
  是**互相引用**的六个对象，必须先预留真实对象号再构造。踩过一次：内部引用指向了别处，
  `ToUnicode` 指到一个页面内容流上，**自读的文本全成了 `\u0000`**（自家 `validate()` 照样通过）
- **淡化用浅灰填充**（`0.9 g`）而不是透明度软掩码（ExtGState SMask）：结果确定、少一层资源依赖
- 页码支持格式串（`{page}` / `{n}` / `{total}`）、起始编号与页眉/页脚位置；中文页码如「第 {page} 页 / 共 {total} 页」

> **第三方解析器验收（`npm run verify:pdf-cjk-watermark`）**：给真实 3 页 PDF 加中文贴边水印后，
> Word 的 PDF 重排仍读到**全部 66 个词的原正文**（标题与表格都在），并在 **footer story** 里读到了
> 我们叠加的中文水印「机密 · 中文水印 1/3」。
>
> 同时记一条**观察**：Word 的重排对「整份文档都盖着居中斜向大水印」会退化 ——
> 实测 3 页全覆盖时它只吐出 31 个词，且页数从 3 变 5；换成贴边（bottom）位置或只盖 1–2 页就一切正常。
> 我们的读取器与结构校验在这两种情况下都正常，因此判断是 **Word 重排的版面切分行为**，不是文件损坏；
> 验收用贴边位置，这样它既能独立核对水印文字，也能核对原正文。

**对象图重写（回收字节 / 拆分）** —— `office_optimize_pdf` / `office_split_pdf`：

- 与增量更新相反：只写出从 trailer 出发**可达**的对象，重新编号、重建经典 xref 表。
  增量更新累积的旧对象、旧交叉引用段、对象流容器全部消失，页面/字体/图片/元数据保留；
  内容流**保持原压缩字节**搬运，不做解压再压缩
- **两种写法的取舍要讲清楚**：增量更新「原字节一个都不动」（最小风险），代价是文件只增不减、
  被删页面的对象仍留在字节里；重写「字节里只剩该有的东西」，代价是整份文件被重建
- **拆分必须做引用清理**：把页面从 `/Kids` 摘掉是不够的 —— 书签（`/Outlines`）、命名目标、
  链接批注的 `/Dest`、结构树的 `/Pg` 都可能直接指向那个页面对象。只要还有一处引用，
  可达性遍历就会把整页内容又写回新文件（**实测踩到过：拆 1 页出来，字节里仍有 3 个页面对象**）。
  实现上按容器删除这类引用（数组项删除、`/Dest [页 …]` 整项删除、字典键删除），并报告清理条数
- **硬不变量**：重写后字节里 `/Type /Page` 的出现次数**恰好等于页数**。
  这也是「拆分能不能安全对外分发」的判据 —— 用文本搜索做泄漏检查是不可靠的（内容流是压缩的，
  正文本来就不以明文出现）
- 原地重写走事务（暂存 → 重开校验 → 原子替换）；另存/拆分写完会**重新打开校验**，
  页数或结构不符就把刚写的文件删掉并报错，绝不留半个产物
- 引用指向不可解析的对象时**拒绝重写**，而不是产出悄悄缺内容的文件；加密 PDF 拒绝

**合并** —— `office_merge_pdfs`：跨文件搬运对象图（每份文档整体错开编号、引用跟着改写），
页面树用嵌套写法挂到新建的 `/Pages` 根下，因此各份文档的**页尺寸、字体与图片各自独立**。
只搬运页面内容与资源：目录（书签）、命名目标、表单、结构树与页面标签**不会带过来**
（跨文档合并它们需要重映射，硬带只会得到互相打架的引用），元数据取第一份。
合并产物仍是标准 PDF，可以继续旋转/拆分/重写。

**图片提取** —— `office_extract_pdf_images`：只解出 PDF 里**已经存在的图像流**，不做页面光栅化。

- `/DCTDecode`（JPEG）与 `/JPXDecode`（JPEG 2000）**原样导出**，不重新编码（不损失画质）
- `/FlateDecode` + 8 位分量 → 还原预测器（Predictor 1/2/10–15，含 Paeth）后包成 PNG；
  索引色展开为 RGB；`/SMask` 软掩码合成成带透明通道的 PNG
- **不支持的形态明确跳过并给出原因**（1/2/4 位分量、CMYK、CCITT/JBIG2、色彩空间无法解析、
  像素数据不足、软掩码尺寸不符），而不是导出打不开的文件
- 同一张图被多页共用时**只导出一次**，结果里列出用到它的页码；写完会做一次自检
  （真实类型与尺寸必须与计算结果一致），不一致就删掉文件并报错
- **独立验证**：`test/image-decode-check.mjs` 用第三方解码器（sharp/libvips，从宿主环境解析，
  插件本体不依赖它）打开导出的图片，并检查**通道均值**。样本是纯色图 `(31,119,180)`，
  实测均值正好是 `[31, 119, 180]` —— 这证明「导出的确实是被插入的那张图」，
  而不只是「一张能打开的图」

**验证**：

1. 结构契约测试——`startxref` 指向 `xref` 关键字、每个 xref 条目的偏移精确指向 `N G obj`、
   `/Prev` 等于原文件的 `startxref`、原文件是输出的**字节前缀**（增量路径）；
   重写路径则断言 trailer 的引用是 `num gen R` 三段式、经典 xref 表覆盖全部对象、
   字节里 `/Type /Page` 计数恰好等于页数
2. 独立第三方阅读器——用 **Microsoft Word 的 PDF 重排**（不是本插件、也不是 LibreOffice）打开改写后的文件：
   - 页面重排：文本顺序变为「本季度进展 → 表格页 → 季度业务报告」
   - 删除页面：被删页的文本从 Word 读到的内容里消失
   - **页码**：Word 把它读进了**页脚故事**（`FOOTER: 1 / 3`）
   - **水平水印**：Word 把它读进了**页眉故事**（`HEADER: HORIZONTAL MARK`）
   - **重写/拆分后**：Word 仍能读出全部文本（含叠加的水印），拆分件的内容顺序与页序一致
   - 原有中文全部完好（说明资源合并没丢字体）
3. 样本本身由真实 LibreOffice 引擎从中文演示文稿导出（PDF 1.6、交叉引用流、压缩内容流、子集化 CJK 字体）

> **验证方法论上的一个坑（值得记下来）**：Word 报的 `PAGES` 是**它自己重排后的页数**，不是 PDF 页数 ——
> 实测 1 页的 PDF 被排成 2 页、3 页的被排成 4 页（它会把同一页的多个内容流当成不同段）。
> 所以页数一律以「插件读取 + xref 结构契约 + `/Type /Page` 计数」为准，Word 只用来验证**文本内容**。
> 同理，用文本搜索判断「被删页内容是否还在文件里」不可靠：内容流是压缩的，正文本来就不以明文出现。

**已知边界（如实列出）**：

- **斜向（45°）水印没有在第三方阅读器里验证过**：Word 的 PDF 重排**会丢弃旋转文字**——
  这一点用最小样本逐项隔离确认过（同一绘制指令，水平能读到、旋转读不到）。
  斜向水印目前只验证到「内容流写对了、文件结构合法、本插件能读回文字」。
  要视觉确认需要能渲染 PDF 的引擎，本机没有（宿主内置的是 LibreOfficeKit，不接受 PDF 输入；
  Edge 无头模式不加载 PDF 插件；profile 里也没有 pdfjs/canvas）。**水平水印（`position: top|bottom`）则是被独立阅读器读到的**
- **中文水印/页码需要嵌入字体子集** —— 现已实现（自动嵌入，也可用 `cjk_font_path` 指定字体）
- 增量更新会让文件**变大**（旧对象与旧 xref 保留在文件里）；用 `office_optimize_pdf` 重写即可回收，
  但重写后普通对象不再走对象流压缩，**体积不一定比原始文件更小**（会返回 `NO_SIZE_GAIN` 提示）
- 重写会**丢弃对象流与交叉引用流**（换成经典 xref 表与普通对象）：语义等价、布局不同，
  对绝大多数阅读器无影响，但依赖对象流的增量差异工具可能看出差别
- 重写/拆分时，指向被丢弃页面的引用会被**删除**（书签目标、链接批注、结构树的 `/Pg`），
  返回值里给出清理条数；被删掉书签的目录项会保留标题但不再跳转
- 嵌套页面树（`/Kids` 里有 `/Pages` 子节点）能**读**，拆分也支持；但重排/删除会明确拒绝
  （`UNSUPPORTED_FEATURE`），而不是猜着改
- 尚未实现：PDF 合并、原地文本改写、批注、表单填写、OCR、图片提取、PDF → Office 转换

### PPTX 能力（阶段 4）

**读取** —— `office_read_pptx` 五种视图：`summary` / `structure` / `slides` / `slide` / `text`。

- **演示结构**：幻灯片尺寸与宽高比（识别 16:9 / 4:3 / 16:10）、母版 / 版式 / 主题数量
- **每页内容**：标题、各形状文本（含占位符类型：居中标题/副标题/正文/页脚…）、表格/图表/SmartArt/图片形状分类
- **演讲者备注**：只取讲稿占位符，不混入备注页上的页码占位符
- **隐藏幻灯片**：标记实际写在幻灯片部件根元素的 `show="0"` 上，而不是 `presentation.xml` 的 `sldId` —— 这一点用真实样本确认，并用「转 PDF 只导出 3 页」独立交叉验证
- **动画与切换**：`<p:timing>` 即使无动画也存在，因此会进一步检查是否真有 `p:anim*` 节点，避免把全部幻灯片误报为有动画
- **显式提示**：SmartArt、音视频、嵌入对象、动画都会作为警告返回，不静默忽略

**写入** —— `office_update_slide_text` / `office_add_slide` / `office_delete_slide` / `office_reorder_slides` / `office_insert_slide_image` / `office_add_slide_table` / `office_duplicate_slide` / `office_add_text_box`。

- `update_slide_text`：改写某个形状的文本，`\n` 分段并复用原段落的 `<a:pPr>`（保住项目符号层级与缩进）
- `add_slide`：新增空白页，**继承参照页的版式**（引用同一个 slideLayout），因此标题/正文占位符样式与母版一致；会一并补齐部件、内容类型 Override、演示文稿关系与 `sldIdLst` 条目
- `delete_slide`：删除页面的同时清理其部件、关系部件与内容类型声明，**不留孤儿部件**（只从 `sldIdLst` 摘引用会让文件里留下永不被引用的幻灯片）
- `duplicate_slide`：复制幻灯片到源页**紧后面**，连带复制该页的关系部件（图片、图表），但**丢掉备注页关系** —— 共用同一个 `notesSlide` 部件会让 PowerPoint 报「需要修复」。实测：源页 `HASNOTES=yes`、副本页 `HASNOTES=no`
- `add_text_box`：插入文本框，写齐 `p:nvSpPr` / `p:spPr` / `p:txBody` 三段，形状 id 取全页最大值 +1（**重复 id 会让 PowerPoint 报修复**），位置尺寸按 96 DPI 换算 EMU；WPS 演示把它识别为 `type=17` 文本框（390×90pt = 520×120px）
- `reorder_slides`：只重排 `sldIdLst` 条目，不动任何页面内容
- `insert_slide_image`：插入图片，类型按**真实字节**判定；未指定尺寸用原始像素（96 DPI），未指定位置**按幻灯片居中**（默认落左上角会盖住标题）；一并补齐媒体部件、内容类型与幻灯片关系
- `add_slide_table`：插入表格，默认套用 PowerPoint 内置表格样式并标记首行为表头（不指定样式时表格会渲染成**无边框裸文本**）；行列参差时按最大列数补齐
- `create_presentation`：**从零生成**一份 .pptx，不依赖任何模板文件。必需部件全部程序化写出：
  内容类型、包级关系、`presentation.xml`（sldMasterIdLst / sldIdLst / sldSz / notesSz）、
  母版（**必须带 `p:clrMap` 十项**）、版式、主题（`clrScheme` + `fontScheme` + `fmtScheme`）、首页、文档属性。
  配色与字体是**本插件自己的**（主题名 `office-plugin`），不复制 Office 文件里的内容；
  生成后立刻用自家读取器自检，不通过就**不留文件**
  - **踩坑（只有真机能发现）**：主题里渐变停靠点的颜色必须**直接**写在 `<a:gs pos="…">` 下；
    包一层 `<a:solidFill>` 看着像、schema 上却不是合法的颜色选择 —— PowerPoint 会因此**拒绝打开整个文件**。
    定位方法：把自家部件逐个换成 PowerPoint 自己写的同名部件，哪一次能打开就说明被换掉的那份 XML 有问题
    （从「整份主题」一路二分到 `fmtScheme` 里的一行）
  - 生成出来的是**活文件**：实测在其上继续 `add_slide` / `add_text_box` / `add_slide_table` /
    `add_slide_chart` / `set_slide_layout` 后，PowerPoint 读到 2 页、3 个形状、`charttype=51` 且数值正确，
    WPS 演示读到 2 页与 `TABLE_2_2 rows=2 cols=2`，两边都无修复提示
- `set_theme`：换母版主题（配色/字体）。**不能直接把母版指向别处已在用的主题部件** ——
  实测把幻灯片母版指向「备注母版的主题」后，PowerPoint **拒绝打开整个文件**（对照组：只把主题内容换掉则正常打开）。
  所以实现与 PowerPoint 自己一致：**克隆出一个新主题部件**（`themeN+1.xml`）再挂上去；
  当前主题若只被这一个母版引用，则就地改它的内容，因此来回切换不会无限堆积部件。
  主题名里的**零宽字符**（简体中文版 PowerPoint 会写 `Office 主题` + 两个 `U+200B`）在入口统一清掉，
  否则「按名字匹配」永远匹配不上、而错误信息看起来一模一样
- `add_slide_chart`：插入**原生图表**（`column` / `bar` / `line` / `pie`），写齐图表部件、嵌入工作簿、图表关系、
  幻灯片关系、`p:graphicFrame` 与内容类型六处接线。系列数据写**单元格引用 + 缓存**，并嵌入一份
  **由本插件自己的 XLSX 写入器生成**的工作簿，因此在 PowerPoint 里「编辑数据」可用。
  下面三条是实测踩出来的（自家校验器与 Excel 都能读通，但 PowerPoint 拒绝打开或静默丢弃）：
  - `c:externalData` 的位置卡得很死 —— CT_ChartSpace 的顺序是 `… chart spPr? txPr? externalData? …`，
    必须排在 `c:spPr` **之后**、`</c:chartSpace>` 之前；放到 `c:chart` 或 `c:spPr` 之前，
    PowerPoint 直接报「**无法打开该文件**」（不是提示修复）
  - `a:graphicData` 的 `uri` 必须是 `…/drawingml/2006/chart`：写错时**不报错也不显示**，图形静默消失
  - 系列数据**不能**只用字面量（`c:strLit` / `c:numLit`）：那样写 PowerPoint 同样拒绝打开；
    改成 `c:strRef` / `c:numRef` + 缓存（Excel 与 PowerPoint 自己的写法）后才正常
  - **复制含图表的幻灯片要深拷贝图表部件**：两页共用一个 `chartN.xml` 时 PowerPoint 会拒绝打开整个文件
    （PowerPoint 自己复制时也会新建一份）。`duplicate_slide` 现在连图表部件与嵌入工作簿一起复制，
    实测复制后 PowerPoint 读到 6 张图表、两页数值一致
- `set_slide_layout`：换版式。幻灯片的版式关联**只由关系决定**（`slideN.xml.rels` 里 `…/slideLayout` 关系的 `Target`），
  所以改的是关系目标而不是幻灯片 XML；**不挪动也不删除形状**，若新版式没有对应占位符，文本会以自由形状继续显示
- `export_presentation`：PPTX → PDF（宿主内置引擎）。导出**前**先读一遍演示文稿，把「导不进去的东西」
  （动画、切换、音视频、嵌入对象、隐藏页）写进 `LOSSY_CONVERT` 警告；导出**后**再用插件自己的 PDF 读取器
  核对页数，与可见幻灯片数不一致时给 `PAGE_COUNT_MISMATCH`。实测 `deck.pptx`（4 页含 1 页隐藏）→ 3 页 PDF

**真实 PowerPoint 验证（版式）**：`npm run verify:pptx-layout` 读回 `LAYOUT_1 空白` / `LAYOUT_2 仅标题` / `LAYOUT_3 空白` / `LAYOUT_4 标题和文本`
（前两张是插件改的，后两张原样），且 `TITLE_1` 仍是「季度业务报告」——证明换版式不会丢形状文本。
导出的 PDF 用 **Word 自带的 PDF 解析器**独立复核：`REFLOW_PAGES 3`、`MATCH_COUNT 2`（`XLSX`），与插件读取结果一致。
- **读写共用同一个形状枚举器**：读和写各写一套遍历、或各自做一次类型细化（把 graphicFrame 细分出 table/chart），迟早会出现下标错位或类型不一致。细化只在一个地方做
- 只改目标形状的 `<p:txBody>`，其余形状、版式、母版、主题与动画一律不动
- 表格/图表等没有自己文本体的形状会被明确拒绝（不能把表格单元格的文本体误当成形状文本体）
- 形状含域（页码/日期）或超链接时默认**拒绝**，需 `allow_markup_loss=true`

**校验** —— `office_validate_pptx`：包结构、必需部件、幻灯片可解析、关系有效、图片/图表引用可解析、图表嵌入工作簿可解析、母版与主题保留，并如实列出需渲染才能判定的四项。

**真实 PowerPoint 验证**：改写标题与正文后用 PowerPoint 打开——标题显示新文本、正文三行新内容、其余幻灯片与尺寸（960pt = 13.33 英寸）不变、`PPT_OPEN_OK`。
增删与排序后再次用 PowerPoint 打开：幻灯片数与顺序（`本季度进展` / 空页 / `季度业务报告` / 空页）与插件读取结果**逐项吻合**。
复制幻灯片 + 插入文本框后（`npm run verify:pptx-slide-ops`）：PowerPoint 读到 5 页、副本页 3 个形状、第 3 个形状文本为新增文本框内容，且 `HASNOTES_3=no` / `HASNOTES_2=yes` —— 证明备注关系被正确丢弃而源页备注保留；
WPS 演示（`npm run verify:pptx-slide-ops-wps`）把它识别为 `type=17`、名称「文本框 4」、尺寸 390×90pt。

### DOCX 能力（阶段 3）

**读取** —— `office_read_docx`，五种视图：`summary` / `structure` / `paragraphs` / `tables` / `text`。

- **段落**：文本、样式名、大纲级别、对齐、是否列表项、是否含加粗片段
- **表格**：行列数与每格文本
- **页眉页脚**：文本与页码域识别
- **结构识别**：批注数、修订、域代码、目录、书签、内容控件、脚注尾注、嵌入对象、宏、数字签名、页面尺寸
- **修订语义**：`w:del`（已删除）不计入正文，`w:ins`（已插入）计入 —— 即 Word 显示修订后的可见文本

**写入** —— `office_insert_paragraph` / `office_update_paragraph` / `office_delete_paragraph` / `office_find_and_replace_docx` / `office_set_docx_header` / `office_set_docx_footer` / `office_update_docx_table_cell` / `office_insert_docx_image`。

- `update_paragraph` **保留段落属性**（样式、大纲级别、对齐、编号），只替换 run 内容 —— 因此填充模板不会破坏标题层级或编号
- 插入只新增一个 `<w:p>` 节点，不重建文档；`\n` 转软换行而非新段落
- 查找替换在 `<w:t>` 内进行；命中文本被格式拆散在多个 run 中时返回 `TEXT_SPLIT_ACROSS_RUNS` 警告而**不静默跳过**
- 页眉页脚写入可选附加**页码域**（fldChar 三段式，即 Word 自己写出来的形态，兼容性优于 fldSimple）；页码由 Word 按页自动更新，插件不写死
- 文档**没有**页眉页脚时自动新建并一次接线：部件 → 内容类型 Override → 关系 → `sectPr` 引用（少任何一环 Word 都会忽略它）
- 图片插入：类型按**真实字节**判定而非扩展名；未指定尺寸时用原始像素（96 DPI），只给宽度时按比例缩放；一并补齐媒体部件、内容类型与关系
- 删除段落是破坏性操作，需 `confirm=true`

**表格操作** —— `office_create_docx_table` / `office_set_docx_table_style` / `office_insert_docx_table_row` / `office_delete_docx_table_row` / `office_merge_docx_table_cells`：

- **建表**：`w:tblGrid` 给出每列宽度、每个单元格写 `w:tcW`（不写 Word 会按内容猜列宽）；
  首行可加粗并标记 `<w:trPr><w:tblHeader/></w:trPr>`（跨页重复表头）；
  表格成为正文最后一个内容元素时**补一个空段落**（Word 自己也这么做）
- **样式**：`w:tblStyle` 指向样式定义、`w:tblLook` 决定首行/首列/镶边行等条件格式是否生效 ——
  两件都要做；样式 ID 不存在时**补一个最小定义**（名称 + 单线边框），否则 Word 静默退回「普通表格」
- **插行从参照行复制格式**（`w:trPr` + 每个单元格的 `w:tcPr`）：只造一个不带单元格属性的 `<w:tr>`，
  Word 会按默认列宽**重排整张表** —— 列宽、边框、底纹都得跟着走
- **合并分两步**，缺一步 Word 就显示成没合并：横向写 `<w:gridSpan>` 并删掉被并入的 `<w:tc>`；
  纵向在首行写 `<w:vMerge w:val="restart"/>`、其余行写 `<w:vMerge/>`
- 被并入的文字按 Word 的行为**搬进左上角单元格**（换行分隔）并把原格**清空** ——
  不清空的话，日后「取消合并」会冒出重复文字
- 删行沿用与删段落同一套**保护性检查**（书签/批注锚点/修订/域代码默认拒绝）；表格只剩一行时拒绝删除

> **踩坑记录（同一个病根犯了两次）**：`XmlDoc` 的 `children` 是**改前快照**，删掉节点后再从里面找插入锚点，
> 拿到的可能正是刚被删的那一个 —— 实测写出**两个 `w:tblStyle`**（Word 里表现为样式来回跳）。
> 正确做法是**先取锚点、再删除**。另外 `w:tcPr` 的子元素必须按 schema 顺序插入
> （`cnfStyle → tcW → gridSpan → hMerge → vMerge → tcBorders → shd → …`），顺序错了 Word 报「文档内容有问题」。

**批注与修订（只读视图）** —— `office_read_docx` 的 `detail="comments"` / `detail="revisions"`：

- 批注：`word/comments.xml` 的 id、作者、缩写、日期、正文，以及它在正文里的**锚点段落下标**
  （由 `<w:commentReference w:id>` 反查所在段落）
- Word 2013+ 的**线程化批注**把「回复关系 / 已解决状态 / 人员标识」放在
  `commentsExtended.xml`、`commentsIds.xml`、`commentsExtensible.xml`、`people.xml` 里 —— 这些部件
  **只报告存在、不解析内容**，并在 `note` 里说明，而不是假装批注只有主部件里的字段
- 修订：`w:ins` / `w:del` / `w:moveFrom` / `w:moveTo` 逐条列出（作者、日期、文本、段落下标），
  外加 `rPrChange` / `pPrChange` 等**格式变更**计数
- 删除的文字在 `w:delText` 而不是 `w:t` 里，取错标签会得到一串空字符串；`w:ins`/`w:del` 出现在
  `w:pPr/w:rPr` 时表示**段落标记**本身被增删（合并/拆分段落），它天然没有文字 —— 单独标注为
  「删除（段落标记）」，不与「删掉一段文字」混为一谈

**Word 真机交叉验证**（`test/fixtures/word-complex.docx`，插件读数与 Word 读数逐项吻合）：

| 项目 | 插件 | Word COM |
|---|---|---|
| 批注 | `count 1`、作者 `审阅者`、正文 `批注：这里的风险项需要法务复核。` | `COMMENTS 1`、`COMMENT_1 author=审阅者 text=批注：这里的风险项需要法务复核。` |
| 修订 | `count 2`（1 插入 + 1 段落标记删除）、插入文本 `本段由修订模式插入。` | `REVISIONS 2`、`REVISION_1 type=1 text=本段由修订模式插入。`、`REVISION_2 type=2 text=`（Word 也读不出文字） |
| 书签 | `bookmarks()[0] name=RiskSection` | `BOOKMARK_1 name=RiskSection` |

**真实 Word 验证**：

| 样本 | 命令 | Word 读数 |
|---|---|---|
| 增删行 + 合并 | `npm run verify:docx-table` | `T1_ROWS 2 / T1_COLUMNS 2`、合并格 `新增A⏎1月`、被并入格 `<merged>` |
| 建表 + 套样式 | `npm run verify:docx-table-build` | `TABLES 2`、`T1_ROWS 3 / T1_COLUMNS 3`、`T1_STYLE 网格型`、单元格 `产品/金额/笔记本/显示器` |
| 纸张/方向/页边距 + 页码域 | `npm run verify:docx-layout` | `PAGE_WIDTH_PT 1190.55 / PAGE_HEIGHT_PT 841.9`（42×29.7cm）、`ORIENTATION 1`（横向）、`MARGIN_TOP_PT 56.7`（2cm）、`MARGIN_LEFT_PT 70.8`（2.5cm）、`FOOTER 第 1 页`、`FOOTER_FIELDS 1`、`FOOTER_FIELD_TYPES 33`（`wdFieldPage`） |
| 图片环绕 | `npm run verify:docx-wrap` | `INLINESHAPES 0` → `FLOATSHAPES 1`、`SHAPE1_WRAP 0`（`wdWrapSquare` 四周型）、`SHAPE1_POS 9,4.5`（114300 / 57150 EMU 精确落位）、`SHAPE1_SIZE 180x90`（= 240×120px，尺寸不受环绕方式影响） |

**图片环绕**（`office_set_docx_image_wrap`）：

- Word 里「行内图」与「浮动图」是**两种容器**：`<wp:inline>` 与 `<wp:anchor>`。转换要**连结束标签一起换**
  （只改起始标签会得到 `<wp:anchor>…</wp:inline>`，连自家解析器重新扫描时都报「XML 标签不匹配」）
- `wp:anchor` 的必需属性一个都不能少（`distT/distB/distL/distR/simplePos/relativeHeight/behindDoc/locked/layoutInCell/allowOverlap`），
  元素顺序也受 schema 约束：`simplePos → positionH → positionV → extent → effectExtent → 环绕元素 → docPr`。
  环绕元素插在 **`wp:docPr` 之前**、位置元素插在 **`wp:extent` 之前**
- 同一位置的三次零长插入按**登记顺序**生效，正好拼出 schema 要求的顺序（先 `simplePos`、再 `positionH`、再 `positionV`）
- 「衬于文字下方」与「浮于文字上方」在 XML 里都是 `<wp:wrapNone/>`，只差 `behindDoc`（同时把 z 序 `relativeHeight` 降到 0）。
  Word **没有**独立的「无环绕」选项，所以 `none` 是 `inFront` 的别名，读回统一报 `inFront`
- `tight`（紧密）与 `through`（穿越）需要 `wrapPolygon` 包围多边形 —— 要按图片实际轮廓算，属于需要渲染的能力，
  **明确拒绝**而不是伪造一个矩形
- 偏移用 EMU（1 pt = 12700 EMU），参照系可指定 `column/page/margin/character` 与 `paragraph/page/margin/line`

> **踩坑记录（只有插件级测试才抓得到的那种）**：适配器返回的 `offset_emu` 只在「浮动」分支里存在，
> 而工具的 `summaryOf` 无条件读 `extra.offset_emu.x` —— 于是「浮动 → 行内」这一步在
> `INTERNAL_ERROR: Cannot read properties of undefined (reading 'x')` 上炸掉。
> 适配器测试全绿（它们不经过 `summaryOf`），是插件级端到端用例把它揪出来的。
> 现在两个返回分支都带 `offset_emu`（行内为 `null`），`summaryOf` 也做了判空。

**页面设置与页码域**（`office_set_docx_page_layout` / `office_set_docx_margins` / `office_insert_docx_page_number`）：

- 页面尺寸与页边距都写在**末节**的 `w:sectPr` 上。只有 `w:body` 的**直接子** `sectPr` 才代表最后一节；
  段落内部的 `sectPr` 是分节符，改它只影响前面那一节 —— 所以选择器只看 body 直接子节点
- `w:pgSz` / `w:pgMar` 必须按 CT_SectPr 顺序落位（`pgSz → pgMar → …`），顺序错了 Word 报「文档内容有问题」
- 方向用**交换宽高**实现，横向才写 `w:orient="landscape"`（纵向是默认值）；尺寸以 twips 存储，对外用厘米
- 新建 `w:pgMar` 时**写齐七项**：缺属性的 `pgMar` 会让不同阅读器各取各的默认值，文档在不同机器上排版不一致
- 页码是 `PAGE` **域**（fldChar 三段式），由 Word 按页计算，插件不写死数字；工具会返回
  `FIELD_NEEDS_UPDATE` 警告提示打开后按 F9 刷新

> **踩坑记录（补丁区间同起点）**：给一个**没有 `w:pPr`** 的段落写页码时，第一次实现是
> 「先整段替换正文，再在正文起点插入 `<w:pPr>`」两个补丁。零长插入与替换补丁**起点相同**、不算重叠，
> 于是插入排在替换**之后** —— 结果段落里正文被写了两遍，旧内容留在段尾（生成 XML 一看就发现：
> `<w:pPr>` 排在 run 后面）。正确做法是**一次补丁同时写 `pPr` 与正文**；`pPr` 已存在时才分开改
> （正文补丁从 `pPr.end` 起，与 `pPr` 内部的补丁天然不重叠）。回归用例：
> 「页码域：pPr 在前、正文只有一份」。

**书签**（`office_insert_bookmark`）：书签是**一对**标记 —— `<w:bookmarkStart w:id w:name/>` 放在段落属性之后、`<w:bookmarkEnd w:id/>` 放在段落内容末尾；两处 `w:id` 必须一致且全文档唯一，重号会让目录、交叉引用与书签跳转失效。名称按 Word 的规则校验（字母含中文或下划线开头、最长 40 字符、**不区分大小写**地在文档内唯一）。空段落上两处插入落在同一偏移，按登记顺序生效（先 start 后 end），正好是正确顺序。

**Word 真机读数**（`npm run verify:docx-bookmark`）：`BOOKMARKS 2`、
`BOOKMARK_1 name=title_mark text=书签验证文档`、`BOOKMARK_2 name=正文_1 text=本文件由 dsh-exp-office 生成。加粗片段`
—— 中文书签名与它圈定的范围都被 Word 正确识别。

> **保护性检查（重要）**：段落里若含书签、批注锚点、修订标记、域代码或超链接，整段改写/删除会让它们失效 —— 部件还在，但引用消失，**Word 会静默丢弃对应功能**。此时工具默认**拒绝执行**并列出命中的标记，需显式传 `allow_markup_loss=true` 才继续。
>
> 这个检查来自一次真实教训：字节级测试断言「`comments.xml` 未变」判定通过，但用 Word 打开时批注从 1 条变成 0 条 —— 锚点被删了。改段落前可先用 `office_read_docx` 或适配器的 `paragraphSafety(index)` 看清哪些段落可安全改写。

**校验** —— `office_validate_docx`：ZIP 完整性、必需部件、XML 可解析性、关系目标存在性、图片/超链接/页眉页脚引用可解析、样式与编号引用有定义、目录与域代码状态、批注与修订保留、字体清单，并支持与基线做部件级差异比对。

> **不假装通过**：空白页、元素重叠、表格越界、页眉页脚重叠、字体替换这五项必须渲染成图片才能判定，工具会在 `not_checked` 中如实列出并给出原因，而不是返回一个「全部通过」。这个校验器第一次运行就抓出了我自己样本里引用未定义 `TableGrid` 样式的真实缺陷。

**比较** —— `office_compare_docx`：两份 Word 的**段落级差异**（只读，不改文件、不写修订标记）。

- 一级对齐用最长公共子序列，键是「样式 + 文字」：中间插入一段不会把后面所有段落都报成「修改」；
- 二级对齐在「删+增」区块内按相似度（文字相似度 ×2 + 样式相同 ×1）做保序最大权匹配，因此能区分
  「这段文字被改了」「这段只是换了样式」「这里真的插入了一段新内容」——直接按位置配对会把新插入的段落说成旧段落被改；
- 表格：表格个数、行列数变化、逐格文本变化（定位到第几张表第几行第几列）；另外比对段落样式集合与元数据字段；
- 段落数超过精确对齐的计算上限时**退化为按位置配对并明确告知**（`COMPARE_ALIGNMENT_FALLBACK`），不假装精确；
- 条目超过 `max_items` 时截断，但 `total_*` 仍是真实总数。

> 实测（真实 Word 交叉验证）：`report.docx`（Word 读到 9 段）与 `report-edited.docx`（Word 读到 10 段、第 2 段是「本段由插件插入」）→
> 插件比较结果恰好是「新增 1 段『本段由插件插入』+ 1 处正文文字修改」，与 Word 自己的说法一致。

## 转 PDF

`office_convert_document` 把 `docx` / `xlsx` / `pptx`（含旧版二进制格式）转为 PDF，使用**宿主内置的 LibreOffice 引擎**，不需要本机安装 Office。

- **不新增依赖**：`@deepseek-ai/libreoffice-kit` 作为**可选** peerDependency 延迟解析；解析不到时该工具返回 `OFFICE_NOT_INSTALLED`，其余能力不受影响
- **不自建进程管理**：引擎自带超时、取消、字节上限与输出路径保护，自己再包一层只会更差
- **返回缺失字体清单**：`missing_fonts` 是判断「是否发生字体替换」的唯一可靠依据，读 XML 无法回答 —— 它补上了 `validate_docx` 里那项 `not_checked`
- 转换不会覆盖已存在的输出文件；支持取消

**真实 Word 验证**：插件写入的 .docx 被 Word 打开后——插入段被识别为**「标题 1」样式**、查找替换生效、表格完好、**页脚页码域仍渲染为「第1页」**、页眉与文档属性未变。设置页眉页脚后 Word 显示新页眉文本、页脚 PAGE 域渲染为实际页码。插入图片后 Word 报告 `InlineShapes: 1` 且尺寸精确（240px → 180pt）。

**真实 Word 复杂样本回归**：用 Word 创作含**批注、未接受修订、目录域、书签、脚注、超链接**的文档，插件编辑后 Word 复核的特性数量与原样本**完全一致**（批注 1、修订 2、书签 1、目录 1、脚注 1），且除 `word/document.xml` 外所有部件逐字节不变。

## 跨软件交叉验证

按需求 §九.3「不能只验证文件能否打开」，下表是**同一份插件输出**在多个真实软件里的复核结果。

| 输出 | 校验方 | 结果 |
|---|---|---|
| `.xlsx`（含图表/条件格式/合并/公式） | Microsoft Excel 16 | 数值、图表 1、条件格式 1、合并区域全部保留 |
| `.xlsx`（含样式） | WPS 表格 12 | 打开正常，样式保留 |
| `.docx`（含批注/修订/目录/书签/脚注/超链接） | Microsoft Word | 批注 1、修订 2、书签 1、目录 1、脚注 1 |
| 同上 | **WPS 文字 12** | **批注 1、修订 2、书签 1、域 4 —— 与 Word 逐项吻合** |
| `.docx`（插件写入的页眉页脚与图片） | Microsoft Word | 页眉正确、页码域渲染为「第 1」、`InlineShapes: 1` |
| `.pptx`（插件改写文本后） | Microsoft PowerPoint 16 | 标题与正文为新内容，尺寸 960pt |
| 同上 | **WPS 演示 12** | **与 PowerPoint 逐项吻合** |
| `.pptx`（插件增删排序后） | Microsoft PowerPoint 16 | 顺序：本季度进展 / 空页 / 季度业务报告 / 空页 |
| 同上 | **WPS 演示 12** | **顺序与标题完全一致** |
| `.pptx`（插件插入图片后） | Microsoft PowerPoint 16 | 形状数 2→3，`type=13`（图片）命名为 `Picture 4`，尺寸 **300×150pt**（400px/96dpi 精确换算） |
| `.pptx`（插件插入表格后） | **WPS 演示 12** | 第 1 页 3 个形状，表格 `type=19` 命名 `Table 4`、尺寸 360×120pt、**4 行 3 列**，与 PowerPoint 与插件自身读取结果一致 |
| `.pptx`（插件插入图片后） | **WPS 演示 12** | 形状类型/名称/尺寸逐项复核（见 `test/wps-office-check.ps1 -Kind presentation` 的 `SHAPE_*` 输出） |
| `.pdf`（转自 docx / pptx） | PDF 结构检查 | 合法 PDF、页数正确（隐藏页未导出）、字体缺失 0 |
| `.pdf`（插件旋转/重排/删除后） | **Microsoft Word 的 PDF 重排**（独立解析器） | 重排后文本顺序变为「本季度进展 → 表格页 → 季度业务报告」；删除后剩 2 页且被删页文本消失；旋转后文件仍被正常解析 |
| `.pdf`（插件加页码 / 水平水印后） | **Microsoft Word 的 PDF 重排** | 页码被读进**页脚故事**（`FOOTER: 1 / 3`）、水平水印被读进**页眉故事**（`HEADER: HORIZONTAL MARK`），原有中文全部完好 |
| `.pptx`（复制幻灯片 + 插入文本框后） | Microsoft PowerPoint 16 / **WPS 演示 12** | PowerPoint：5 页、副本页 3 个形状、第 3 个形状文本 = 新增文本框内容、`HASNOTES_2=yes` / `HASNOTES_3=no`；WPS 演示：`type=17` 文本框 390×90pt |
| `.docx`（A3 横向 + 页边距 + 页码域） | Microsoft Word / **WPS 文字 12** | Word：42×29.7cm、`ORIENTATION 1`、页边距 2/2.5cm、`FOOTER 第 1 页`、`FOOTER_FIELDS 1`（类型 33 = `wdFieldPage`）；WPS 文字正常打开 |
| `.docx`（图片改为四周型环绕） | Microsoft Word | `INLINESHAPES 0` → `FLOATSHAPES 1`、`SHAPE1_WRAP 0`（`wdWrapSquare`）、`SHAPE1_POS 9,4.5`、尺寸不变 |
| `.docx`（插入中文名书签） | Microsoft Word | `BOOKMARKS 2`、`BOOKMARK_1 name=title_mark text=书签验证文档`、`BOOKMARK_2 name=正文_1 …` |
| `.docx`（批注/修订只读视图） | Microsoft Word | `COMMENTS 1` + 批注作者与正文、`REVISIONS 2`（`type=1` 插入文本、`type=2` 段落标记）**与插件读数逐项吻合** |
| `.pptx`（插入原生图表后） | Microsoft PowerPoint 16 | `charttype=51/57/65/5` 四种类型全部识别、系列数与数值与插件写入**逐项吻合**（`Values=120,150,180` 等）、标题存在、无修复提示 |
| 同上（图表） | **WPS 演示 12** | 9 页全部正常打开（每张图表页 1 个图形框），`WPS_OFFICE_OK` |
| `.docx`（导出副本） | Microsoft Word | 副本与源文件**逐字节一致**；Word 打开后段落 9 / 表格 1 / 页脚域类型 33 与源文件读数相同 |
| `.pptx`（**从零生成**） | Microsoft PowerPoint 16 / **WPS 演示 12** | PowerPoint：`SLIDES 1`、`SLIDE_WIDTH_PT 960`、`DESIGN_NAME office-plugin`、`LAYOUT_1 标题幻灯片`；编辑后 `SLIDES 2` + `charttype=51`（数值 3,1,2）且无修复提示。WPS 演示：`SLIDES 2`、`TABLE_2_2 rows=2 cols=2`、`WPS_OFFICE_OK` |
| `.pptx`（换主题后） | Microsoft PowerPoint 16 | `DESIGN_NAME` 可读、`ACCENT1_RGB` 从默认值变成 **255（红）** —— 证明换主题真的生效，而不只是「文件能打开」 |
| `.pptx`（切换版式后） | Microsoft PowerPoint 16 | `LAYOUT_1 空白` / `LAYOUT_2 仅标题` / `LAYOUT_3 空白` / `LAYOUT_4 标题和文本`，且 `TITLE_1` 仍是原标题 |
| `.pdf`（`export_presentation` 导出） | 插件 PDF 读取器 + **Word PDF 重排** | 4 张幻灯片（1 张隐藏）→ 3 页；Word 独立读到 `REFLOW_PAGES 3`、`MATCH_COUNT 2`（`XLSX`） |
| `.pdf`（`office_search_pdf` 全文搜索） | **Word PDF 重排** | 搜 `XLSX` 命中 2 处（第 2、3 页），Word 独立计数 `MATCH_COUNT 2` |
| `.pdf`（加中文便签批注后） | 插件 PDF 读取器 + **Word PDF 重排** | 读取器逐字段读回（`Text` / 中文内容 / 作者 / rect / `flags=4` / `open`）、`structure.annotations 1`、结构校验通过；Word 仍能正常解析（`REFLOW_PAGES 2 / WORDS 43`），增量更新不触碰原有字节 |

复现脚本：`npm run verify:excel` / `verify:wps` / `verify:table` / `verify:chart` / `verify:docx-table` / `verify:docx-table-build` / `verify:docx-style` /
`verify:docx-layout` / `verify:docx-wrap` / `verify:docx-bookmark` / `verify:word` / `verify:pptx` / `verify:pptx-image` / `verify:pptx-slide-ops` / `verify:pptx-layout` / `verify:pptx-charts` / `verify:pptx-theme` / `verify:pptx-created` / `verify:docx-export` / `verify:pdf`，
以及 `test/wps-office-check.ps1 -Kind writer|presentation` 与
`pwsh -File test/pdf-open-check.ps1 -Path <pdf> -Match <ASCII关键词>`。

### 打开 → 重新保存 → 再次打开（需求 §十八.6）

只「能打开」还不够：Office/WPS 必须能把插件输出当成正常文档重新保存，且重存后再次打开内容不变。

| 文件 | 软件 | 往返结果 |
|---|---|---|
| `.xlsx`（图表 + 条件格式 + 合并 + 公式） | Microsoft Excel 16 | 重存后再次打开：数值、图表 1、条件格式 1 **全部保留** |
| 同上 | WPS 表格 12 | 往返通过 |
| `.docx`（批注 + 修订 + 目录 + 书签 + 脚注） | Microsoft Word | 重存后再次打开：**批注 1、修订 2、书签 1、目录 1、脚注 1 全部保留** |
| `.pptx`（含插件插入的图片） | Microsoft PowerPoint 16 | 重存后再次打开：幻灯片 4、第 2 页形状 3、标题完好 |

## 性能

对照需求 §14.4 参考目标（Node 24 / i5-10300H，详见 [技术方案 §六](docs/阶段0-技术方案.md)）：

| 场景 | 实测 | 目标 |
|---|---|---|
| 读取小型 XLSX（1000×10，10010 单元格） | **75 ms**（打开 1.1 + 读全表 73.8） | ≤ 2000 ms ✅ |
| 修改小型 XLSX（改 2 格并保存） | **149 ms** | ≤ 3000 ms ✅ |
| 仅解析 ZIP 中央目录（7 个部件） | **0.1 ms** | ≤ 500 ms ✅ |
| 20 万单元格表**读 100 行** | **118 ms** | — |
| 50 万单元格表**读 100 行** | **162 ms**（其中约 130 ms 是解压工作表部件） | — |
| 50 万单元格表读全表（50 万格） | **2.3 s**（RSS 峰值含前序步骤，单步堆增量近 0） | — |
| 10 万单元格表改 1 格并保存 | **1.02 s**（改 20 万格 1.7 s） | — |

`npm run test:perf:large` 复现（最后一次运行：**8 项达标 / 0 项未达标**）。

### 区域读取是真的按区域读（阶段 7）

`office_read_range` **不再把整张表建成 DOM**，而是直接扫工作表部件的原始字节：

- 扫到区域最后一行之后的第一个 `<row>` **立即停止**，因此读 100 行的代价只与这 100 行有关；
- 共享字符串表与样式表只在真遇到对应单元格时才解析（纯数字区域根本不碰 `sharedStrings.xml`）；
- 确定性证据（不是计时）：样本第 1000 行故意放一个非法引用，读 `A1:B2` 不报错、整表读取必报错、读 `A999:B1000` 报错（`npm run test:scan`）；
- 等价性：套件里保留了一份**独立的 DOM 参照实现**（改造前的算法），对六种单元格类型、缺 `r` 的行/格、自闭合行、共享公式、CDATA、前置空格、越界共享下标与 13 个区域逐格比对。

**诚实的边界**：

- 首次读取某个工作表仍需**解压整个工作表部件**（50 MB 级表部件约 0.5–1 s），之后同进程内复用缓存。真正的流式（边解压边解析）属于后续工作。
- 单元格解析约 **3 µs/格**（扫描 + 属性解析），50 万格读全表约 2.3 s。这比旧路径（建 DOM）快且省内存，但不是零成本。
- **写入路径仍需把整张表解析成 DOM**（这是「只改目标区间、其余字节不变」的前提）：实测 5 万格 0.46 s / 150 MB、10 万格 0.94 s / 265 MB、20 万格 1.70 s / 495 MB。因此单表写入有硬上限 **30 万单元格**（`maxEditableCells` 可调）：超过就给 `MEMORY_LIMIT` 与出路，而不是把宿主进程拖进 GC（50 万格实测 >10 分钟不收敛）。
- 读取上限：单次区域 **20 万格**（`MAX_RANGE_CELLS`）、整表读取 **50 万格**（`MAX_SHEET_CELLS`）；都是**在分配内存之前**拒绝，并给出可操作建议。

## 已知限制

- **公式不重算**：写入公式时请同时给出缓存值，或让用户在 Excel/WPS 中打开重算。会返回 `FORMULA_NOT_RECALCULATED` 提示。
- **行列移位不重写公式引用**：插入/删除行列只重排行号与单元格引用，公式文本、合并区域、条件格式中的引用不会自动调整，会明确返回警告。
- **大文件**：读取按区域扫描（见上），但**写入**要把整张表建成 DOM，单表 >30 万单元格默认拒绝；单文件上限 512 MB。
- **不执行宏**：`.xlsm` 的 VBA 按原字节保留但从不执行。
- **CSV 导出是有损的**：公式、样式、图表、批注与其它工作表都不会保留，会返回 `LOSSY_CONVERT` 警告。
- **样式能力限于字体/填充/数字格式/水平对齐**，不支持边框、垂直对齐、条件格式创建。
- **不支持**：图表创建、透视表修改、PDF 表单填写/批注/合并拆分、OCR、PDF → Office 反向转换（阶段 5–7 计划）。
- **`<c><v/></c>`（空值元素）读出来是 `0` 而不是 `null`**：`Number('')` 的历史行为，改造前后一致；要改成 `null` 属于语义变更，需要单独决策（已记在任务清单）。
- **本地引擎联动是「单用户桌面」级别**：默认关闭、只清理自己启动的进程、不做受限账户/沙箱隔离（开发要求建议的隔离需要运维层面配合）；视觉是否变化也没有回归测试（需要渲染成图片才能判定）。
- **页码/日期等域不写死数字**：域的值由 Word 计算，插件写入的是域代码与一个缓存值；请按 F9（或打印预览）刷新。WPS 文字能正常打开带域文档，但 WPS COM 的 `Fields.Count` **只统计正文域**，页眉页脚里的域不出现在该集合中 —— 那里读到 0 不代表域没写进去（Word 的 `Section.Headers/Footers.Range.Fields` 读到 1，类型 33 即 `wdFieldPage`）。
- **PDF 写入限于页面级操作与叠加文字**（旋转/删除/重排、水印、页码）：不做**原地改写已有文字**、表单填写与内容流重写；要安全地做到那些需要解析并重建内容流与字体子集，不是这一阶段能诚实交付的。
- **加密文件**：正确识别并返回 `PASSWORD_REQUIRED`，不解密。

## 部署要求

- Node.js ≥ 22.19.0（依赖内置 `zlib.crc32` 与 `node:crypto`）
- DSH ≥ 0.1.0-rc.6
- 无第三方运行时依赖，无网络访问需求
- 插件配置：`workspaceRoot`（默认进程工作目录）、`maxFileBytes`、`allowOverwriteOriginals`、`maxEditableCells`、`allowLocalAutomation`（默认 false）、`automationTimeoutMs`、`powershellPath`

## 本地 Office/WPS 联动（阶段 6）

三个工具，**默认全部关闭**（`allowLocalAutomation` 必须显式为 `true`）：

| 工具 | 做什么 | 会不会动源文件 |
|---|---|---|
| `office_detect_engines` | 读注册表列出本机 Office/WPS 组件与版本；`probe_com=true` 时实测 COM 可用性（需授权） | 不动任何文件 |
| `office_recalculate` | 用真实引擎打开**副本**、重算（Excel 公式缓存、Word 域）后另存为新文件 | **不会**：输入只读打开，产物写新路径 |
| `office_rerender` | 用真实引擎打开**副本**并整份重写后另存为新文件（修「Office 能打开但提示修复」的文件） | **不会** |

**安全契约**（开发要求 §十二.4 逐条落地，另有 [docs/本地联动安全说明.md](docs/本地联动安全说明.md)）：

| 要求 | 做法 |
|---|---|
| 独立进程 | 引擎在它自己的进程里跑；插件只通过结果 JSON 读结论，不把 COM 对象暴露给上层 |
| 禁止宏自动执行 | `AutomationSecurity = 3`（msoAutomationSecurityForceDisable）；宏文件按宏格式重存但**绝不执行宏** |
| 禁止外部链接/模板自动更新 | `AskToUpdateLinks = false`、`Open(UpdateLinks:=0)`；不调用 `AddIns`/`AttachedTemplate` |
| 设置脚本执行超时 | 默认 120 s；到点先杀 PowerShell 进程树，再按状态文件补一次 cleanup |
| 防止弹窗阻塞 | `DisplayAlerts = false`、`-NonInteractive`、尽量隐藏窗口 |
| 检测进程异常退出 | 脚本结束返回退出码与 JSON；插件按「结果文件是否存在」判定，绝不当成功 |
| 任务结束清理进程 | 开始时快照引擎 PID，结束时**只杀新增的那些**（不碰用户自己开着的 Office/WPS），并回报 `killed_pids` / `leftover_pids` |
| 文件锁 | 调用期间对源文件加插件自己的文件锁，避免与其它任务同时处理同一文件 |
| 审计 | 每次调用追加一行 JSONL（动作/引擎/文件名/字节数/耗时/错误），**不记绝对路径、不记文档内容** |

真机验证（`npm run verify:automation`，本机 Office 16.0.17932 + WPS 12.1）：故意把 `=1+2` 的缓存值写成 999 →
Excel 重算后读到 **3**、WPS 重算后同样读到 **3**；Word / PowerPoint 重写后的产物都能被自身读取器解析且规模一致；
每一次都断言**源文件逐字节未变**、**没有留下本次启动的引擎进程**。

> 明确没做的：受限账户/沙箱隔离（需要运维层面配合）、视觉回归（要渲染成图片）、把 COM 对象直接暴露给 Agent（开发要求明令禁止）。

## 从零生成 PDF（阶段 5 收尾）

`office_create_pdf` **不依赖模板与任何第三方库**：页面树、内容流、字体资源、xref 全部自己写，
并同时写入 `/ToUnicode`，因此生成的文本**可以被提取**（自己与第三方解析器都能读回）。

| 能力 | 说明 |
|---|---|
| 排版 | 自动折行、自动分页、`\f` 前缀强制分页；纸张 A3/A4/A5/Letter/Legal 或自定义点尺寸；纵向/横向；页边距/字号/行距/字体可调 |
| 拉丁文本 | PDF 标准 14 字体（Helvetica / Times / Courier 各四体），**无需嵌入字体**，任何阅读器都能打开 |
| **中文等非拉丁文本** | **自动嵌入子集化的 TrueType 字体**：解析本机字体（`.ttf`/`.ttc`）→ 只保留用到的字形 → 写 Type0 / CIDFontType2 / FontDescriptor / FontFile2 / CIDToGIDMap / ToUnicode。实测含 119 个汉字的 2 页 PDF 只有 **28.7 KB** |
| 文本 | 括号/反斜杠/高位字符按 PDF 语法转义；两类字体都写 `/ToUnicode`，因此**文本都能被提取** |
| 元数据 | Title/Author/Subject/Keywords/Creator/Producer；非 ASCII 值按 UTF-16BE + BOM 写 |
| 自校验 | 写完立刻用自己的读取器读回：页数不对或第一行文本读不回就**不留文件** |

**字体子集是怎么做的（关键设计）**：PDF 的 `CIDFontType2` 用 `/CIDToGIDMap` 把 CID 映射到字形编号，
所以**只要子集里字形编号不变，就不需要重排编号** —— 重排编号正是子集化最容易出错的地方。
实现是保留 `loca` 的长度、把未用到的字形写成零长度（TrueType 里的「空字形」），
并按 maxp / hhea / hmtx 同步裁掉编号更高的部分；复合字形（带重音的拉丁字母、部分汉字）会递归闭包进来。

**实测（第三方解析器独立核对）**：`npm run verify:pdf-cjk` 生成一份含中文标题、中文正文、中英混排、
第二页与 ASCII 标记的 PDF，再用 **Word 自带的 PDF 解析器**读回来逐项比对 —— 5/5 全部读到。
这一步不是形式：本轮就是它抓出了「`/W` 数组被当成 PDF 名称写出去」的缺陷 —— **自读完全正常，Word 读出 0 个字**。

**字体从哪来**：默认按顺序找本机 `simhei.ttf` → `Deng.ttf` → `simfang.ttf` → `simkai.ttf` → `msyh.ttc` →
`simsun.ttc` → `NotoSansSC-VF.ttf`；也可以用 `cjk_font_path` 指定。**没有任何可用字体时明确报错**
（列出试过的路径），而不是画乱码。

**仍然不做**：富文本（行内混排样式）、图片、矢量图形、表格框线与页码。
中文**水印 / 页码**已经做了 —— `office_add_pdf_watermark` / `office_add_pdf_page_numbers` 复用同一套字体嵌入机制（见「叠加式写入」一节）。

## 表格提取（阶段 5 收尾）

`office_extract_pdf_tables` 从 PDF 里还原表格：**按文本片段的位置推断**——
先把 y 坐标的间隔推成行距、据此聚类成行，再按 x 坐标聚类成列，行 × 列得到单元格文本。
不依赖表格框线（很多 PDF 根本没有框线）。

| 细节 | 做法 |
|---|---|
| 行容差 | **自动推导**：同一行里汉字与数字的基线常差几个点（实测 4.9 pt），写死容差会把一行拆成两行 |
| 列间隔 | 用估算字宽（CJK 1 em、其余 0.5 em）判断「这是词间距还是列间隔」；正文里的词间距不会被误判成列 |
| 靠空格对齐的表格 | 一个字符串里用连续空格分列的（等宽排版），会按两空格以上切成虚拟列 —— 包括**本插件自己生成的 PDF** |
| 多张表 | 连续的「表格行」组成一张表；正文行会中断成表过程，因此同一页的多张表会分别报出 |
| 诚实报告 | 返回值里带 `method`、`approximation`（近似之处）与 `not_done`（跨页合并、合并单元格、扫描件） |

实测（真实 PDF，均由 LibreOffice 从带表格的 DOCX 导出）：

| 样本 | 提取结果 |
|---|---|
| `sample.pdf`（2×2 表） | `[["月份","金额"],["1月","12000"]]` ✅ 与源文档一致 |
| `sample-multipage.pdf`（第 3 页 3×3 表） | `[["模块","状态","测试数"],["XLSX","已完成","46"],["DOCX","已完成","73"]]` ✅ 只在第 3 页报表 |
| 无表格的 PDF（含本插件生成的） | 报 0 张表，**不把正文误判成表格** |
| 独立交叉验证 | Word 自带的 PDF 解析器在同一份 PDF 里同样读到 `12000`（`MATCH_COUNT 1`） |

**明确不做**：跨页表格自动合并、合并单元格的跨行跨列还原、扫描件与纯图片表格（需要 OCR，当前不支持）、
单元格内换行与富文本样式。

## 并发与文件锁

同一文件被并发修改时的行为是**明确契约**，有专门的测试套件（`npm run test:concurrency`）盯着：

| 场景 | 行为 |
|---|---|
| 两个写入同时打同一文件 | **串行化**：第二个等锁（默认最多等 5 秒），两边都成功、结果都落盘、文件仍合法 |
| 别的任务正持锁 | 返回 `FILE_LOCKED`，**文件逐字节不变**（宁可拒绝也不写坏） |
| 进程异常退出留下残留锁 | 超过存活时间（10 分钟）自动回收，不会永久卡死 |
| 并发读 | 不加排他锁，互不阻塞 |
| 成功路径 | 不留锁文件 |

## 安装 / 升级 / 卸载

**安装**（普通 npm 包，也可以作为 profile 的 `file:` 依赖）：

```bash
npm install dsh-exp-office            # 从 registry 安装
npm install /path/to/dsh-exp-office   # 或从目录安装
```

**升级**：

```bash
npm install dsh-exp-office@latest     # registry 升级
npm install /path/to/new-version      # 从新目录升级（会替换旧版本）
```

> 注意：`file:<目录>` 形式的依赖**同版本不会触发替换**（npm 认为已是最新），升级时请递增 `version`；
> 另外这种依赖装出来是**符号链接**指向源目录 —— 所以「改安装副本里的文件」实际改的是源码，别这么干。

**卸载**：`npm uninstall dsh-exp-office`。卸载后目录被移除，不留锁文件、暂存目录或临时文件。

**演练**：`npm run verify:lifecycle` 会在临时目录里完整走一遍「安装 → 升级（版本号 +1）→ 卸载」，
并断言**全程未改动工作区源码**（比对 `lib/` 与三个清单的指纹）。

**发布包自检**：`npm run verify:pack` 打 `npm pack` → 校验文件清单（必需文件齐全、没混入测试与运行残留）
→ 在临时目录安装 → 用**安装后的那份代码**跑真实 Cordis 装载并核对工具数。

## 目录结构

```
lib/                      18 个模块，无第三方运行时依赖
  index.js        Cordis 插件入口：解析 defineTool 并注册工具
  capabilities.js 能力清单（插件版本/协议版本/已实现与未实现能力）
  tools.js        90 个工具定义 + 统一响应信封 + 预览-执行-回滚五件套
  ooxml.js        格式无关底座：ZIP 容器、字节级最小修改 XML 引擎、OPC 容器操作
  workspace.js    路径安全 / 真实类型检测 / 哈希 / 文件锁 / 临时目录 / 事务回滚
  errors.js       错误码体系 + 结构化响应
  define-tool.js  defineTool 与宿主模块的分层解析
  xlsx.js         XLSX 适配器（读写/工作表/查找替换/合并/行列/样式/导出/校验）
  docx.js         DOCX 适配器（读写/段落/表格/页眉页脚/图片/文档比较/保护性检查/校验）
  pptx.js         PPTX 适配器（读取/形状文本/幻灯片增删排序/图片表格/图表/版式主题/校验）
  pptx-template.js 空白演示文稿模板生成器（主题/母版/版式程序化生成）
  pdf.js          PDF 适配器（读取/文本/页面操作/叠加写入/重写·拆分·合并/图片提取）
  image.js        格式无关图像能力（PNG 编码）
  convert.js      Office → PDF（延迟解析宿主内置 LibreOffice 引擎）
  automation.js   本地 Office/WPS 联动（开关 / 超时 / 进程回收 / 审计）


  office-automation.ps1  唯一与本地 Office/WPS 通话的脚本（固定模板、纯 ASCII、参数化）

test/                     57 个文件
  all.mjs               总入口（npm test）
  engine.test.mjs       引擎：ZIP / XML 最小修改 / 安全防护
  xlsx.test.mjs         XLSX 适配器
  xlsx-scan.test.mjs    XLSX 区域读取的字节扫描路径（早停证明 / 上限 / 与 DOM 参照逐格等价）
  docx.test.mjs         DOCX 适配器（读取 / 写入 / 表格 / 图片 / 样式 / 页面设置 / 校验）
  pptx.test.mjs         PPTX 适配器（读取 / 写入 / 幻灯片管理 / 复制 / 文本框）
  pdf.test.mjs          PDF 适配器（读取 / 文本提取 / 页面级写操作）
  fidelity.test.mjs     保真度：真实 Excel 多特性样本
  word-fidelity.test.mjs 保真度：真实 Word 复杂样本（批注/修订/目录）
  macro.test.mjs        宏启用工作簿（检测与字节保留）
  plugin.test.mjs       插件集成：注册 / 端到端 / 事务 / 任务五件套
  convert.test.mjs      转 PDF（真实引擎，单独套件）
  perf.test.mjs         性能基准（对照文档 §14.4 目标）
  concurrency.test.mjs  并发与文件锁（串行化 / FILE_LOCKED / 残留锁回收）
  automation.test.mjs   本地联动的 Node 侧（开关 / 超时 / 进程回收 / 审计，用桩脚本）
  automation-check.mjs  本地联动真机验证（真实 Excel/WPS/Word/PowerPoint，会启动本机 Office）
  make-xlsx-large-fixture.mjs 大样本生成器（直接拼 XML，不走 DOM）

  read-cells-check.mjs  逐格读数，用于与 Excel/WPS 读数对照
  pack-check.mjs        发布包自检（打包 → 校验清单 → 临时安装 → 真实装载）
  profile-lifecycle-check.mjs 安装 → 升级 → 卸载 演练（含源码指纹保护）
  release-privacy-check.mjs   推前隐私扫描（本地模式表 + 内置规则）
  make-*-fixture.*      样本生成（Excel / Word / PPTX / 宏 / 图片 / PDF，需本机 Office 或宿主引擎）
  profile-boot-check.mjs profile 启动自检（副本漂移 + 真实 Cordis 装载，不开服务）
  image-decode-check.mjs 用第三方解码器（sharp）验证导出的图片能打开、内容一致
  extract-images-check.mjs 图片提取的端到端检查（提取 → 落盘 → 报告）
  *-open-check.ps1      真实软件验证：Excel / WPS / Word / PowerPoint / PDF（Word PDF 重排）
  wps-office-check.ps1  WPS 文字与 WPS 演示交叉验证（含逐形状类型/尺寸/表格行列）
```

## 安装到 DSH profile

插件是普通的 DSH 组合包（bundle）：profile 的 `package.json` 里加依赖与 `dsh.profile.bundles` 条目，
包自带的 `cordis.patch.yml`（`dsh.bundle.patch`）负责把插件行插进配置树，**不需要改 profile 的 patch 层**。

```bash
# 以 web profile 为例（本地工作区作为 file: 依赖）
# 1) profiles/web/package.json: dependencies 增加
#      "dsh-exp-office": "file:E:/deepseek harness/插件工作区/dsh-exp-office"
#    并把它加进 dsh.profile.bundles
# 2) 首次安装：**在 DSH 停止时**于 profile 目录执行一次 pnpm install
cd %USERPROFILE%/.dsh/profiles/web && pnpm install
# 3) 检查组合后的配置树（只组合配置，不会 apply 插件，也拦不住契约错误）
dsh --profile web --dump-config | Select-String dsh-exp-office
# 4) 之后改源码只需同步副本 —— 用复制，不要用 pnpm
npm run sync:profile     # 只复制本包的 lib/ 与清单文件到安装副本
npm run verify:boot      # 漂移 + 真实装载 + 工具数
# 5) 重启 profile 后生效
```

`npm run verify:boot` 做三件事：比对安装副本与工作区的逐字节一致性、从 profile 按包名解析插件、
用 profile 那份**真实 Cordis** 装载并 await，最后断言工具数与 `dsh.plugin.json` 一致。

> **重要：DSH 运行期间不要在 profile 里跑 `pnpm install`。**
>
> `pnpm install` 会按 `package.json` **收敛整棵 node_modules**：它不只处理本包，还会去替换别的插件，
> 其中包括正被运行中的服务占用的原生包（实测报
> `failed to remove existing directory … 拒绝访问 (os error 5)` 并以退出码 1 结束）。
> 结果是 **profile 里其它插件也可能被牵连**（实测 `dsh-context`、`dsh-ponytail-skills`、
> `dsh-whale-widget`、`dsh-better-sidebar` 等目录被改写，随后整个插件树罢工）。
>
> 因此日常只用 `npm run sync:profile`（纯文件复制，只碰 `node_modules/dsh-exp-office/` 一个目录）；
> 需要 `pnpm install` 时请**先停掉 DSH**（它同时会补齐被占用而没能换位的包）。

### Cordis 插件契约（踩过一次事故，写在这里）

> **`apply` 的返回值只能是清理函数 / `null` / `undefined`（或 resolve 成它们的 Promise/迭代器）。**
>
> 2026-09-23 早间事故：`apply` 末尾 `return { registered, manifest }` 返回了普通对象摘要，
> Cordis 的 `safeCollect` 抛 `TypeError: Invalid effect`，**整个插件树加载失败、`dsh web` 起不来**。
> `dsh --profile web --dump-config` 当时是通过的 —— 它只组合配置树，不 apply 插件。
> 要暴露摘要请写 `ctx.logger.info`，不要 return。
>
> 现在有三道闸门守着：`apply` 返回值契约断言、**真实 Cordis 加载器装载测试**、
> 以及一条**反例测试**（apply resolve 成普通对象时加载器必须拒绝）——反例保证前两道不是碰巧通过。

已实测：`dsh --profile web --dump-config` 输出里出现 `id: dsh-exp-office`；`npm run verify:boot`
显示「安装副本与工作区逐字节一致」且「真实装载注册 90 个工具，与清单一致」。

> **注意**：`pnpm install` 会按 `package.json` **收敛** profile 的 `node_modules` ——
> profile 里那些「曾经装过、但已不在 `package.json` 里」的遗留依赖会被清理掉。
> 本插件安装时清理掉了 443 个这类遗留包（来源是两次被中断的插件安装留下的残骸，
> 见 `.plugin-manager/logs/*/pnpm.log` 里 `Command failed with exit code 3221225786`）。
> 安装前请先备份 `package.json` 与 `pnpm-lock.yaml`。

## 许可

MIT
