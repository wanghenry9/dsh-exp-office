# 用真实 Microsoft PowerPoint 创作一份「多特性」演示文稿，用于 PPTX 适配器测试与保真度回归。
# 覆盖：标题页、项目符号正文、表格、备注、隐藏幻灯片、自定义主题色。
#
# 注意：本机 PowerShell 的 COM 绑定对「循环里用变量拼地址」会失败，
# 因此这里统一使用字面地址与显式索引。
param([string]$OutPath = "$PSScriptRoot\fixtures\deck.pptx")

$ErrorActionPreference = 'Stop'
$resolved = [System.IO.Path]::GetFullPath($OutPath)
$dir = Split-Path -Parent $resolved
if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Force }

$app = $null
$pres = $null
try {
  $app = New-Object -ComObject PowerPoint.Application
  # msoFalse = 0：不显示窗口
  $pres = $app.Presentations.Add(0)

  # ---- 第 1 张：标题页（ppLayoutTitle = 1）----
  $s1 = $pres.Slides.Add(1, 1)
  $s1.Shapes.Title.TextFrame.TextRange.Text = "季度业务报告"
  $s1.Shapes.Item(2).TextFrame.TextRange.Text = "技术方案与实施进展 · 2026 Q3"

  # ---- 第 2 张：项目符号正文（ppLayoutText = 2）----
  $s2 = $pres.Slides.Add(2, 2)
  $s2.Shapes.Title.TextFrame.TextRange.Text = "本季度进展"
  $body = $s2.Shapes.Item(2).TextFrame.TextRange
  $body.Text = "完成 XLSX 读写与保真度验证`r完成 DOCX 读写与转 PDF`rPPTX 适配器开发中"
  $s2.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = "讲稿：重点说明保真度验证方法与真实软件复核。"

  # ---- 第 3 张：表格（ppLayoutBlank = 12）----
  $s3 = $pres.Slides.Add(3, 12)
  $shape = $s3.Shapes.AddTable(3, 3, 60, 80, 500, 200)
  $tbl = $shape.Table
  $tbl.Cell(1,1).Shape.TextFrame.TextRange.Text = "模块"
  $tbl.Cell(1,2).Shape.TextFrame.TextRange.Text = "状态"
  $tbl.Cell(1,3).Shape.TextFrame.TextRange.Text = "测试数"
  $tbl.Cell(2,1).Shape.TextFrame.TextRange.Text = "XLSX"
  $tbl.Cell(2,2).Shape.TextFrame.TextRange.Text = "已完成"
  $tbl.Cell(2,3).Shape.TextFrame.TextRange.Text = "46"
  $tbl.Cell(3,1).Shape.TextFrame.TextRange.Text = "DOCX"
  $tbl.Cell(3,2).Shape.TextFrame.TextRange.Text = "已完成"
  $tbl.Cell(3,3).Shape.TextFrame.TextRange.Text = "73"

  # ---- 第 4 张：会被隐藏（用于验证隐藏幻灯片识别）----
  $s4 = $pres.Slides.Add(4, 2)
  $s4.Shapes.Title.TextFrame.TextRange.Text = "附录：内部草稿"
  $s4.Shapes.Item(2).TextFrame.TextRange.Text = "本页不应出现在正式放映中。"
  $s4.SlideShowTransition.Hidden = -1   # msoTrue

  # 注意：不设置 BuiltInDocumentProperties —— PowerPoint COM 在该属性上
  # 会抛「Object reference not set to an instance of an object」，
  # 而 PowerPoint 自己会填好 Application / Creator，不影响测试。

  # ppSaveAsOpenXMLPresentation = 24
  $pres.SaveAs($resolved, 24)
  $pres.Close()
  $pres = $null
  Write-Output "CREATED: $resolved"
}
catch {
  Write-Output "CREATE_ERROR: $($_.Exception.Message)"
  exit 1
}
finally {
  if ($pres -ne $null) { try { $pres.Close() } catch {} }
  if ($app -ne $null) {
    try { $app.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
