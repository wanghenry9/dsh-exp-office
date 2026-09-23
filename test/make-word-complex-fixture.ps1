# 用真实 Microsoft Word 创作一份「复杂特性」样本，用于保真度回归。
# 覆盖：多级标题、目录域、批注、修订（未接受）、脚注、书签、超链接、项目符号列表。
#
# 注意：本机 PowerShell 的 COM 绑定对「循环里用变量拼地址」会失败，
# 因此这里统一使用字面地址与 $doc.Range(start, end) 定位。
param([string]$OutPath = "$PSScriptRoot\fixtures\word-complex.docx")

$ErrorActionPreference = 'Stop'
$resolved = [System.IO.Path]::GetFullPath($OutPath)
$dir = Split-Path -Parent $resolved
if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Force }

$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0

  $doc = $word.Documents.Add()

  # ---- 标题与正文（用内置样式，产生大纲级别）----
  $doc.Content.Text = "季度业务报告`r技术方案与实施进展`r本报告由测试脚本生成。`r待确认的风险项包括预算与排期。`r"
  $doc.Paragraphs.Item(1).Range.Style = $doc.Styles.Item(-2)   # wdStyleHeading1
  $doc.Paragraphs.Item(2).Range.Style = $doc.Styles.Item(-3)   # wdStyleHeading2

  # ---- 项目符号列表 ----
  $listPara = $doc.Paragraphs.Item(3).Range
  $listPara.ListFormat.ApplyBulletDefault()

  # ---- 书签 ----
  $bookmarkRange = $doc.Paragraphs.Item(4).Range
  $doc.Bookmarks.Add("RiskSection", $bookmarkRange)

  # ---- 批注 ----
  $commentTarget = $doc.Paragraphs.Item(4).Range
  $doc.Comments.Add($commentTarget, "批注：这里的风险项需要法务复核。")

  # ---- 修订（未接受的插入与删除）----
  $doc.TrackRevisions = $true
  $tail = $doc.Range($doc.Content.End - 2, $doc.Content.End - 1)
  $tail.Text = "本段由修订模式插入。"
  $doc.TrackRevisions = $false

  # ---- 脚注 ----
  $noteRange = $doc.Paragraphs.Item(3).Range
  $doc.Footnotes.Add($noteRange, "", "脚注：数据来源为财务系统。")

  # ---- 目录域 ----
  $tocRange = $doc.Range(0, 0)
  $toc = $doc.TablesOfContents.Add($tocRange, $true, 1, 3)

  # ---- 超链接 ----
  $linkRange = $doc.Paragraphs.Item(4).Range
  $doc.Hyperlinks.Add($linkRange, "https://github.com/deepseek-ai/deepseek-harness", [Type]::Missing, "DeepSeek Harness", "点此访问")

  $doc.SaveAs([ref]$resolved, [ref]16)   # wdFormatDocumentDefault = 16 (.docx)
  $doc.Close($false)
  $doc = $null
  Write-Output "CREATED: $resolved"
}
catch {
  Write-Output "CREATE_ERROR: $($_.Exception.Message)"
  exit 1
}
finally {
  if ($doc -ne $null) { try { $doc.Close($false) } catch {} }
  if ($word -ne $null) {
    try { $word.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
