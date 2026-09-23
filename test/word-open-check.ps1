# 用真实 Microsoft Word 打开 DOCX，验证结构与内容。
# 只读打开、禁用弹窗与宏，读完立即关闭，不保存、不改动源文件。
param(
  [Parameter(Mandatory = $true)][string]$Path
)

$ErrorActionPreference = 'Stop'
$resolved = (Resolve-Path -LiteralPath $Path).Path
Write-Output "FILE: $resolved"

$app = $null
$doc = $null
try {
  $app = New-Object -ComObject Word.Application
  $app.Visible = $false
  $app.DisplayAlerts = 0
  try { $app.AutomationSecurity = 3 } catch {}   # 禁止宏

  $doc = $app.Documents.Open($resolved, $false, $true)
  Write-Output "PARAGRAPHS: $($doc.Paragraphs.Count)"
  Write-Output "TABLES: $($doc.Tables.Count)"
  Write-Output "WORDS: $($doc.Words.Count)"

  $limit = [Math]::Min(10, $doc.Paragraphs.Count)
  foreach ($i in 1..$limit) {
    $text = $doc.Paragraphs.Item($i).Range.Text.Trim()
    if ($text -ne '') {
      Write-Output "P${i}: [$($doc.Paragraphs.Item($i).Style.NameLocal)] $text"
    }
  }

  if ($doc.Tables.Count -gt 0) {
    Write-Output "T1R1C1: $($doc.Tables.Item(1).Cell(1,1).Range.Text.Trim())"
    # Table structure: rows / columns / each cell text. Merged cells make Cell(r,c) throw,
    # so every access is guarded and reported as "merged" instead of crashing the probe.
    $table = $doc.Tables.Item(1)
    Write-Output "T1_ROWS: $($table.Rows.Count)"
    Write-Output "T1_COLUMNS: $($table.Columns.Count)"
    # Style + conditional-formatting switches: proves w:tblStyle / w:tblLook were written correctly
    # (an unknown style id silently falls back to "Normal Table").
    $styleName = 'n/a'
    try { $styleName = $table.Style.NameLocal } catch {}
    Write-Output "T1_STYLE: $styleName"
    try {
      Write-Output "T1_LOOK: headingRows=$($table.ApplyStyleHeadingRows) lastRow=$($table.ApplyStyleLastRow) firstCol=$($table.ApplyStyleFirstColumn) rowBands=$($table.ApplyStyleRowBands)"
    } catch {
      Write-Output "T1_LOOK: not-exposed"
    }
    for ($r = 1; $r -le $table.Rows.Count; $r++) {
      for ($c = 1; $c -le $table.Columns.Count; $c++) {
        $text = ''
        try { $text = $table.Cell($r, $c).Range.Text.Trim() } catch { $text = '<merged>' }
        Write-Output "T1_${r}_${c}: $text"
      }
    }
  }
  Write-Output "INLINESHAPES: $($doc.InlineShapes.Count)"
  # 浮动形状（设置环绕方式后，图片会从 InlineShapes 移到 Shapes）。
  # WrapFormat.Type: 0=square 1=tight 2=through 3=none 4=topBottom 5=behind 6=front
  Write-Output "FLOATSHAPES: $($doc.Shapes.Count)"
  if ($doc.Shapes.Count -gt 0) {
    $sh = $doc.Shapes.Item(1)
    $wrapType = 'n/a'
    try { $wrapType = $sh.WrapFormat.Type } catch {}
    Write-Output "SHAPE1_WRAP: $wrapType"
    Write-Output "SHAPE1_POS: $([math]::Round($sh.Left,1)),$([math]::Round($sh.Top,1))"
    Write-Output "SHAPE1_SIZE: $([math]::Round($sh.Width,1))x$([math]::Round($sh.Height,1))"
  }
  # Paragraph styles of the first three paragraphs + page count:
  # proves w:pStyle references resolved (an unknown style id silently falls back to Normal).
  for ($i = 1; $i -le [Math]::Min(3, $doc.Paragraphs.Count); $i++) {
    $style = 'n/a'
    try { $style = $doc.Paragraphs.Item($i).Style.NameLocal } catch {}
    $text = ''
    try { $text = $doc.Paragraphs.Item($i).Range.Text.Trim() } catch {}
    Write-Output "PARA_${i}: style='$style' text='$($text.Substring(0, [Math]::Min(18, $text.Length)))'"
  }
  Write-Output "PAGES: $($doc.ComputeStatistics(2))"
  if ($doc.InlineShapes.Count -gt 0) {
    $shape = $doc.InlineShapes.Item(1)
    Write-Output "SHAPE1: type=$($shape.Type) w=$([math]::Round($shape.Width,1))pt h=$([math]::Round($shape.Height,1))pt"
  }
  Write-Output "COMMENTS: $($doc.Comments.Count)"
  if ($doc.Comments.Count -gt 0) {
    $c = $doc.Comments.Item(1)
    Write-Output "COMMENT_1: author=$($c.Author) text=$($c.Range.Text.Trim())"
  }
  Write-Output "REVISIONS: $($doc.Revisions.Count)"
  if ($doc.Revisions.Count -gt 0) {
    # Type: 1=insert 2=delete 3=format ...; text proves we read the same revisions Word sees
    for ($i = 1; $i -le [Math]::Min(4, $doc.Revisions.Count); $i++) {
      $rev = $doc.Revisions.Item($i)
      $revText = ''
      try { $revText = $rev.Range.Text.Trim() } catch {}
      Write-Output "REVISION_${i}: type=$($rev.Type) text=$revText"
    }
  }
  Write-Output "BOOKMARKS: $($doc.Bookmarks.Count)"
  if ($doc.Bookmarks.Count -gt 0) {
    for ($i = 1; $i -le [Math]::Min(4, $doc.Bookmarks.Count); $i++) {
      $bm = $doc.Bookmarks.Item($i)
      $bmText = ''
      try { $bmText = $bm.Range.Text.Trim() } catch {}
      Write-Output "BOOKMARK_${i}: name=$($bm.Name) text=$bmText"
    }
  }
  Write-Output "HYPERLINKS: $($doc.Hyperlinks.Count)"
  Write-Output "TOCS: $($doc.TablesOfContents.Count)"
  Write-Output "FOOTNOTES: $($doc.Footnotes.Count)"
  Write-Output "HEADER: $($doc.Sections.Item(1).Headers.Item(1).Range.Text.Trim())"
  Write-Output "FOOTER: $($doc.Sections.Item(1).Footers.Item(1).Range.Text.Trim())"
  Write-Output "PAGE_WIDTH_PT: $($doc.PageSetup.PageWidth)"
  # 页面设置：长宽、方向（0=纵向 1=横向）、页边距（磅）。用 [math]::Round 避免 COM 返回浮点噪声。
  Write-Output "PAGE_HEIGHT_PT: $($doc.PageSetup.PageHeight)"
  Write-Output "ORIENTATION: $($doc.PageSetup.Orientation)"
  Write-Output "MARGIN_TOP_PT: $([math]::Round($doc.PageSetup.TopMargin,1))"
  Write-Output "MARGIN_BOTTOM_PT: $([math]::Round($doc.PageSetup.BottomMargin,1))"
  Write-Output "MARGIN_LEFT_PT: $([math]::Round($doc.PageSetup.LeftMargin,1))"
  Write-Output "MARGIN_RIGHT_PT: $([math]::Round($doc.PageSetup.RightMargin,1))"
  Write-Output "SECTIONS: $($doc.Sections.Count)"
  # 页码域是否真的在页脚里（33 = wdFieldPage）；域计数能区分「域」与「普通文本 1」。
  $footerFields = 0
  $footerFieldTypes = @()
  try {
    $footerRange = $doc.Sections.Item(1).Footers.Item(1).Range
    $footerFields = $footerRange.Fields.Count
    for ($i = 1; $i -le $footerFields; $i++) { $footerFieldTypes += $footerRange.Fields.Item($i).Type }
  } catch {}
  Write-Output "FOOTER_FIELDS: $footerFields"
  Write-Output "FOOTER_FIELD_TYPES: $($footerFieldTypes -join ',')"
  $headerFields = 0
  try { $headerFields = $doc.Sections.Item(1).Headers.Item(1).Range.Fields.Count } catch {}
  Write-Output "HEADER_FIELDS: $headerFields"

  $doc.Close($false)
  $doc = $null
  Write-Output 'WORD_OPEN_OK'
}
catch {
  Write-Output "WORD_ERROR: $($_.Exception.Message)"
}
finally {
  if ($doc -ne $null) { try { $doc.Close($false) } catch {} }
  if ($app -ne $null) {
    try { $app.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
