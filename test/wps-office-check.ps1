# 用 WPS 打开 DOCX / PPTX，做跨软件交叉验证。
# 只读打开、禁用弹窗，读完立即关闭，不保存、不改动源文件。
#
# 用法：
#   pwsh -File test/wps-office-check.ps1 -Path x.docx -Kind writer
#   pwsh -File test/wps-office-check.ps1 -Path x.pptx -Kind presentation
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [ValidateSet('writer', 'presentation')][string]$Kind = 'writer'
)

$ErrorActionPreference = 'Stop'
$resolved = (Resolve-Path -LiteralPath $Path).Path
Write-Output "FILE: $resolved"
Write-Output "KIND: $Kind"

$app = $null
$doc = $null
try {
  if ($Kind -eq 'writer') {
    $app = New-Object -ComObject KWPS.Application
    $app.Visible = $false
    try { $app.DisplayAlerts = 0 } catch {}
    $doc = $app.Documents.Open($resolved, $false, $true)

    Write-Output "PARAGRAPHS: $($doc.Paragraphs.Count)"
    Write-Output "TABLES: $($doc.Tables.Count)"
    Write-Output "COMMENTS: $($doc.Comments.Count)"
    Write-Output "REVISIONS: $($doc.Revisions.Count)"
    Write-Output "BOOKMARKS: $($doc.Bookmarks.Count)"
    Write-Output "FIELDS: $($doc.Fields.Count)"
    Write-Output "INLINESHAPES: $($doc.InlineShapes.Count)"
    Write-Output "P1: $($doc.Paragraphs.Item(1).Range.Text.Trim())"

    $doc.Close($false)
    $doc = $null
  }
  else {
    $app = New-Object -ComObject KWPP.Application
    try { $app.DisplayAlerts = 0 } catch {}
    # Open(FileName, ReadOnly=-1, Untitled=0, WithWindow=0)
    $doc = $app.Presentations.Open($resolved, -1, 0, 0)

    Write-Output "SLIDES: $($doc.Slides.Count)"
    Write-Output "SLIDE_WIDTH_PT: $([math]::Round($doc.PageSetup.SlideWidth, 1))"
    for ($i = 1; $i -le $doc.Slides.Count; $i++) {
      $sl = $doc.Slides.Item($i)
      $t = ''
      try { $t = $sl.Shapes.Title.TextFrame.TextRange.Text } catch {}
      Write-Output "TITLE_${i}: $t"
      # 形状清单：Type 13 = 图片，19 = 表格（与 PowerPoint 的 msoShapeType 一致）
      Write-Output "SHAPES_${i}: $($sl.Shapes.Count)"
      for ($j = 1; $j -le $sl.Shapes.Count; $j++) {
        $sh = $sl.Shapes.Item($j)
        $type = 'n/a'
        try { $type = $sh.Type } catch {}
        $name = 'n/a'
        try { $name = $sh.Name } catch {}
        $w = 0
        $h = 0
        try { $w = [math]::Round($sh.Width, 1); $h = [math]::Round($sh.Height, 1) } catch {}
        Write-Output "SHAPE_${i}_${j}: type=$type name=$name size=${w}x${h}pt"
        if ("$type" -eq '19') {
          $rows = 'n/a'
          $cols = 'n/a'
          try { $rows = $sh.Table.Rows.Count; $cols = $sh.Table.Columns.Count } catch {}
          Write-Output "TABLE_${i}_${j}: rows=$rows cols=$cols"
        }
      }
    }

    $doc.Close()
    $doc = $null
  }
  Write-Output 'WPS_OFFICE_OK'
}
catch {
  Write-Output "WPS_OFFICE_ERROR: $($_.Exception.Message)"
}
finally {
  if ($doc -ne $null) { try { $doc.Close() } catch {} }
  if ($app -ne $null) {
    try { $app.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
