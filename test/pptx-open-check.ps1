# 用真实 Microsoft PowerPoint 打开演示文稿，验证结构与文本。
# 只读打开、不显示窗口，读完立即关闭。
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [int]$Slide = 1,
  [int]$Shape = 1
)

$ErrorActionPreference = 'Stop'
$resolved = (Resolve-Path -LiteralPath $Path).Path
Write-Output "FILE: $resolved"

$app = $null
$pres = $null
try {
  $app = New-Object -ComObject PowerPoint.Application
  # Open(FileName, ReadOnly=-1, Untitled=0, WithWindow=0)
  $pres = $app.Presentations.Open($resolved, -1, 0, 0)

  Write-Output "SLIDES: $($pres.Slides.Count)"
  Write-Output "SLIDE_WIDTH_PT: $([math]::Round($pres.PageSetup.SlideWidth,1))"
  Write-Output "SLIDE_HEIGHT_PT: $([math]::Round($pres.PageSetup.SlideHeight,1))"

  $s = $pres.Slides.Item($Slide)
  Write-Output "SHAPES_ON_SLIDE_${Slide}: $($s.Shapes.Count)"
  # 注意：PowerShell 变量名大小写不敏感，不能用 $shape 承接（会覆盖参数 $Shape）
  $targetShape = $s.Shapes.Item($Shape)
  Write-Output "TEXT_${Slide}_${Shape}: $($targetShape.TextFrame.TextRange.Text)"

  # 主题：Design.Name 是主题名（PowerPoint 确实暴露它），Colors(5) 是 accent1 的 RGB
  $designName = 'n/a'
  try { $designName = $pres.SlideMaster.Design.Name } catch {}
  Write-Output "DESIGN_NAME: $designName"
  $accent1 = 'n/a'
  try { $accent1 = $pres.SlideMaster.Theme.ThemeColorScheme.Colors(5).RGB } catch {}
  Write-Output "ACCENT1_RGB: $accent1"

  # 逐页：标题、形状数、备注页有无、以及目标页每个形状的文本（用于核对新增文本框）
  for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $sl = $pres.Slides.Item($i)
    $t = ''
    try { $t = $sl.Shapes.Title.TextFrame.TextRange.Text } catch {}
    Write-Output "TITLE_${i}: $t"
    Write-Output "SHAPECOUNT_${i}: $($sl.Shapes.Count)"
    $hasNotes = 'no'
    try { if ($sl.HasNotesPage -and $sl.NotesPage.Shapes.Count -gt 0) { $hasNotes = 'yes' } } catch {}
    Write-Output "HASNOTES_${i}: $hasNotes"
    # 版式名：改版式改的是 slide 关系目标，只有 PowerPoint 能告诉我们它实际按哪个版式渲染
    $layoutName = 'n/a'
    try { $layoutName = $sl.CustomLayout.Name } catch {}
    Write-Output "LAYOUT_${i}: $layoutName"
    # 图表：HasChart + ChartType + 系列数与数值（xlColumnClustered=51 / xlLine=4 / xlPie=5 / xlBarClustered=57）
    for ($j = 1; $j -le $sl.Shapes.Count; $j++) {
      $sh = $sl.Shapes.Item($j)
      $hasChart = $false
      try { $hasChart = [bool]$sh.HasChart } catch {}
      if ($hasChart) {
        Write-Output "CHART_${i}_${j}: charttype=$($sh.Chart.ChartType) hastitle=$($sh.Chart.HasTitle)"
        try {
          $sc = $sh.Chart.SeriesCollection()
          Write-Output "CHART_${i}_${j}_SERIES: $($sc.Count)"
          for ($k = 1; $k -le $sc.Count; $k++) {
            $vals = ''
            try { $vals = ($sc.Item($k).Values -join ',') } catch {}
            Write-Output "CHART_${i}_${j}_V${k}: $vals"
          }
        } catch {
          Write-Output "CHART_${i}_${j}_SERIES: not-exposed"
        }
      }
    }
  }

  $pres.Close()
  $pres = $null
  Write-Output 'PPT_OPEN_OK'
}
catch {
  Write-Output "PPT_ERROR: $($_.Exception.Message)"
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
