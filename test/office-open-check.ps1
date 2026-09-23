# 用真实 Microsoft Excel 打开插件输出的工作簿，验证兼容性。
# 只读打开、禁用弹窗与外部链接更新，读完立即关闭，不保存、不改动源文件。
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [string]$ExpectA2 = '',
  [string]$ExpectB2 = ''
)

$ErrorActionPreference = 'Stop'
$resolved = (Resolve-Path -LiteralPath $Path).Path
Write-Output "FILE: $resolved"

$excel = $null
$wb = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.EnableEvents = $false
  $excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable，禁止宏

  # UpdateLinks=0 不更新外部链接；ReadOnly=$true 只读；Notify=$false 不弹窗
  $wb = $excel.Workbooks.Open($resolved, 0, $true, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true)

  $sheets = @()
  foreach ($ws in $wb.Worksheets) { $sheets += $ws.Name }
  Write-Output "SHEETS: $($sheets -join ' | ')"

  $ws1 = $wb.Worksheets.Item(1)
  Write-Output "A1: $($ws1.Range('A1').Text)"
  Write-Output "A2: $($ws1.Range('A2').Text)"
  Write-Output "B2: $($ws1.Range('B2').Text)"
  Write-Output "D2_FORMULA: $($ws1.Range('D2').Formula)"
  Write-Output "USEDRANGE: $($ws1.UsedRange.Address($false, $false))"

  # 表格对象（ListObject）：插件建的表格必须被 Excel 认出来，否则文件会被要求修复
  foreach ($sheet in $wb.Worksheets) {
    $tables = $sheet.ListObjects
    Write-Output "LISTOBJECTS_$($sheet.Name): $($tables.Count)"
    for ($i = 1; $i -le $tables.Count; $i++) {
      $t = $tables.Item($i)
      Write-Output "  TABLE_$($sheet.Name)_${i}: name=$($t.Name) range=$($t.Range.Address($false, $false)) header=$($t.HeaderRowRange.Address($false, $false)) totals=$($t.ShowTotals) style=$($t.TableStyle.Name)"
    }
  }
  Write-Output "CHARTCOUNT: $($ws1.ChartObjects().Count)"
  # 图表细节：类型（51=柱状簇、57=条形簇、4=折线、5=饼图）、标题、系列数、尺寸
  for ($i = 1; $i -le $ws1.ChartObjects().Count; $i++) {
    $co = $ws1.ChartObjects().Item($i)
    $chart = $co.Chart
    $title = ''
    try { if ($chart.HasTitle) { $title = $chart.ChartTitle.Text } } catch {}
    Write-Output "CHART_${i}: type=$($chart.ChartType) title='$title' series=$($chart.SeriesCollection().Count) size=$([math]::Round($co.Width,1))x$([math]::Round($co.Height,1))px"
  }

  if ($ExpectA2 -ne '' -and $ws1.Range('A2').Text -ne $ExpectA2) {
    Write-Output "MISMATCH_A2: expected '$ExpectA2' got '$($ws1.Range('A2').Text)'"
  }
  if ($ExpectB2 -ne '' -and $ws1.Range('B2').Text -ne $ExpectB2) {
    Write-Output "MISMATCH_B2: expected '$ExpectB2' got '$($ws1.Range('B2').Text)'"
  }

  $wb.Close($false)
  $wb = $null
  Write-Output "OPEN_OK"
}
catch {
  Write-Output "EXCEL_ERROR: $($_.Exception.Message)"
}
finally {
  if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
  if ($excel -ne $null) {
    try { $excel.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
