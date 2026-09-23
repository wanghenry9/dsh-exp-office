# 用真实 Microsoft Excel 创作一份「多特性」样本工作簿，用于保真度回归测试。
# 覆盖：条件格式、图表、合并单元格、公式、多工作表、数字格式。
# 生成物为 test/fixtures/excel-rich.xlsx，随后由 fidelity.test.mjs 修改并断言保真度。
#
# 注意：本机 PowerShell 的 COM 绑定对「变量地址 + Cells.Item」组合会抛
# "Unable to cast object of type 'System.Int32' to type 'System.String'"，
# 因此这里统一使用字面地址（Range('B2')）逐格赋值，不构造二维数组、不用 Cells.Item。
param([string]$OutPath = "$PSScriptRoot\fixtures\excel-rich.xlsx")

$ErrorActionPreference = 'Stop'
$resolved = [System.IO.Path]::GetFullPath($OutPath)
$dir = Split-Path -Parent $resolved
if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Force }

$excel = $null
$wb = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  $wb = $excel.Workbooks.Add()
  while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = '销售数据'

  $ws.Range('A1').Value2 = '月份'
  $ws.Range('A2').Value2 = '1月'; $ws.Range('A3').Value2 = '2月'; $ws.Range('A4').Value2 = '3月'
  $ws.Range('A5').Value2 = '4月'; $ws.Range('A6').Value2 = '5月'; $ws.Range('A7').Value2 = '6月'
  $ws.Range('B1').Value2 = '销售额'
  $ws.Range('B2').Value2 = 12000; $ws.Range('B3').Value2 = 15500; $ws.Range('B4').Value2 = 9800
  $ws.Range('B5').Value2 = 21000; $ws.Range('B6').Value2 = 17500; $ws.Range('B7').Value2 = 24000
  $ws.Range('C1').Value2 = '成本'
  $ws.Range('C2').Value2 = 8000; $ws.Range('C3').Value2 = 9200; $ws.Range('C4').Value2 = 7100
  $ws.Range('C5').Value2 = 13000; $ws.Range('C6').Value2 = 11000; $ws.Range('C7').Value2 = 14200

  # 公式列 + 汇总公式 + 数字格式
  $ws.Range('D1').Value2 = '利润'
  $ws.Range('D2:D7').Formula = '=B2-C2'
  $ws.Range('B8').Formula = '=SUM(B2:B7)'
  $ws.Range('B8').NumberFormat = '#,##0.00'

  # 合并单元格
  $ws.Range('A10:D10').Merge()
  $ws.Range('A10').Value2 = '合并标题行'

  # 条件格式：销售额大于 20000 标红加粗
  $cf = $ws.Range('B2:B7').FormatConditions.Add(1, 3, '20000')
  $cf.Interior.Color = 255
  $cf.Font.Bold = $true

  # 图表
  $chart = $ws.Shapes.AddChart2(251, 51, 260, 10, 400, 260)
  $chart.Chart.SetSourceData($ws.Range('A1:B7'))
  $chart.Chart.HasTitle = $true
  $chart.Chart.ChartTitle.Text = '月度销售额'

  # 第二张工作表（跨表公式）
  $ws2 = $wb.Worksheets.Add()
  $ws2.Name = '汇总'
  $ws2.Range('A1').Value2 = '指标'; $ws2.Range('B1').Value2 = '数值'
  $ws2.Range('A2').Value2 = '总销售额'; $ws2.Range('B2').Formula = "='销售数据'!B8"
  $ws2.Range('A3').Value2 = '备注'; $ws2.Range('B3').Value2 = '由真实 Excel 创作的多特性样本'

  $ws.Activate()
  $wb.SaveAs($resolved, 51)  # xlOpenXMLWorkbook
  $wb.Close($false)
  $wb = $null
  Write-Output "CREATED: $resolved"
}
catch {
  Write-Output "CREATE_ERROR: $($_.Exception.Message)"
  exit 1
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
