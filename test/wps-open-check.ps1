# 用真实 WPS 表格（KET.Application）打开插件输出的工作簿，验证兼容性。
# 只读打开、禁用弹窗，读完立即关闭，不保存、不改动源文件。
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [string]$ExpectA2 = '',
  [string]$ExpectB2 = ''
)

$ErrorActionPreference = 'Stop'
$resolved = (Resolve-Path -LiteralPath $Path).Path
Write-Output "FILE: $resolved"

$app = $null
$wb = $null
try {
  $app = New-Object -ComObject KET.Application
  $app.Visible = $false
  $app.DisplayAlerts = $false
  try { $app.AskToUpdateLinks = $false } catch {}
  try { $app.EnableEvents = $false } catch {}

  $wb = $app.Workbooks.Open($resolved)
  Start-Sleep -Milliseconds 500

  $sheets = @()
  foreach ($ws in $wb.Worksheets) { $sheets += $ws.Name }
  Write-Output "SHEETS: $($sheets -join ' | ')"

  $ws1 = $wb.Worksheets.Item(1)
  Write-Output "A1: $($ws1.Range('A1').Text)"
  Write-Output "A2: $($ws1.Range('A2').Text)"
  Write-Output "B2: $($ws1.Range('B2').Text)"
  Write-Output "D2_FORMULA: $($ws1.Range('D2').Formula)"

  # ListObject (table objects). WPS may not expose this collection; report honestly instead of crashing.
  try {
    $tables = $ws1.ListObjects
    Write-Output "LISTOBJECTS: $($tables.Count)"
    for ($i = 1; $i -le $tables.Count; $i++) {
      $t = $tables.Item($i)
      $name = 'n/a'
      try { $name = $t.Name } catch {}
      $range = 'n/a'
      try { $range = $t.Range.Address($false, $false) } catch {}
      Write-Output "  TABLE_${i}: name=$name range=$range"
    }
  } catch {
    Write-Output "LISTOBJECTS: not-exposed-by-wps"
  }

  if ($ExpectA2 -ne '' -and $ws1.Range('A2').Text -ne $ExpectA2) {
    Write-Output "MISMATCH_A2: expected '$ExpectA2' got '$($ws1.Range('A2').Text)'"
  }
  if ($ExpectB2 -ne '' -and $ws1.Range('B2').Text -ne $ExpectB2) {
    Write-Output "MISMATCH_B2: expected '$ExpectB2' got '$($ws1.Range('B2').Text)'"
  }

  $wb.Close($false)
  $wb = $null
  Write-Output "WPS_OPEN_OK"
}
catch {
  Write-Output "WPS_ERROR: $($_.Exception.Message)"
}
finally {
  if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
  if ($app -ne $null) {
    try { $app.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
