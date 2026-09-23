# dsh-exp-office: local Office/WPS bridge (fixed template; ASCII-only by design)
#
# This is the ONLY place in the plugin that talks to local Office/WPS.
# NOTE: this file is deliberately ASCII-only. Windows PowerShell 5.1 reads a BOM-less
# UTF-8 script as ANSI, so non-ASCII text here would be mangled and would break parsing.
# All user-facing Chinese text is produced on the Node side.
#
# Safety contract (spec section 12.4):
#   * user data only arrives through parameters; nothing is interpolated into code;
#   * macros never run: AutomationSecurity = 3 (msoAutomationSecurityForceDisable);
#   * external links are not updated: AskToUpdateLinks = false + Open(UpdateLinks:=0);
#   * no dialogs: DisplayAlerts = false, engine window hidden when supported;
#   * the result is written to a JSON file (no stdio pipes: avoids encoding/pipe issues);
#   * only processes started by THIS run are cleaned up (PID set difference);
#   * a state file lets the Node side run one extra cleanup after a hard timeout.
param(
  [Parameter(Mandatory = $true)][ValidateSet('detect', 'recalc', 'rerender', 'cleanup')][string]$Action,
  [string]$InputPath = '',
  [string]$OutputPath = '',
  [ValidateSet('xlsx', 'xlsm', 'docx', 'docm', 'pptx', 'pptm')][string]$Kind = 'xlsx',
  [ValidateSet('auto', 'excel', 'word', 'powerpoint', 'wps')][string]$Engine = 'auto',
  [switch]$ProbeCom,
  [Parameter(Mandatory = $true)][string]$ResultPath,
  [string]$StatePath = ''
)

$ErrorActionPreference = 'Stop'

# Engine table: COM ProgID, executable name, ASCII label, supported kinds.
$ENGINES = @(
  @{ id = 'excel';      prog_id = 'Excel.Application';      exe = 'EXCEL.EXE';    label = 'Microsoft Excel';      kinds = @('xlsx', 'xlsm') },
  @{ id = 'word';       prog_id = 'Word.Application';       exe = 'WINWORD.EXE';  label = 'Microsoft Word';       kinds = @('docx', 'docm') },
  @{ id = 'powerpoint'; prog_id = 'PowerPoint.Application'; exe = 'POWERPNT.EXE'; label = 'Microsoft PowerPoint'; kinds = @('pptx', 'pptm') },
  @{ id = 'wps';        prog_id = 'KET.Application';        exe = 'et.exe';       label = 'WPS Spreadsheets';     kinds = @('xlsx', 'xlsm') },
  @{ id = 'wps';        prog_id = 'KWPS.Application';       exe = 'wps.exe';      label = 'WPS Writer';           kinds = @('docx', 'docm') },
  @{ id = 'wps';        prog_id = 'KWPP.Application';       exe = 'wpp.exe';      label = 'WPS Presentation';     kinds = @('pptx', 'pptm') }
)

$ENGINE_EXE_NAMES = @('EXCEL', 'WINWORD', 'POWERPNT', 'et', 'wps', 'wpp', 'wpscloudsvr', 'wpscenter')

function Get-EnginePids {
  $found = @()
  foreach ($name in $ENGINE_EXE_NAMES) {
    foreach ($p in @(Get-Process -Name $name -ErrorAction SilentlyContinue)) { $found += [int]$p.Id }
  }
  return $found
}

function Write-Result($obj) {
  $json = $obj | ConvertTo-Json -Depth 8 -Compress
  $dir = Split-Path -Parent $ResultPath
  if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  [System.IO.File]::WriteAllText($ResultPath, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Get-FileVersion($path) {
  try { return (Get-Item -LiteralPath $path).VersionInfo.FileVersion } catch { return $null }
}

function Get-RegistryValue($path, $name) {
  try { return (Get-ItemProperty -LiteralPath $path -Name $name -ErrorAction Stop).$name } catch { return $null }
}

# LocalServer32 may carry arguments ("...EXCEL.EXE /automation") and quotes: keep the exe only.
function ConvertTo-ExePath($raw) {
  if (-not $raw) { return $null }
  $text = $raw.Trim()
  if ($text.StartsWith('"')) {
    $end = $text.IndexOf('"', 1)
    if ($end -gt 0) { return $text.Substring(1, $end - 1) }
  }
  $cut = $text.IndexOf(' /')
  if ($cut -gt 0) { return $text.Substring(0, $cut).Trim() }
  return $text
}

# Resolve the COM ProgID to its LocalServer32 executable, then fall back to the
# Kingsoft install root (WPS registers its servers under WOW6432Node, and some
# builds expose no LocalServer32 at all). Registry only: starts nothing.
function Get-ProgIdExe($progId, $exeName) {
  $cur = Get-RegistryValue "Registry::HKEY_CLASSES_ROOT\$progId\CurVer" '(default)'
  $key = if ($cur) { "Registry::HKEY_CLASSES_ROOT\$cur\CLSID" } else { "Registry::HKEY_CLASSES_ROOT\$progId\CLSID" }
  $clsid = Get-RegistryValue $key '(default)'
  if ($clsid) {
    $raw = Get-RegistryValue "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\LocalServer32" '(default)'
    if (-not $raw) { $raw = Get-RegistryValue "Registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID\$clsid\LocalServer32" '(default)' }
    $exe = ConvertTo-ExePath $raw
    if ($exe -and (Test-Path -LiteralPath $exe)) { return $exe }
  }
  foreach ($root in @('HKLM:\SOFTWARE\WOW6432Node\Kingsoft\Office\6.0\common', 'HKLM:\SOFTWARE\Kingsoft\Office\6.0\common', 'HKCU:\SOFTWARE\Kingsoft\Office\6.0\common')) {
    $install = Get-RegistryValue $root 'InstallRoot'
    if ($install) {
      $candidate = Join-Path (Join-Path $install 'office6') $exeName
      if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
  }
  $exe = ConvertTo-ExePath $raw
  return $exe
}

$startedPids = @(Get-EnginePids)

if ($StatePath -ne '') {
  $state = @{ action = $Action; started_pids = $startedPids; result_path = $ResultPath; time = (Get-Date).ToString('o') }
  [System.IO.File]::WriteAllText($StatePath, ($state | ConvertTo-Json -Compress), (New-Object System.Text.UTF8Encoding($false)))
}

$app = $null
$doc = $null
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$result = @{ action = $Action; ok = $false; engine = $null; version = $null; error = $null; warnings = @() }

try {
  if ($Action -eq 'detect') {
    $list = @()
    foreach ($e in $ENGINES) {
      $exe = Get-ProgIdExe $e.prog_id $e.exe
      $exists = $false
      $version = $null
      if ($exe -and (Test-Path -LiteralPath $exe)) {
        $exists = $true
        $version = Get-FileVersion $exe
      }
      $entry = @{
        id           = $e.id
        label        = $e.label
        prog_id      = $e.prog_id
        exe_path     = $exe
        installed    = [bool]$exists
        file_version = $version
        com_version  = $null
        com_ok       = $false
        com_error    = $null
        kinds        = $e.kinds
      }
      if ($ProbeCom -and $exists) {
        try {
          $probe = New-Object -ComObject $e.prog_id
          try { $entry.com_version = [string]$probe.Version } catch { $entry.com_version = $null }
          $entry.com_ok = $true
          try { $probe.Quit() } catch { }
          try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($probe) | Out-Null } catch { }
        } catch {
          $entry.com_error = $_.Exception.Message
        }
      }
      $list += $entry
    }
    $result.engines = $list
    $result.ok = $true
  }
  elseif ($Action -eq 'cleanup') {
    $result.ok = $true
    $result.note = 'cleanup only reclaims the PIDs recorded in the state file'
  }
  elseif ($Action -eq 'recalc' -or $Action -eq 'rerender') {
    if ($InputPath -eq '' -or $OutputPath -eq '') { throw 'recalc/rerender need -InputPath and -OutputPath' }
    if (-not (Test-Path -LiteralPath $InputPath)) { throw "input file not found: $InputPath" }
    # Office automation resolves relative paths against ITS OWN working directory,
    # so both paths must be absolute before they reach the engine.
    $InputPath = (Resolve-Path -LiteralPath $InputPath).Path
    $outDir = Split-Path -Parent $OutputPath
    if ($outDir -eq '') { $outDir = (Get-Location).Path }
    if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    $outDir = (Resolve-Path -LiteralPath $outDir).Path
    $OutputPath = Join-Path $outDir (Split-Path -Leaf $OutputPath)

    $candidates = @($ENGINES | Where-Object { $_.kinds -contains $Kind })
    if ($Engine -ne 'auto') { $candidates = @($candidates | Where-Object { $_.id -eq $Engine }) }
    if ($candidates.Count -eq 0) { throw "no engine matches kind=$Kind engine=$Engine" }
    # Prefer Microsoft Office over WPS unless the caller asked for WPS explicitly.
    if ($Engine -eq 'auto') {
      $candidates = @($candidates | Sort-Object @{ Expression = { if ($_.id -eq 'wps') { 1 } else { 0 } } })
    }

    $used = $null
    foreach ($c in $candidates) {
      try {
        $app = New-Object -ComObject $c.prog_id
        $used = $c
        break
      } catch {
        $result.warnings += "engine $($c.prog_id) unavailable: $($_.Exception.Message)"
        $app = $null
      }
    }
    if ($app -eq $null) { throw "no candidate engine could be created (kind=$Kind)" }

    $result.engine = $used.id
    $result.prog_id = $used.prog_id
    try { $result.version = [string]$app.Version } catch { $result.version = $null }

    try { $app.DisplayAlerts = $false } catch { }
    try { $app.AskToUpdateLinks = $false } catch { }
    try { $app.AutomationSecurity = 3 } catch { }
    try { $app.EnableEvents = $false } catch { }
    try {
      $app.Visible = $false
    } catch {
      $result.warnings += 'engine does not support a hidden window (PowerPoint); dialogs stay disabled anyway'
    }

    if ($Kind -eq 'xlsx' -or $Kind -eq 'xlsm') {
      $doc = $app.Workbooks.Open($InputPath, 0, $false)
      try { $app.CalculateFullRebuild() } catch { $result.warnings += "CalculateFullRebuild failed: $($_.Exception.Message)" }
      $fmt = if ($Kind -eq 'xlsm') { 52 } else { 51 }
      $doc.SaveAs($OutputPath, $fmt)
    }
    elseif ($Kind -eq 'docx' -or $Kind -eq 'docm') {
      $doc = $app.Documents.Open($InputPath, $false, $false, $false)
      try { $doc.Fields.Update() | Out-Null } catch { $result.warnings += "Fields.Update failed: $($_.Exception.Message)" }
      try { $doc.Repaginate() } catch { }
      $fmt = if ($Kind -eq 'docm') { 13 } else { 16 }
      $doc.SaveAs2($OutputPath, $fmt)
    }
    else {
      $doc = $app.Presentations.Open($InputPath, $true, $false, $false)
      $fmt = if ($Kind -eq 'pptm') { 25 } else { 24 }
      try { $doc.SaveAs($OutputPath, $fmt) } catch { $doc.SaveAs($OutputPath) }
    }
    $result.ok = $true
  }
  else {
    throw "unknown action: $Action"
  }
} catch {
  $result.ok = $false
  $result.error = $_.Exception.Message
} finally {
  if ($doc -ne $null) { try { $doc.Close($false) } catch { } ; $doc = $null }
  if ($app -ne $null) {
    try { $app.Quit() } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
    $app = $null
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()

  # Kill only engine processes started by THIS run (PID set difference).
  # Poll a few rounds: WPS spawns its worker lazily and can reappear right after Quit().
  $killed = @()
  for ($round = 0; $round -lt 5; $round++) {
    Start-Sleep -Milliseconds 400
    $now = @(Get-EnginePids)
    $mine = @($now | Where-Object { $startedPids -notcontains $_ })
    if ($mine.Count -eq 0 -and $round -ge 1) { break }
    foreach ($procId in $mine) {
      try {
        Stop-Process -Id $procId -Force -ErrorAction Stop
        if ($killed -notcontains $procId) { $killed += $procId }
      } catch { }
    }
  }
  Start-Sleep -Milliseconds 300
  $result.started_pids = $startedPids
  $result.killed_pids = $killed
  $result.leftover_pids = @(Get-EnginePids | Where-Object { $startedPids -notcontains $_ })
  $sw.Stop()
  $result.elapsed_ms = [int]$sw.ElapsedMilliseconds
  $result.output_size = if ($OutputPath -ne '' -and (Test-Path -LiteralPath $OutputPath)) { (Get-Item -LiteralPath $OutputPath).Length } else { $null }
  Write-Result $result
  if ($StatePath -ne '' -and (Test-Path -LiteralPath $StatePath)) { Remove-Item -LiteralPath $StatePath -Force -ErrorAction SilentlyContinue }
}
