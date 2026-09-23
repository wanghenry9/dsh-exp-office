# 用 Microsoft Word 的 PDF 重排（PDF Reflow）独立打开 PDF，验证文件仍可被第三方解析器读取。
#
# 为什么用 Word：本机没有 pdftotext / pdfinfo，而 Word 2013+ 自带独立的 PDF 解析器，
# 它可以作为「不是我们自己写的解析器」的第三方裁判 —— 用来确认插件写出的 PDF
# 没有被写坏（页序、页数、文本内容都能独立读出）。
#
# 注意：本脚本刻意只用 ASCII 标签。Windows PowerShell 5.1 按 ANSI 读取无 BOM 的 .ps1，
# 中文字符串字面量会变成乱码（注释乱码无害，但字符串里的 $() 会被吃掉），
# 所以标签一律用英文，中文只出现在注释与 Word 读回的数据里。
#
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File test\pdf-open-check.ps1 -Path <pdf> [-Raw]
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [switch]$Raw,
  # Optional ASCII keyword: report how many times Word's *independent* PDF parser sees it.
  # Used to cross-check our own full-text search hit counts (ASCII terms only).
  [string]$Match = '',
  # Optional path: write Word's parsed text to this file as UTF-8 (encoding-safe channel
  # for Chinese content; stdout would be mangled by the console code page).
  [string]$TextOut = ''
)

$ErrorActionPreference = 'Stop'
$target = (Resolve-Path -LiteralPath $Path).Path
Write-Output "FILE: $target"

# 只清理本次探测启动的 Word，绝不动用户自己开的窗口。
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

$word = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($target, $false, $true)
  $text = $doc.Content.Text
  $pages = $doc.ComputeStatistics(2)
  $words = $doc.ComputeStatistics(0)
  # 注意：这是 **Word 重排后的页数**，不是 PDF 的页数 —— PDF 重排会按 Word 的排版重新分页
  # （实测：1 页 PDF 排出 2 页，3 页 PDF 排出 4 页）。页数请以插件读取 + xref 结构契约为准，
  # 这里的文本与词数才是可靠的独立信号。
  Write-Output "REFLOW_PAGES: $pages"
  Write-Output "WORDS: $words"
  if ($Match -ne '') {
    $hits = ([regex]::Matches($text, [regex]::Escape($Match))).Count
    Write-Output "MATCH_COUNT: $hits"
    Write-Output "MATCH_TERM: $Match"
  }
  # PDF 重排会把 PDF 元数据映射到 Word 的内置文档属性上，可用来独立验证元数据写入。
  try {
    foreach ($prop in @('Title', 'Author', 'Subject', 'Keywords', 'Comments')) {
      $value = $doc.BuiltInDocumentProperties($prop).Value
      if ($value) { Write-Output "PROP_${prop}: $value" }
    }
  } catch {
    Write-Output "PROP-ERROR: $($_.Exception.Message)"
  }
  # PDF reflow moves text inside the page margins into the header/footer story,
  # so page-number overlays only show up there (Content.Text cannot see them).
  $stories = @()
  try {
    foreach ($section in $doc.Sections) {
      $footer = $section.Footers.Item(1).Range.Text
      $header = $section.Headers.Item(1).Range.Text
      # 先收集：文档 Close 之后就取不到这些 story 了
      $stories += "HEADER: $header"
      $stories += "FOOTER: $footer"
      if ($footer -and $footer.Trim() -ne '') { Write-Output "FOOTER: $($footer.Trim())" }
      if ($header -and $header.Trim() -ne '') { Write-Output "HEADER: $($header.Trim())" }
    }
  } catch {
    Write-Output "HEADERFOOTER-ERROR: $($_.Exception.Message)"
  }
  $doc.Close($false)
  # 把 Word 解析出的文本写成 UTF-8 文件：中文经 stdout 回传会被控制台编码弄坏，
  # 写文件再让调用方按字节比对才可靠（从零生成的中文 PDF 就靠这条做独立验证）。
  if ($TextOut -ne '') {
    $full = [System.IO.Path]::GetFullPath($TextOut)
    [System.IO.File]::WriteAllText($full, $text, (New-Object System.Text.UTF8Encoding($false)))
    Write-Output "TEXT_OUT: $full"
    # 页眉页脚 story 另写一个伴随文件：贴边水印/页码会被重排放进 footer story，
    # Content.Text 看不到它们 —— 校验中文水印必须看这里。
    [System.IO.File]::WriteAllText("$full.stories", ($stories -join "`n"), (New-Object System.Text.UTF8Encoding($false)))
  }
  if ($Raw) {
    Write-Output 'TEXT-RAW:'
    Write-Output $text
  } else {
    Write-Output 'TEXT:'
    foreach ($line in ($text -split "`r|`n|`a" | Where-Object { $_.Trim() -ne '' })) {
      Write-Output $line.Trim()
    }
  }
} finally {
  if ($word) {
    try { $word.Quit() } catch { Write-Output "QUIT-ERROR: $($_.Exception.Message)" }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
  }
  Start-Sleep -Milliseconds 500
  foreach ($proc in @(Get-Process WINWORD -ErrorAction SilentlyContinue)) {
    if ($before -notcontains $proc.Id) {
      $stale = $proc.Id
      Write-Output "CLEANUP: stopped Word process $stale"
      Stop-Process -Id $stale -Force -ErrorAction SilentlyContinue
    }
  }
}
