#Requires -Version 5.1
param(
  [Parameter(Mandatory=$true)][string]$Folder
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Parse-AmountFromText {
  param([Parameter(Mandatory)][string]$Text)
  $patterns = @(
    '(?im)\b(?:grand\s+total|total\s+amount|amount\s+paid|total\s+paid|net\s+total|invoice\s+total|remittance\s+total)\D*(\$?\s?-?\d{1,3}(?:,\d{3})*(?:\.\d{2})?)',
    '(?im)\b(?:AUD|NZD)\s*(\$?\s?-?\d{1,3}(?:,\d{3})*(?:\.\d{2})?)',
    '(?im)\btotal(?:\s+amount)?\b\D*(\$?\s?-?\d{1,3}(?:,\d{3})*(?:\.\d{2})?)'
  )
  foreach($p in $patterns){ $m=[regex]::Match($Text,$p); if($m.Success){ $raw=$m.Groups[1].Value.Trim(); return ($raw -replace '\s','') } }
  return ''
}

function Try-RenameWithAmount {
  param([Parameter(Mandatory)][string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { return $Path }
  $ext = [IO.Path]::GetExtension($Path)
  if ($ext -notmatch '^\.pdf$' -and $ext -notmatch '^\.PDF$') { return $Path }
  # Try pdftotext if present
  $text = ''
  $cmd = Get-Command pdftotext -ErrorAction SilentlyContinue
  if (-not $cmd) {
    $here = (Get-Location).Path
    $cands = @(
      (Join-Path $here 'tools\poppler\Library\bin\pdftotext.exe'),
      (Join-Path $here 'tools\poppler\bin\pdftotext.exe')
    )
    foreach($p in $cands){ if(Test-Path -LiteralPath $p){ $cmd = @{ Source = $p }; break } }
    if (-not $cmd) {
      try {
        $pop = Join-Path $here 'tools\poppler'
        if (Test-Path -LiteralPath $pop) {
          $hit = Get-ChildItem -LiteralPath $pop -Recurse -Filter 'pdftotext.exe' -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
          if ($hit) { $cmd = @{ Source = $hit } }
        }
      } catch {}
    }
  }
  if ($cmd) {
    try {
      $psi = New-Object System.Diagnostics.ProcessStartInfo
      $psi.FileName = $cmd.Source
      $psi.Arguments = ('-layout -nopgbrk -q -f 1 -l 6 -enc UTF-8 "{0}" -' -f $Path)
      $psi.UseShellExecute = $false
      $psi.RedirectStandardOutput = $true
      $psi.RedirectStandardError = $true
      $p = [System.Diagnostics.Process]::Start($psi)
      $text = $p.StandardOutput.ReadToEnd()
      $null = $p.StandardError.ReadToEnd()
      $p.WaitForExit()
    } catch {}
  }
  if (-not $text) {
    # Adobe Acrobat COM export to text
    try {
      $app = New-Object -ComObject AcroExch.App
      $av  = New-Object -ComObject AcroExch.AVDoc
      if ($av.Open($Path, "")) {
        $pd = $av.GetPDDoc()
        $js = $pd.GetJSObject()
        try { $null = $js.ocr.Invoke() } catch { }
        $tmp = Join-Path ([IO.Path]::GetTempPath()) (([IO.Path]::GetRandomFileName())+'.txt')
        try { $js.SaveAs($tmp, 'com.adobe.acrobat.accesstext') } catch { try { $js.SaveAs($tmp, 'com.adobe.acrobat.plain-text') } catch { } }
        try { if (Test-Path -LiteralPath $tmp) { $text = Get-Content -LiteralPath $tmp -Raw -ErrorAction SilentlyContinue } } catch {}
        try { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force } } catch {}
        $av.Close($true) | Out-Null
      }
      $app.Exit() | Out-Null
    } catch { }
  }
  if (-not $text) {
    try {
      $word = New-Object -ComObject Word.Application
      $word.Visible = $false
      $doc = $word.Documents.Open($Path, $true, $true, $true)
      $text = $doc.Content.Text
      $doc.Close($false)
      $word.Quit()
    } catch { try { if($doc){$doc.Close($false)} } catch{}; try { if($word){$word.Quit()} } catch{} }
  }
  if (-not $text) { return $Path }
  $amt = Parse-AmountFromText -Text $text
  if (-not $amt) { return $Path }
  $dir = Split-Path -Parent $Path
  $name = [IO.Path]::GetFileNameWithoutExtension($Path)
  $ext  = [IO.Path]::GetExtension($Path)
  if ($name -match ' - \d[\d,]*\.?\d{0,2}$') { return $Path }
  $new = Join-Path $dir ("{0} - {1}{2}" -f $name, ($amt -replace '[\\/:*?"<>|]','_'), $ext)
  $n=1; $cand=$new; while(Test-Path -LiteralPath $cand){ $cand = Join-Path $dir ("{0} - {1} ({2}){3}" -f $name, ($amt -replace '[\\/:*?"<>|]','_'), $n, $ext); $n++ }
  Move-Item -LiteralPath $Path -Destination $cand
  Write-Host ("Renamed with amount: {0}" -f $cand)
  return $cand
}

if (-not (Test-Path -LiteralPath $Folder)) { Write-Error "Folder not found: $Folder"; exit 1 }

$files = Get-ChildItem -LiteralPath $Folder -File -Filter *.pdf -ErrorAction SilentlyContinue
foreach($f in $files){
  $null = Try-RenameWithAmount -Path $f.FullName
}

# Prune originals when there is an amount-suffixed twin
$files = Get-ChildItem -LiteralPath $Folder -File -Filter *.pdf -ErrorAction SilentlyContinue
$map = @{}
foreach($f in $files){
  $name = [IO.Path]::GetFileNameWithoutExtension($f.Name)
  $stem = $name
  $hasAmt = $false
  if ($name -match '^(.*) - \d[\d,]*\.?\d{0,2}$') { $stem = $Matches[1]; $hasAmt = $true }
  if (-not $map.ContainsKey($stem)) { $map[$stem] = @{ originals=@(); withAmt=@() } }
  if ($hasAmt) { $map[$stem].withAmt += $f } else { $map[$stem].originals += $f }
}
foreach($k in $map.Keys){ if ($map[$k].withAmt.Count -gt 0 -and $map[$k].originals.Count -gt 0) { foreach($o in $map[$k].originals){ try { Remove-Item -LiteralPath $o.FullName -Force } catch {} } } }

Write-Host "Done."
