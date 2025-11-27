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

function Normalize-AmountString {
  param([string]$Value)
  if (-not $Value) { return '' }
  $trimmed = $Value.Trim()
  if (-not $trimmed) { return '' }
  $isNegative = $false
  if ($trimmed.StartsWith('(') -and $trimmed.EndsWith(')')) {
    $isNegative = $true
    $trimmed = $trimmed.Trim('(', ')')
  }
  $clean = $trimmed -replace '[^\d\.,-]', ''
  if (-not $clean) { return '' }
  $clean = $clean -replace ',', ''
  if ($isNegative -and $clean -notmatch '^-') {
    $clean = "-$clean"
  }
  try {
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $number = [decimal]::Parse($clean, $culture)
    return $number.ToString('0.00', $culture)
  } catch {
    return $clean
  }
}

function Parse-DocumentReference {
  param([Parameter(Mandatory)][string]$Text)
  $patterns = @(
    '(?is)document\s+ref[\s\S]{0,120}?no[:\s]*([A-Za-z0-9-]+)',
    '(?is)reference\s+number[:\s]*([A-Za-z0-9-]+)'
  )
  foreach($p in $patterns){
    $m = [regex]::Match($Text,$p)
    if ($m.Success) {
      $raw = $m.Groups[1].Value.Trim()
      if ($raw) { return ($raw -replace '\s','') }
    }
  }
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
  $docRef = Parse-DocumentReference -Text $text
  $dir = Split-Path -Parent $Path
  $name = [IO.Path]::GetFileNameWithoutExtension($Path)
  $ext  = [IO.Path]::GetExtension($Path)
  $hasAmtSuffix = $false
  $baseName = $name
  $existingAmtSuffix = ''
  $suffixPattern = '^(.*) - ([\$]?\s*\(?-?\d[\d,]*(?:\.\d{1,2})?\)?)$'
  if ($name -match $suffixPattern) {
    $hasAmtSuffix = $true
    $baseName = $Matches[1]
    $existingAmtSuffix = $Matches[2]
  }
  $safeDoc = $null
  if ($docRef) { $safeDoc = ($docRef -replace '[\\/:*?"<>|]','_') }
  $normalizedAmt = Normalize-AmountString -Value $amt
  $normalizedExistingAmt = Normalize-AmountString -Value $existingAmtSuffix
  $amountMatches = $false
  if ($normalizedAmt -and $normalizedExistingAmt) {
    $amountMatches = ($normalizedAmt -eq $normalizedExistingAmt)
  } elseif ($existingAmtSuffix) {
    $amountMatches = (($existingAmtSuffix -replace '\s','') -eq ($amt -replace '\s',''))
  }
  if ($hasAmtSuffix -and $amountMatches) {
    if (-not $safeDoc -or $baseName -eq $safeDoc) { return $Path }
  }
  $prefix = if ($safeDoc) { $safeDoc } else { $baseName }
  $safePrefix = ($prefix -replace '[\\/:*?"<>|]','_')
  $displayAmt = if ($normalizedAmt) { $normalizedAmt } else { ($amt -replace '\s','') }
  $safeAmt = ($displayAmt -replace '[\\/:*?"<>|]','_')
  $new = Join-Path $dir ("{0} - {1}{2}" -f $safePrefix, $safeAmt, $ext)
  $n=1; $cand=$new; while(Test-Path -LiteralPath $cand){ $cand = Join-Path $dir ("{0} - {1} ({2}){3}" -f $safePrefix, $safeAmt, $n, $ext); $n++ }
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
  if ($name -match '^(.*) - ([\$]?\s*\(?-?\d[\d,]*(?:\.\d{1,2})?\)?)$') { $stem = $Matches[1]; $hasAmt = $true }
  if (-not $map.ContainsKey($stem)) { $map[$stem] = @{ originals=@(); withAmt=@() } }
  if ($hasAmt) { $map[$stem].withAmt += $f } else { $map[$stem].originals += $f }
}
foreach($k in $map.Keys){ if ($map[$k].withAmt.Count -gt 0 -and $map[$k].originals.Count -gt 0) { foreach($o in $map[$k].originals){ try { Remove-Item -LiteralPath $o.FullName -Force } catch {} } } }

Write-Host "Done."
