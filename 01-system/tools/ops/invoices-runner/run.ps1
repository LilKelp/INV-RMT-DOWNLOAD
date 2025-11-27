#Requires -Version 5.1
param(
  [Parameter(Mandatory=$true)][string]$Date,
  [string]$TimeZoneId = 'AUS Eastern Standard Time',
  [switch]$Recurse,
  [switch]$Broad,
  [string[]]$AllowSenders,
  [switch]$FastScan,
  [int]$MaxItems = 400
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Resolve-Path (Join-Path $PSScriptRoot '..\\..\\..\\..')
$runner = Join-Path $PSScriptRoot 'run_invoices_today.ps1'
if (-not (Test-Path -LiteralPath $runner)) {
  Write-Error "Runner not found: $runner"
  exit 1
}

$selectedDate = $null
$formats = @('yyyyMMdd','yyyy-MM-dd','yyyy/MM/dd')
foreach($fmt in $formats){ try { $selectedDate = [datetime]::ParseExact($Date,$fmt,$null) ; break } catch {} }
if (-not $selectedDate) { try { $selectedDate = [datetime]$Date } catch { Write-Error "Invalid -Date: $Date"; exit 1 } }
$dateFolder = $selectedDate.ToString('yyyy-MM-dd')

$outDir = Join-Path (Join-Path $root '03-outputs') 'invoices-runner'
$outDir = Join-Path $outDir $dateFolder
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$filesDir = Join-Path $outDir 'files'
New-Item -ItemType Directory -Force -Path $filesDir | Out-Null
$log = Join-Path $outDir ("run-" + ((Get-Date).ToString('HHmmss')) + '.log')

Start-Transcript -Path $log -Append | Out-Null
try {
  $splat = @{ Date = $dateFolder; TimeZoneId = $TimeZoneId; SaveRoot = $filesDir }
  if ($Recurse) { $splat.Recurse = $true }
  if ($Broad) { $splat.Broad = $true }
  if ($AllowSenders) {
    $senderList = @($AllowSenders | Where-Object { $_ })
    if ($senderList.Count -gt 0) { $splat.AllowSenders = $senderList }
  }
  if ($FastScan) { $splat.FastScan = $true }
  if ($MaxItems -gt 0) { $splat.MaxItems = $MaxItems }
  & $runner @splat

  $store = 'AZhao@novabio.com'
  $folder = Join-Path $filesDir $store
  $summary = Join-Path $outDir 'summary.txt'
  if (Test-Path -LiteralPath $folder) {
    $files = @(Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue)
    "Save folder: $folder" | Out-File -FilePath $summary -Encoding UTF8
    ("{0} files" -f $files.Count) | Out-File -Append -FilePath $summary

    # Filter out obvious non-invoices and collapse duplicates by stem (prefer unsuffixed name)
    $neg = { param($name) $l=$name.ToLower(); return ($l -like '*form*' -or $l -like '*supplier*form*' -or $l -like '*statement*' -or $l -like '*stmt*' -or $l -like '*purchase*order*' -or $l -like '*purchaseorder*' -or $l -like '* order *' -or $l -like '*remit*' -or $l -like '*remittance*' -or $l -like '*advice*') }
    $byStem = @{}
    foreach($f in $files){
      if (& $neg $f.Name) { continue }
      $stem = [IO.Path]::GetFileNameWithoutExtension($f.Name)
      $stemNorm = ($stem -replace ' \(\d+\)$','')
      if (-not $byStem.ContainsKey($stemNorm)) { $byStem[$stemNorm] = @{ nosuffix=$null; withsuffix=@() } }
      if ($stem -match ' \(\d+\)$') { $byStem[$stemNorm].withsuffix += $f } else { $byStem[$stemNorm].nosuffix = $f }
    }
    $picks = @()
    foreach($k in $byStem.Keys){
      if ($byStem[$k].nosuffix) { $picks += $byStem[$k].nosuffix }
      elseif ($byStem[$k].withsuffix.Count -gt 0) { $picks += $byStem[$k].withsuffix | Select-Object -First 1 }
    }
    $rows = @()
    foreach($f in $picks){
      $rows += [pscustomobject]@{ Name=$f.Name; FullPath=$f.FullName; PresentedPath=$f.FullName; SizeKB=[math]::Round($f.Length/1kb,2) }
    }
    if ($rows.Count -gt 0) { $rows | Export-Csv -NoTypeInformation -Encoding UTF8 -Path (Join-Path $outDir 'manifest.csv') }
  } else {
    "Save folder: $folder (missing)" | Out-File -FilePath $summary -Encoding UTF8
  }
} finally {
  Stop-Transcript | Out-Null
}

Write-Host ("Artifacts: {0}" -f $outDir)
