#Requires -Version 5.1
param(
  [Parameter(Mandatory=$true)][string[]]$Stores,
  [Parameter(Mandatory=$true)][string]$Date,
  [string]$TimeZoneId = 'AUS Eastern Standard Time',
  [switch]$FastScan,
  [int]$MaxItems = 400,
  [switch]$Recurse,
  [switch]$PruneOriginals,
  [switch]$Broad,
  [string[]]$AllowSenders
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Resolve-Path (Join-Path $PSScriptRoot '..\\..\\..\\..')
$runner = Join-Path $PSScriptRoot 'run_remittance_today.ps1'
if (-not (Test-Path -LiteralPath $runner)) { Write-Error "Runner not found: $runner"; exit 1 }

# Normalize date folder
$selectedDate = $null
$formats = @('yyyyMMdd','yyyy-MM-dd','yyyy/MM/dd')
foreach($fmt in $formats){ try { $selectedDate = [datetime]::ParseExact($Date,$fmt,$null) ; break } catch {} }
if (-not $selectedDate) { try { $selectedDate = [datetime]$Date } catch { Write-Error "Invalid -Date: $Date"; exit 1 } }
$dateFolder = $selectedDate.ToString('yyyy-MM-dd')

$outDir = Join-Path (Join-Path $root '03-outputs') 'remittance-runner'
$outDir = Join-Path $outDir $dateFolder
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$filesDir = Join-Path $outDir 'files'
New-Item -ItemType Directory -Force -Path $filesDir | Out-Null
$log = Join-Path $outDir ("run-" + ((Get-Date).ToString('HHmmss')) + '.log')

Start-Transcript -Path $log -Append | Out-Null
try {
  $splat = @{ Stores = $Stores; Date = $dateFolder; TimeZoneId = $TimeZoneId }
  if ($FastScan) { $splat.FastScan = $true }
  if ($MaxItems -gt 0) { $splat.MaxItems = $MaxItems }
  if ($Recurse) { $splat.Recurse = $true }
  if ($PruneOriginals) { $splat.PruneOriginals = $true } else { $splat.PruneOriginals = $true }
  if ($Broad) { $splat.Broad = $true }
  if ($AllowSenders -and $AllowSenders.Count -gt 0) { $splat.AllowSenders = $AllowSenders }
  $splat.SaveRoot = $filesDir

  & $runner @splat

  # Build manifest from presented folders
  $rows = @()
  foreach($store in $Stores){
    $folder = Join-Path $filesDir $store
    $summary = Join-Path $outDir ("summary-" + ($store -replace '[\\/:*?\"<>|]','_') + '.txt')
    if (Test-Path -LiteralPath $folder) {
      $files = @(Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue)
      "Save folder: $folder" | Out-File -FilePath $summary -Encoding UTF8
      ("{0} files" -f $files.Count) | Out-File -Append -FilePath $summary
      foreach($f in $files){
        $rows += [pscustomobject]@{
          Store = $store
          Date  = $dateFolder
          Name  = $f.Name
          FullPath = $f.FullName
          PresentedPath = $f.FullName
          SizeKB = [math]::Round($f.Length/1kb,2)
        }
      }
    } else {
      "Save folder: $folder (missing)" | Out-File -FilePath $summary -Encoding UTF8
    }
  }
  if ($rows.Count -gt 0) {
    $csv = Join-Path $outDir 'manifest.csv'
    $rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
  }
} finally {
  Stop-Transcript | Out-Null
}

Write-Host ("Artifacts: {0}" -f $outDir)
