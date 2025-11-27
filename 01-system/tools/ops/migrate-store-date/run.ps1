#Requires -Version 5.1
param([switch]$DryRun)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Resolve-Path (Join-Path $PSScriptRoot '..\\..\\..\\..')
$script = Join-Path $PSScriptRoot 'migrate_to_store_date.ps1'
if (-not (Test-Path -LiteralPath $script)) { Write-Error "Script not found: $script"; exit 1 }

$outDir = Join-Path (Join-Path $root '03-outputs') 'migrate-store-date'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$log = Join-Path $outDir ("run-" + ((Get-Date).ToString('yyyyMMdd-HHmmss')) + '.log')

Start-Transcript -Path $log -Append | Out-Null
try {
  if ($DryRun) { & $script -DryRun } else { & $script }
} finally {
  Stop-Transcript | Out-Null
}

Write-Host ("Artifacts: {0}" -f $outDir)
