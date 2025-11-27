#Requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$Folder)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Resolve-Path (Join-Path $PSScriptRoot '..\\..\\..\\..')
$script = Join-Path $PSScriptRoot 'rename_amount_in_folder.ps1'
if (-not (Test-Path -LiteralPath $script)) { Write-Error "Script not found: $script"; exit 1 }

$outDir = Join-Path (Join-Path $root '03-outputs') 'remit-rename-amount'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$log = Join-Path $outDir ("run-" + ((Get-Date).ToString('yyyyMMdd-HHmmss')) + '.log')

Start-Transcript -Path $log -Append | Out-Null
try {
  & $script -Folder $Folder
} finally {
  Stop-Transcript | Out-Null
}

Write-Host ("Artifacts: {0}" -f $outDir)
