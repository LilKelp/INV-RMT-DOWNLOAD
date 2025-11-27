#Requires -Version 5.1
param(
  [string]$SaveRoot = $(Join-Path (Get-Location) 'Inv&Remit_Today'),
  [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-UniquePath {
  param([Parameter(Mandatory)][string]$Path)
  $dir = Split-Path -Parent $Path
  $name = [IO.Path]::GetFileNameWithoutExtension($Path)
  $ext  = [IO.Path]::GetExtension($Path)
  $cand = $Path; $n=1
  while(Test-Path -LiteralPath $cand){ $cand = Join-Path $dir ("{0} ({1}){2}" -f $name,$n,$ext); $n++ }
  return $cand
}

if (-not (Test-Path -LiteralPath $SaveRoot)) { Write-Error "SaveRoot not found: $SaveRoot"; exit 1 }

$dateDirs = Get-ChildItem -LiteralPath $SaveRoot -Directory -ErrorAction SilentlyContinue
foreach($dateDir in $dateDirs){
  # Assume current structure: SaveRoot\YYYY-MM-DD\<Store>\files
  # Move to: SaveRoot\<Store>\YYYY-MM-DD\files
  $stores = Get-ChildItem -LiteralPath $dateDir.FullName -Directory -ErrorAction SilentlyContinue
  foreach($store in $stores){
    $destRoot = Join-Path $SaveRoot $store.Name
    $destDate = Join-Path $destRoot $dateDir.Name
    if ($DryRun) { Write-Host ("Would create: {0}" -f $destDate) } else { New-Item -ItemType Directory -Path $destDate -Force | Out-Null }

    # Move files
    $items = Get-ChildItem -LiteralPath $store.FullName -Force -ErrorAction SilentlyContinue
    foreach($it in $items){
      $target = Join-Path $destDate $it.Name
      if (Test-Path -LiteralPath $target) { $target = Get-UniquePath -Path $target }
      if ($DryRun) { Write-Host ("Would move: {0} -> {1}" -f $it.FullName,$target) }
      else { Move-Item -LiteralPath $it.FullName -Destination $target }
    }
    # Remove empty source store dir
    try { if (-not $DryRun) { Remove-Item -LiteralPath $store.FullName -Force -Recurse } } catch {}
  }
  # Remove empty date dir
  try {
    $rem = Get-ChildItem -LiteralPath $dateDir.FullName -Force -ErrorAction SilentlyContinue
    if (-not $rem -or $rem.Count -eq 0) { if (-not $DryRun) { Remove-Item -LiteralPath $dateDir.FullName -Force -Recurse } }
  } catch {}
}

Write-Host "Migration complete."
