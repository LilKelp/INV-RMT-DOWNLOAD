#Requires -Version 5.1
param(
    [string]$SaveRoot,
    [string[]]$AllowedStores = @('AZhao@novabio.com','Australia AR','New Zealand AR'),
    [string]$DateFilter = '*',
    [switch]$FlattenAllowedStores,
    [switch]$RemoveOtherStores,
    [switch]$DeleteLooseRootFiles,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-UniquePath {
    param([Parameter(Mandatory)][string]$Path)
    $dir = Split-Path -Parent $Path
    $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $ext = [System.IO.Path]::GetExtension($Path)
    $candidate = $Path
    $n = 1
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $dir "$name ($n)$ext"
        $n++
    }
    return $candidate
}

if (-not $SaveRoot) {
    $SaveRoot = Join-Path -Path (Get-Location) -ChildPath 'Inv&Remit_Today'
}

if (-not (Test-Path -LiteralPath $SaveRoot)) {
    Write-Host "SaveRoot not found: $SaveRoot"
    exit 0
}

if (-not $PSBoundParameters.ContainsKey('FlattenAllowedStores')) { $FlattenAllowedStores = $true }
if (-not $PSBoundParameters.ContainsKey('RemoveOtherStores')) { $RemoveOtherStores = $true }

Write-Host "Cleanup root: $SaveRoot"
Write-Host "Allowed stores: $($AllowedStores -join ', ')"
Write-Host "Date filter: $DateFilter | Flatten: $FlattenAllowedStores | RemoveOthers: $RemoveOtherStores | DeleteLooseRootFiles: $DeleteLooseRootFiles | DryRun: $DryRun"

$dateDirs = Get-ChildItem -LiteralPath $SaveRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like $DateFilter }

foreach ($dateDir in $dateDirs) {
    Write-Host "\n=== Cleaning date: $($dateDir.Name) ==="

    # 0) Handle loose files directly under the date folder (legacy saves)
    $looseFiles = Get-ChildItem -LiteralPath $dateDir.FullName -File -ErrorAction SilentlyContinue
    foreach ($lf in $looseFiles) {
        if ($DeleteLooseRootFiles) {
            if ($DryRun) { Write-Host ("Would delete loose file: {0}" -f $lf.FullName) }
            else { Remove-Item -LiteralPath $lf.FullName -Force }
        } else {
            $miscDir = Join-Path -Path $dateDir.FullName -ChildPath '_misc'
            if (-not (Test-Path -LiteralPath $miscDir)) { if ($DryRun) { Write-Host ("Would create: {0}" -f $miscDir) } else { New-Item -ItemType Directory -Path $miscDir | Out-Null } }
            $target = Join-Path -Path $miscDir -ChildPath $lf.Name
            if (Test-Path -LiteralPath $target) { $target = Get-UniquePath -Path $target }
            if ($DryRun) { Write-Host ("Would move loose: {0} -> {1}" -f $lf.FullName, $target) }
            else { Move-Item -LiteralPath $lf.FullName -Destination $target }
        }
    }

    # 1) Flatten allowed stores: move all files up to store root, then remove empty subdirs
    foreach ($storeName in $AllowedStores) {
        $storePath = Join-Path -Path $dateDir.FullName -ChildPath $storeName
        if (-not (Test-Path -LiteralPath $storePath)) { continue }

        if ($FlattenAllowedStores) {
            Write-Host ("Flattening store: {0}" -f $storeName)
            $files = Get-ChildItem -LiteralPath $storePath -Recurse -File -ErrorAction SilentlyContinue
            foreach ($f in $files) {
                $target = Join-Path -Path $storePath -ChildPath $f.Name
                if (Test-Path -LiteralPath $target) { $target = Get-UniquePath -Path $target }
                if ($DryRun) {
                    Write-Host ("Would move: {0} -> {1}" -f $f.FullName, $target)
                } else {
                    Move-Item -LiteralPath $f.FullName -Destination $target
                }
            }
            # Remove empty subfolders
            $subdirs = Get-ChildItem -LiteralPath $storePath -Directory -Recurse -ErrorAction SilentlyContinue |
                Sort-Object FullName -Descending
            foreach ($d in $subdirs) {
                try {
                    $hasItems = (Get-ChildItem -LiteralPath $d.FullName -Force | Measure-Object).Count -gt 0
                    if (-not $hasItems) {
                        if ($DryRun) { Write-Host ("Would remove empty dir: {0}" -f $d.FullName) }
                        else { Remove-Item -LiteralPath $d.FullName -Force -Recurse }
                    }
                } catch {}
            }
        }
    }

    # 2) Remove other store folders entirely
    if ($RemoveOtherStores) {
        $storeDirs = Get-ChildItem -LiteralPath $dateDir.FullName -Directory -ErrorAction SilentlyContinue
        foreach ($s in $storeDirs) {
            $isAllowed = $false
            foreach ($allow in $AllowedStores) { if ($s.Name -ieq $allow) { $isAllowed = $true; break } }
            if (-not $isAllowed -and $s.Name -ne '_archive' -and $s.Name -ne '_misc') {
                if ($DryRun) {
                    Write-Host ("Would delete store folder: {0}" -f $s.FullName)
                } else {
                    Write-Host ("Deleting store folder: {0}" -f $s.FullName)
                    Remove-Item -LiteralPath $s.FullName -Force -Recurse
                }
            }
        }
    }
}

Write-Host "\nCleanup complete."
