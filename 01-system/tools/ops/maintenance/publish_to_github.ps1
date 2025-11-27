#Requires -Version 5.1
param(
  [string]$Remote,
  [switch]$ForceReinit
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-GitAvailable {
  $git = Get-Command git -ErrorAction SilentlyContinue
  if ($git) { return $git.Source }
  # Try portable Git under tools\git
  try {
    $here = (Get-Location).Path
    $gitDir = Join-Path $here 'tools\git'
    if (Test-Path -LiteralPath $gitDir) {
      $hit = Get-ChildItem -LiteralPath $gitDir -Recurse -Filter 'git.exe' -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
      if ($hit) {
        # Prepend to PATH for this session
        $env:Path = (Join-Path (Split-Path $hit -Parent) '') + ';' + $env:Path
        return $hit
      }
    }
  } catch {}
  Write-Error "git is not installed or not on PATH. Install Git for Windows (https://gitforwindows.org/), run .\install_portable_git.ps1, or use GitHub Desktop to publish this folder."
}

try {
  $gitPath = Ensure-GitAvailable
} catch { exit 1 }

try {
  $isRepo = Test-Path -LiteralPath .git
  if ($ForceReinit -and $isRepo) {
    & git rev-parse --is-inside-work-tree | Out-Null
  }
  if (-not $isRepo) {
    & git init | Out-Null
  }
  & git add -A
  $status = (& git status --porcelain)
  if ($status) {
    & git commit -m "chore: add Outlook invoice/remittance automation scripts and docs" | Out-Null
  }
  & git branch -M main | Out-Null
  if ($Remote) {
    $hasRemote = (& git remote) -ne $null
    if (-not $hasRemote) { & git remote add origin $Remote }
    & git push -u origin main
  } else {
    Write-Host "Local repo ready. Set remote then push:"
    Write-Host "  git remote add origin https://github.com/<account>/<repo>.git"
    Write-Host "  git push -u origin main"
  }
} catch {
  Write-Error $_
  exit 1
}
