#Requires -Version 5.1
param(
  [string]$InstallDir = $(Join-Path (Get-Location) 'tools\git')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

function Get-LatestPortableGitUrl {
  try {
    $api = 'https://api.github.com/repos/git-for-windows/git/releases/latest'
    $resp = Invoke-RestMethod -Uri $api -Headers @{ 'User-Agent' = 'ps-portable-git-installer' } -TimeoutSec 60
    foreach ($asset in $resp.assets) {
      if ($asset.name -match '^PortableGit-.*64-bit\.7z\.exe$') { return $asset.browser_download_url }
    }
    return ''
  } catch { return '' }
}

try {
  $url = Get-LatestPortableGitUrl
  if (-not $url) { throw 'Could not resolve latest PortableGit URL (GitHub API).' }
  $tmp = Join-Path $env:TEMP ('portablegit-'+[IO.Path]::GetRandomFileName()+'.exe')
  Write-Host ("Downloading: {0}" -f $url)
  Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing -TimeoutSec 300
  if (Test-Path -LiteralPath $InstallDir) {
    Write-Host ("Clearing existing: {0}" -f $InstallDir)
    Remove-Item -LiteralPath $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
  }
  New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
  Write-Host ("Extracting to: {0}" -f $InstallDir)
  & $tmp -y -o"$InstallDir" | Out-Null
  Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
  # Report git path
  $git = Get-ChildItem -LiteralPath $InstallDir -Recurse -Filter 'git.exe' -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
  if ($git) { Write-Host ("git: {0}" -f $git) } else { Write-Warning 'git.exe not found after extraction.' }
} catch {
  Write-Error $_
  Write-Host 'If network download is blocked, manually download PortableGit-*-64-bit.7z.exe from:'
  Write-Host '  https://github.com/git-for-windows/git/releases'
  Write-Host ("Then extract to: {0}" -f $InstallDir)
  exit 1
}

