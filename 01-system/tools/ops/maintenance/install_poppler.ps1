#Requires -Version 5.1
param(
  [string]$InstallDir = $(Join-Path (Get-Location) 'tools\poppler')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {}

function Get-LatestPopplerZipUrl {
  try {
    $api = 'https://api.github.com/repos/oschwartz10612/poppler-windows/releases/latest'
    $resp = Invoke-RestMethod -Uri $api -Headers @{ 'User-Agent' = 'ps-poppler-installer' } -TimeoutSec 60
    foreach ($asset in $resp.assets) {
      if ($asset.name -match '^Release-.*\.zip$') { return $asset.browser_download_url }
    }
    return ''
  } catch { return '' }
}

try {
  $url = Get-LatestPopplerZipUrl
  if (-not $url) { throw 'Could not resolve latest Poppler release URL (GitHub API).' }
  $tmp = Join-Path $env:TEMP ('poppler-'+[IO.Path]::GetRandomFileName()+'.zip')
  Write-Host ("Downloading: {0}" -f $url)
  Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing -TimeoutSec 300
  if (Test-Path -LiteralPath $InstallDir) {
    Write-Host ("Clearing existing: {0}" -f $InstallDir)
    Remove-Item -LiteralPath $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
  }
  New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
  Write-Host ("Extracting to: {0}" -f $InstallDir)
  Expand-Archive -Path $tmp -DestinationPath $InstallDir -Force
  Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
  Write-Host 'Poppler installed.'
  $bin = Join-Path $InstallDir 'Library\bin\pdftotext.exe'
  if (-not (Test-Path -LiteralPath $bin)) { $bin = Join-Path $InstallDir 'bin\pdftotext.exe' }
  if (Test-Path -LiteralPath $bin) { Write-Host ("pdftotext: {0}" -f $bin) } else { Write-Warning 'pdftotext.exe not found after extraction.' }
} catch {
  Write-Error $_
  Write-Host 'If network download is blocked, manually download a Release-*.zip from:'
  Write-Host '  https://github.com/oschwartz10612/poppler-windows/releases'
  Write-Host ("Then extract to: {0}" -f $InstallDir)
  exit 1
}

