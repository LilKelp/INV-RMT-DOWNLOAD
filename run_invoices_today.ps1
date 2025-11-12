#Requires -Version 5.1
param(
  [string]$SaveRoot,
  [switch]$Recurse,
  [string]$Date,
  [string]$TimeZoneId = 'AUS Eastern Standard Time'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-OutlookApp {
  try { return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { return New-Object -ComObject Outlook.Application }
}

function Get-PdfToTextPath {
  param([string]$Provided)
  if ($Provided) { return $Provided }
  $cmd = Get-Command pdftotext -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }
  $common = @(
    "C:\\Program Files\\poppler\\bin\\pdftotext.exe",
    "C:\\Program Files (x86)\\poppler\\bin\\pdftotext.exe",
    "$env:ProgramFiles\\poppler\\bin\\pdftotext.exe",
    "$env:ProgramFiles(x86)\\poppler\\bin\\pdftotext.exe"
  )
  foreach($p in $common){ if(Test-Path -LiteralPath $p){ return $p } }
  return $null
}

function Get-FileText {
  param([Parameter(Mandatory)][string]$Path,[string]$PdfToText)
  $text=''
  if ($PdfToText) {
    try {
      $psi = New-Object System.Diagnostics.ProcessStartInfo
      $psi.FileName = $PdfToText
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
    try {
      $word = New-Object -ComObject Word.Application
      $word.Visible = $false
      $doc = $word.Documents.Open($Path, $true, $true, $true)
      $text = $doc.Content.Text
      $doc.Close($false)
      $word.Quit()
    } catch { try{ if($doc){$doc.Close($false)} }catch{}; try{ if($word){$word.Quit()} }catch{} }
  }
  return $text
}

function Extract-InvoiceNumber {
  param(
    [string]$FileName,
    [string]$Subject,
    [string]$Text
  )
  $candidates = @()
  if ($FileName) { $candidates += $FileName }
  if ($Subject)  { $candidates += $Subject }
  if ($Text)     { $candidates += $Text }
  foreach ($src in $candidates) {
    # Prefer explicit labels
    $m = [regex]::Match($src,'(?im)\b(?:invoice\s*(?:no\.?|number|#)|inv(?:oice)?\s*(?:no\.?|number|#))\s*[:\-]?\s*([A-Z0-9][A-Z0-9\-\/]{3,})')
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
    # Common pattern like INV-123456 or 0120002755
    $m = [regex]::Match($src,'(?im)\b(?:INV[-_ ]?)?([A-Z0-9]{2,3}-?\d{5,})\b')
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
  }
  return ''
}

function Get-SafeSegment {
  param([Parameter(Mandatory)][string]$Text,[int]$MaxLength=80)
  $t = ($Text).Trim() -replace '[\\/:*?"<>|]','_'
  if ($t.Length -gt $MaxLength) { $t = $t.Substring(0,$MaxLength) }
  if ([string]::IsNullOrWhiteSpace($t)) { return '_' } else { return $t }
}

function Get-SenderSmtp {
  param([Parameter(Mandatory)]$Mail)
  try { if ($Mail.SenderEmailType -eq 'EX') { $ex = $Mail.Sender.GetExchangeUser(); if ($ex -and $ex.PrimarySmtpAddress) { return $ex.PrimarySmtpAddress } } } catch {}
  try { $v = $Mail.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001E'); if ($v) { return [string]$v } } catch {}
  try { return [string]$Mail.SenderEmailAddress } catch { return '' }
}

function Load-SupplierMap {
  param([string]$Path)
  $result = @{}
  if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $result }
  try {
    $obj = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $Path -Raw)
    if ($obj -ne $null) {
      foreach ($prop in $obj.PSObject.Properties) {
        $k = ([string]$prop.Name).ToLower()
        $v = [string]$prop.Value
        $result[$k] = $v
      }
    }
  } catch { }
  return $result
}

function Derive-SupplierName {
  param([Parameter(Mandatory)]$Mail,[hashtable]$Map)
  $sender = Get-SenderSmtp -Mail $Mail
  $senderLower = ([string]$sender).ToLower()
  if ($Map -and $Map.ContainsKey($senderLower)) { return [string]$Map[$senderLower] }
  $domain = ''
  if ($senderLower -match '@(.+)$') { $domain = $Matches[1] }
  if ($domain -and $Map -and $Map.ContainsKey($domain)) { return [string]$Map[$domain] }
  # Fallbacks: SenderName → domain second-level label
  try { $nm = [string]$Mail.SenderName; if ($nm) { return $nm } } catch {}
  if ($domain) {
    $parts = $domain -split '\.'
    if ($parts.Length -ge 2) { return $parts[$parts.Length-2] }
    return $domain
  }
  return 'Supplier'
}

try {
  if (-not $SaveRoot) { $SaveRoot = Join-Path (Get-Location) 'Inv&Remit_Today' }
  $selectedDate = $null
  if ($PSBoundParameters.ContainsKey('Date') -and $Date) {
    $formats = @('yyyyMMdd','yyyy-MM-dd','yyyy/MM/dd')
    foreach($fmt in $formats){ try { $selectedDate = [datetime]::ParseExact($Date,$fmt,$null) ; break } catch {} }
    if (-not $selectedDate) { try { $selectedDate = [datetime]$Date } catch { $selectedDate = (Get-Date) } }
  } else { $selectedDate = (Get-Date) }
  $selectedDate = $selectedDate.Date
  # Timezone window
  try { $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId) } catch { $tz = [System.TimeZoneInfo]::Local }
  $startTZ = $selectedDate
  $endTZ   = $selectedDate.AddDays(1)
  $start = [System.TimeZoneInfo]::ConvertTime($startTZ, $tz, [System.TimeZoneInfo]::Local)
  $end   = [System.TimeZoneInfo]::ConvertTime($endTZ,   $tz, [System.TimeZoneInfo]::Local)
  $fStart = $start.ToString('MM/dd/yyyy hh:mm tt')
  $fEnd   = $end.ToString('MM/dd/yyyy hh:mm tt')

  $dateFolder = $selectedDate.ToString('yyyy-MM-dd')

  $outlook = Get-OutlookApp
  $ns = $outlook.GetNamespace('MAPI')
  $storeName = 'AZhao@novabio.com'
  $store = ($ns.Folders | Where-Object { $_.Name -eq $storeName })
  if (-not $store) { Write-Error "Store not found: $storeName"; exit 1 }
  $rootFolder = $store.Folders.Item('Inbox')
  if (-not $rootFolder) { Write-Error "Inbox not found in $storeName"; exit 1 }

  $restriction = "[ReceivedTime] >= '$fStart' AND [ReceivedTime] < '$fEnd'"
  $subjectRegex = [regex]::new('invoice','IgnoreCase')
  $allowedExtRegex = [regex]::new('\.(pdf)$','IgnoreCase')
  $imageExtRegex = [regex]::new('\.(png|jpg|jpeg|gif|bmp|svg|webp)$','IgnoreCase')

  # No invoice renaming; keep original filenames

  $queue = New-Object System.Collections.Generic.Queue[Object]
  $queue.Enqueue($rootFolder)
  while($queue.Count -gt 0){
    $folder = $queue.Dequeue()
    if ($Recurse) { foreach($sub in $folder.Folders){ $queue.Enqueue($sub) } }
    $items = $folder.Items; $items.IncludeRecurrences=$true; $items.Sort('[ReceivedTime]'); $todayItems=$items.Restrict($restriction)
    foreach($item in $todayItems){
      $isMail = $false; try { if ($item -and $item.Class -eq 43) { $isMail=$true } } catch {}
      if (-not $isMail) { continue }
      $subj=''; try { $subj=[string]$item.Subject } catch {}
      $atts = $item.Attachments
      if (-not $atts -or $atts.Count -le 0) { continue }
      $subjectMatch = $false; try { $subjectMatch = $subjectRegex.IsMatch($subj) } catch { $subjectMatch = $false }
      $attNameMatch = $false
      try { for($ti=1; $ti -le $atts.Count; $ti++){ $tfn=[string]$atts.Item($ti).FileName; if ($tfn -match '(?i)\b(inv|invoice)\b'){ $attNameMatch=$true; break } } } catch { $attNameMatch=$false }
      if (-not ($subjectMatch -or $attNameMatch)) { continue }
      $storeRoot = Join-Path $SaveRoot $storeName
      $storeDir  = Join-Path $storeRoot $dateFolder
      New-Item -ItemType Directory -Path $storeDir -Force | Out-Null
      for($i=1; $i -le $atts.Count; $i++){
        $att = $atts.Item($i); if(-not $att){ continue }
        $fn = [string]$att.FileName
        $isInline = $false
        try { $cid = $att.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F'); if($cid){$isInline=$true} } catch {}
        if ($isInline -and $imageExtRegex.IsMatch($fn)) { continue }
        if (-not $allowedExtRegex.IsMatch($fn)) { continue }
        $safe = $fn -replace '[\\/:*?"<>|]','_'
        $path = Join-Path $storeDir $safe
        $n=1; $cand=$path; while(Test-Path -LiteralPath $cand){ $cand = Join-Path $storeDir ("{0} ({1}){2}" -f ([IO.Path]::GetFileNameWithoutExtension($safe)), $n, [IO.Path]::GetExtension($safe)); $n++ }
        try { $att.SaveAsFile($cand); Write-Host "Saved: $cand" } catch { Write-Warning "Failed to save attachment '$fn': $($_.Exception.Message)" }
      }
    }
  }
  Write-Host ("Save folder: {0}" -f (Join-Path (Join-Path $SaveRoot $storeName) $dateFolder))
} catch {
  Write-Error $_
  exit 1
}
