#Requires -Version 5.1
param(
  [string]$SaveRoot,
  [int]$LookbackDays = 0,
  [switch]$Recurse,
  [string]$Date,
  [string]$TimeZoneId = 'AUS Eastern Standard Time',
  [switch]$PruneOriginals,
  [string[]]$Stores = @('Australia AR','New Zealand AR'),
  [int]$MaxItems = 300,
  [switch]$FastScan,
  [switch]$Broad,
  [string[]]$AllowSenders = @('SharedServicesAccountsPayable@act.gov.au','finance@yourremittance.com.au','noreply_remittances@mater.org.au')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-OutlookApp {
  try { return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { return New-Object -ComObject Outlook.Application }
}

function Get-PdfToTextPath {
  $cmd = Get-Command pdftotext -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }
  $here = (Get-Location).Path
  $common = @(
    (Join-Path $here 'tools\poppler\Library\bin\pdftotext.exe'),
    (Join-Path $here 'tools\poppler\bin\pdftotext.exe'),
    "C:\\Program Files\\poppler\\bin\\pdftotext.exe",
    "C:\\Program Files (x86)\\poppler\\bin\\pdftotext.exe",
    "$env:ProgramFiles\\poppler\\bin\\pdftotext.exe",
    "$env:ProgramFiles(x86)\\poppler\\bin\\pdftotext.exe"
  )
  foreach($p in $common){ if(Test-Path -LiteralPath $p){ return $p } }
  try {
    $pop = Join-Path $here 'tools\poppler'
    if (Test-Path -LiteralPath $pop) {
      $hit = Get-ChildItem -LiteralPath $pop -Recurse -Filter 'pdftotext.exe' -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
      if ($hit) { return $hit }
    }
  } catch {}
  return $null
}

function Parse-AmountFromText {
  param([Parameter(Mandatory)][string]$Text)
  $patterns = @(
    '(?im)\b(?:grand\s+total|total\s+amount|amount\s+paid|total\s+paid|net\s+total|invoice\s+total)\D*(\$?\s?-?\d{1,3}(?:,\d{3})*(?:\.\d{2})?)',
    '(?im)\b(?:AUD|NZD)\s*(\$?\s?-?\d{1,3}(?:,\d{3})*(?:\.\d{2})?)',
    '(?im)\btotal(?:\s+amount)?\b\D*(\$?\s?-?\d{1,3}(?:,\d{3})*(?:\.\d{2})?)'
  )
  foreach($p in $patterns){ $m=[regex]::Match($Text,$p); if($m.Success){ $raw=$m.Groups[1].Value.Trim(); return ($raw -replace '\s','') } }
  return ''
}

function Save-MailAsMsg {
  param(
    [Parameter(Mandatory)][object]$Mail,
    [Parameter(Mandatory)][string]$TargetDir
  )
  try {
    $subj = ''
    try { $subj = [string]$Mail.Subject } catch {}
    if ([string]::IsNullOrWhiteSpace($subj)) { $subj = 'Remittance' }
    $text = ''
    try { $text = $subj + "`n" + [string]$Mail.Body } catch {}
    $amt = if ($text) { Parse-AmountFromText -Text $text } else { '' }
    $safeSubj = ($subj -replace '[\\/:*?"<>|]','_')
    if ($safeSubj.Length -gt 120) { $safeSubj = $safeSubj.Substring(0,120) }
    $fileName = if ($amt) { "{0} - {1}.msg" -f $safeSubj, ($amt -replace '[\\/:*?"<>|]','_') } else { "{0}.msg" -f $safeSubj }
    $dest = Join-Path $TargetDir $fileName
    $n=1; while (Test-Path -LiteralPath $dest) {
      $base = [IO.Path]::GetFileNameWithoutExtension($fileName)
      $dest = Join-Path $TargetDir ("{0} ({1}).msg" -f $base,$n)
      $n++
    }
    $Mail.SaveAs($dest, 3)
    Write-Host ("Saved MSG: {0}" -f $dest)
    return $true
  } catch {
    Write-Warning ("Failed to save MSG: {0}" -f $_.Exception.Message)
    return $false
  }
}

function Get-AmountFromMail {
  param([Parameter(Mandatory)]$Mail)
  try {
    $txt = ''
    try { $txt = [string]$Mail.Subject } catch {}
    try { $txt += "`n" + [string]$Mail.Body } catch {}
    if ($txt) { return (Parse-AmountFromText -Text $txt) } else { return '' }
  } catch { return '' }
}

function Try-RenameWithAmount {
  param([Parameter(Mandatory)][string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { return $Path }
  $ext = [System.IO.Path]::GetExtension($Path)
  if ($ext -notmatch '^\.pdf$' -and $ext -notmatch '^\.PDF$') { return $Path }
  $tool = Get-PdfToTextPath
  $text = ''
  if ($tool) {
    try {
      $psi = New-Object System.Diagnostics.ProcessStartInfo
      $psi.FileName = $tool
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
      # Adobe Acrobat COM text export (requires Acrobat Pro, not Reader)
      $app = New-Object -ComObject AcroExch.App
      $av  = New-Object -ComObject AcroExch.AVDoc
      if ($av.Open($Path, "")) {
        $pd = $av.GetPDDoc()
        $js = $pd.GetJSObject()
        try { $null = $js.ocr.Invoke() } catch { }
        $tmp = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(),([System.IO.Path]::GetRandomFileName()+'.txt'))
        try { $js.SaveAs($tmp, 'com.adobe.acrobat.accesstext') } catch { try { $js.SaveAs($tmp, 'com.adobe.acrobat.plain-text') } catch { } }
        try { if (Test-Path -LiteralPath $tmp) { $text = Get-Content -LiteralPath $tmp -Raw -ErrorAction SilentlyContinue } } catch {}
        try { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force } } catch {}
        $av.Close($true) | Out-Null
      }
      $app.Exit() | Out-Null
    } catch { }
  }
  if (-not $text) {
    try {
      $word = New-Object -ComObject Word.Application
      $word.Visible = $false
      $doc = $word.Documents.Open($Path, $true, $true, $true)
      $text = $doc.Content.Text
      $doc.Close($false)
      $word.Quit()
    } catch { try { if($doc){$doc.Close($false)} } catch{}; try { if($word){$word.Quit()} } catch{} }
  }
  if (-not $text) { return $Path }
  $amt = Parse-AmountFromText -Text $text
  if (-not $amt) { return $Path }
  $dir = Split-Path -Parent $Path
  $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
  $ext = [System.IO.Path]::GetExtension($Path)
  $new = Join-Path $dir ("{0} - {1}{2}" -f $name, ($amt -replace '[\\/:*?"<>|]','_'), $ext)
  $n=1; $cand=$new; while(Test-Path -LiteralPath $cand){ $cand = Join-Path $dir ("{0} - {1} ({2}){3}" -f $name, ($amt -replace '[\\/:*?"<>|]','_'), $n, $ext); $n++ }
  Move-Item -LiteralPath $Path -Destination $cand
  Write-Host ("Renamed with amount: {0}" -f $cand)
  return $cand
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
  $dateFolder = $selectedDate.ToString('yyyy-MM-dd')

  $outlook = Get-OutlookApp
  $ns = $outlook.GetNamespace('MAPI')
  $stores = $Stores

  # Build time window using specified timezone, then convert to local for MAPI Restrict
  try { $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId) } catch { $tz = [System.TimeZoneInfo]::Local }
  $startTZ = $selectedDate
  $endTZ   = $selectedDate.AddDays(1)
  if (-not ($PSBoundParameters.ContainsKey('Date') -and $Date)) {
    $nowTZ = [System.TimeZoneInfo]::ConvertTime([datetime]::UtcNow, [System.TimeZoneInfo]::Utc, $tz)
    $startTZ = $nowTZ.Date.AddDays(-1 * [Math]::Max(0,$LookbackDays))
    $endTZ   = $nowTZ.Date.AddDays(1)
  }
  $start = [System.TimeZoneInfo]::ConvertTime($startTZ, $tz, [System.TimeZoneInfo]::Local)
  $end   = [System.TimeZoneInfo]::ConvertTime($endTZ,   $tz, [System.TimeZoneInfo]::Local)
  $fStart = $start.ToString('MM/dd/yyyy hh:mm tt')
  $fEnd   = $end.ToString('MM/dd/yyyy hh:mm tt')
  $restriction = "[ReceivedTime] >= '$fStart' AND [ReceivedTime] < '$fEnd'"

  $subjectRegex = [regex]::new('remittance','IgnoreCase')
  $fileNameRegex = [regex]::new('remit','IgnoreCase')
  $allowedExtRegex = [regex]::new('\.(pdf)$','IgnoreCase')
  $imageExtRegex = [regex]::new('\.(png|jpg|jpeg|gif|bmp|svg|webp)$','IgnoreCase')
  $blockedAddrs = @('NZ-AR@NOVABIO.COM','AU-AR@NOVABIO.COM','au-orders@novabio.com','azhao@novabio.com') | ForEach-Object { $_.ToLower() }
  $allowAddrs = $AllowSenders | ForEach-Object { $_.ToLower() }
  $seen = New-Object 'System.Collections.Generic.HashSet[string]'

  function Remove-OriginalsWithoutAmountSuffix {
    param([Parameter(Mandatory)][string]$Dir)
    try {
      $files = Get-ChildItem -LiteralPath $Dir -File -ErrorAction SilentlyContinue
      $byStem = @{}
      foreach ($f in $files) {
        $name = [IO.Path]::GetFileNameWithoutExtension($f.Name)
        $stem = $name
        $hasAmt = $false
        if ($name -match '^(.*) - \d[\d,]*\.?\d{0,2}$') { $stem = $Matches[1]; $hasAmt = $true }
        if (-not $byStem.ContainsKey($stem)) { $byStem[$stem] = @{ originals=@(); withAmt=@() } }
        if ($hasAmt) { $byStem[$stem].withAmt += $f } else { $byStem[$stem].originals += $f }
      }
      foreach ($k in $byStem.Keys) {
        if ($byStem[$k].withAmt.Count -gt 0 -and $byStem[$k].originals.Count -gt 0) {
          foreach ($o in $byStem[$k].originals) { try { Remove-Item -LiteralPath $o.FullName -Force } catch {} }
        }
      }
    } catch {}
  }

  foreach($storeName in $stores){
    $store = ($ns.Folders | Where-Object { $_.Name -eq $storeName })
    if (-not $store) { Write-Host "Store not found: $storeName"; continue }
    $rootFolder = $store.Folders.Item('Inbox')
    if (-not $rootFolder) { continue }
    $storeRoot = Join-Path $SaveRoot $storeName
    $saveDir   = Join-Path $storeRoot $dateFolder
    New-Item -ItemType Directory -Path $saveDir -Force | Out-Null
    $queue = New-Object System.Collections.Generic.Queue[Object]
    $queue.Enqueue($rootFolder)
    while($queue.Count -gt 0){
      $folder = $queue.Dequeue()
      if ($Recurse) { foreach($sub in $folder.Folders){ $queue.Enqueue($sub) } }

      $items = $folder.Items
      $items.IncludeRecurrences = $true
      $items.Sort('[ReceivedTime]')
      
      $iter = @()
      if ($FastScan -or $PSBoundParameters.ContainsKey('MaxItems')) {
        $cnt = 0; try { $cnt = [int]$items.Count } catch { $cnt = 0 }
        $startIdx = [Math]::Max(1, $cnt - [Math]::Max(1,$MaxItems) + 1)
        for ($idx = $cnt; $idx -ge $startIdx; $idx--) { $iter += $items.Item($idx) }
      } else {
        try { $items = $items.Restrict($restriction) } catch {}
        $iter = $items
      }

      foreach($item in $iter){
        $isMail = $false; try { if ($item -and $item.Class -eq 43) { $isMail=$true } } catch {}
        if (-not $isMail) { continue }
        try { $rt = [datetime]$item.ReceivedTime } catch { $rt = $null }
        if ($rt -and ($rt -lt $start -or $rt -ge $end)) { continue }
        $sender = ''
        try { $sender = [string]$item.SenderEmailAddress } catch {}
        $senderLower = $sender.ToLower()
        $allowOverride = ($allowAddrs -contains $senderLower)
        if (-not $allowOverride) {
          if ($senderLower -like '*@novabio.com' -or $blockedAddrs -contains $senderLower) { continue }
        }
        $subj=''; try { $subj=[string]$item.Subject } catch {}
        if ($allowOverride) {
          $attsAO = $item.Attachments
          $hasPdfAO = $false
          try {
            if ($attsAO -and $attsAO.Count -gt 0) {
              for($ai=1; $ai -le $attsAO.Count; $ai++){ $afn=[string]$attsAO.Item($ai).FileName; if ($allowedExtRegex.IsMatch($afn)) { $hasPdfAO = $true; break } }
            }
          } catch { $hasPdfAO = $false }
          if (-not $hasPdfAO) {
          [void](Save-MailAsMsg -Mail $item -TargetDir $saveDir)
          continue
        }
          # If it has PDF attachments, fall through to normal PDF save flow
        }
        # Only proceed if subject mentions remittance or at least one attachment filename does
        $subjectMatch = $false; try { $subjectMatch = $subjectRegex.IsMatch($subj) } catch { $subjectMatch = $false }
        $atts = $item.Attachments
        if (-not $atts -or $atts.Count -le 0) { continue }
        $nameMatch = $false
        try {
          for($ti=1; $ti -le $atts.Count; $ti++){ $tfn = [string]$atts.Item($ti).FileName; if ($fileNameRegex.IsMatch($tfn)) { $nameMatch = $true; break } }
        } catch { $nameMatch = $false }
        if (-not $Broad -and -not $allowOverride) {
          if (-not ($subjectMatch -or $nameMatch)) { continue }
        }
        $storeDir = $saveDir
        for($i=1; $i -le $atts.Count; $i++){
          $att = $atts.Item($i); if(-not $att){ continue }
          $fn = [string]$att.FileName
          $isInline = $false
          try { $cid = $att.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F'); if($cid){$isInline=$true} } catch {}
          if ($isInline -and $imageExtRegex.IsMatch($fn)) { continue }
          if (-not $allowedExtRegex.IsMatch($fn)) { continue }
          $safe = $fn -replace '[\\/:*?"<>|]','_'
          $attSize = 0; try { $attSize = [int]$att.Size } catch {}
          $key = ("{0}|{1}|{2}|{3}" -f ([string]$item.EntryID), $i, $safe.ToLower(), $attSize)
          if ($seen.Contains($key)) { continue } else { [void]$seen.Add($key) }

          $amtFromMail = Get-AmountFromMail -Mail $item
          if ($amtFromMail) {
            $base = [IO.Path]::GetFileNameWithoutExtension($safe)
            $ext  = [IO.Path]::GetExtension($safe)
            $cand = Join-Path $storeDir ("{0} - {1}{2}" -f $base, ($amtFromMail -replace '[\\/:*?"<>|]','_'), $ext)
            $n=1; while(Test-Path -LiteralPath $cand){ $cand = Join-Path $storeDir ("{0} - {1} ({2}){3}" -f $base, ($amtFromMail -replace '[\\/:*?"<>|]','_'), $n, $ext); $n++ }
          } else {
            $path = Join-Path $storeDir $safe
            $n=1; $cand=$path; while(Test-Path -LiteralPath $cand){ $cand = Join-Path $storeDir ("{0} ({1}){2}" -f ([IO.Path]::GetFileNameWithoutExtension($safe)), $n, [IO.Path]::GetExtension($safe)); $n++ }
          }

          try {
            $att.SaveAsFile($cand)
            if (-not $amtFromMail) {
              $newPath = Try-RenameWithAmount -Path $cand
              if ($newPath -and (Test-Path -LiteralPath $newPath)) {
                if ($newPath -ne $cand -and (Test-Path -LiteralPath $cand)) { Remove-Item -LiteralPath $cand -Force }
                Write-Host "Saved: $newPath"
              } else {
                Write-Host "Saved: $cand"
              }
            } else {
              Write-Host "Saved: $cand"
            }
          } catch { Write-Warning "Failed to save attachment '$fn': $($_.Exception.Message)" }
        }
      }
    }
  }
  if ($PSBoundParameters.ContainsKey('PruneOriginals') -and $PruneOriginals) {
    foreach($storeName in $stores){
      $storeDir = Join-Path $saveDir $storeName
      if (Test-Path -LiteralPath $storeDir) { Remove-OriginalsWithoutAmountSuffix -Dir $storeDir }
    }
  }
  Write-Host ("Save folder: {0}" -f $saveDir)
} catch {
  Write-Error $_
  exit 1
}
