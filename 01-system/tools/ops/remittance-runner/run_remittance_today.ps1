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
  [string[]]$AllowSenders = @(
    'AccountsPayable@yourremittance.com.au',
    'SharedServicesAccountsPayable@act.gov.au',
    'finance@yourremittance.com.au',
    'noreply_remittances@mater.org.au',
    'payments@nzdf.mil.nz',
    'HSNSW-scnremit@gateway2.messagexchange.com'
  )
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-WorkspaceRoot {
  param([Parameter(Mandatory)][string]$StartPath)
  $current = (Resolve-Path -LiteralPath $StartPath).Path
  while ($true) {
    if (Test-Path -LiteralPath (Join-Path $current 'AGENTS.md')) { return $current }
    $parent = Split-Path -Parent $current
    if (-not $parent -or $parent -eq $current) { break }
    $current = $parent
  }
  throw "Unable to locate workspace root from $StartPath"
}

function Get-OutlookApp {
  try { return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { return New-Object -ComObject Outlook.Application }
}

function Get-PdfToTextPath {
  $cmd = Get-Command pdftotext -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }
  $here = (Get-Location).Path
  $common = @(
    (Join-Path $here '01-system\tools\runtimes\poppler\poppler-25.07.0\Library\bin\pdftotext.exe'),
    (Join-Path $here 'tools\poppler\Library\bin\pdftotext.exe'),
    (Join-Path $here 'tools\poppler\bin\pdftotext.exe'),
    "C:\\Program Files\\poppler\\bin\\pdftotext.exe",
    "C:\\Program Files (x86)\\poppler\\bin\\pdftotext.exe",
    "$env:ProgramFiles\\poppler\\bin\\pdftotext.exe",
    "$env:ProgramFiles(x86)\\poppler\\bin\\pdftotext.exe"
  )
  foreach($p in $common){ if(Test-Path -LiteralPath $p){ return $p } }
  try {
    $pop = Join-Path $here '01-system\tools\runtimes\poppler'
    if (-not (Test-Path -LiteralPath $pop)) {
      $pop = Join-Path $here 'tools\poppler'
    }
    if (Test-Path -LiteralPath $pop) {
      $hit = Get-ChildItem -LiteralPath $pop -Recurse -Filter 'pdftotext.exe' -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
      if ($hit) { return $hit }
    }
  } catch {}
  return $null
}

function Parse-AmountFromText {
  param([Parameter(Mandatory)][string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
  $currencyPattern = '(\$?\s*-?\d{1,3}(?:,\d{3})*(?:\.\d{2}))'
  function Is-DateLikeNumber {
    param([Parameter(Mandatory)][string]$Value)
    # Filter out dd.mm(.yyyy) or dd/mm(.yyyy) without a currency symbol
    if ($Value -match '^\$') { return $false }
    if ($Value -match '^(\d{1,2})[./](\d{1,2})([./]\d{2,4})?$') {
      $d = [int]$Matches[1]; $m = [int]$Matches[2]
      if ($d -ge 1 -and $d -le 31 -and $m -ge 1 -and $m -le 12) { return $true }
    }
    return $false
  }
  $patterns = @(
    "(?is)\b(?:grand\s+total|total\s+amount|amount\s+paid|total\s+paid|net\s+total|invoice\s+total)\b[\s\S]{0,200}?$currencyPattern",
    "(?is)\b(?:total(?:\s+amount)?|balance\s+due)\b[\s\S]{0,200}?$currencyPattern",
    "(?is)\b(?:AUD|NZD)\b[\s\S]{0,80}?$currencyPattern"
  )
  $candidates = New-Object System.Collections.Generic.List[object]
  foreach($p in $patterns){
    foreach($m in [regex]::Matches($Text,$p)){
      $raw = $m.Groups[1].Value
      if (-not $raw) { continue }
      $clean = ($raw -replace '\s','')
      if (-not $clean) { continue }
      if (Is-DateLikeNumber -Value $clean) { continue }
      $digits = ($clean -replace '[$,]','')
      try {
        $value = [decimal]::Parse($digits, [System.Globalization.CultureInfo]::InvariantCulture)
        $candidates.Add([pscustomobject]@{ Clean = $clean; Magnitude = [math]::Abs($value); Priority = 1 })
      } catch { }
    }
  }
  $moneyRegex = [regex]::new($currencyPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  foreach($match in $moneyRegex.Matches($Text)){
    $clean = ($match.Groups[1].Value -replace '\s','')
    if (-not $clean) { continue }
    if (Is-DateLikeNumber -Value $clean) { continue }
    $digits = ($clean -replace '[$,]','')
    try {
      $value = [decimal]::Parse($digits, [System.Globalization.CultureInfo]::InvariantCulture)
      $candidates.Add([pscustomobject]@{ Clean = $clean; Magnitude = [math]::Abs($value); Priority = 0 })
    } catch { }
  }
  if ($candidates.Count -gt 0) {
    $best = $candidates | Sort-Object -Property @{Expression = { $_.Magnitude }; Descending = $true }, @{Expression = { $_.Priority }; Descending = $true } | Select-Object -First 1
    return $best.Clean
  }
  return ''
}

function Parse-DocumentReference {
  param([Parameter(Mandatory)][string]$Text)
  $patterns = @(
    '(?is)document\s+ref[\s\S]{0,120}?no[:\s]*([A-Za-z0-9-]+)',
    '(?is)reference\s+number[:\s]*([A-Za-z0-9-]+)'
  )
  foreach($p in $patterns){
    $m = [regex]::Match($Text,$p)
    if ($m.Success) {
      $raw = $m.Groups[1].Value.Trim()
      if ($raw) { return ($raw -replace '\s','') }
    }
  }
  return ''
}

function Get-SafeName {
  param([string]$Value)
  if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
  return ($Value -replace '[\\/:*?"<>|]','_')
}

function Remember-ProcessedKey {
  param(
    [Parameter(Mandatory)][psobject]$Info,
    [Parameter(Mandatory)][string]$Key
  )
  if (-not $Info) { return }
  try {
    if ($Info.Set.Add($Key)) {
      Add-Content -Path $Info.File -Value $Key -Encoding UTF8
    }
  } catch {}
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
    return $dest
  } catch {
    Write-Warning ("Failed to save MSG: {0}" -f $_.Exception.Message)
    return $null
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
  $docRef = Parse-DocumentReference -Text $text
  $dir = Split-Path -Parent $Path
  $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
  $ext = [System.IO.Path]::GetExtension($Path)
  $hasAmtSuffix = $false
  $baseName = $name
  if ($name -match '^(.*) - \d[\d,]*\.?\d{0,2}$') {
    $hasAmtSuffix = $true
    $baseName = $Matches[1]
  }
  $safeDoc = $null
  if ($docRef) { $safeDoc = ($docRef -replace '[\\/:*?"<>|]','_') }
  if (-not $safeDoc -and $hasAmtSuffix) { return $Path }
  if ($safeDoc -and $hasAmtSuffix -and $baseName -eq $safeDoc) { return $Path }
  $prefix = if ($safeDoc) { $safeDoc } else { $baseName }
  $safePrefix = ($prefix -replace '[\\/:*?"<>|]','_')
  $safeAmt = ($amt -replace '[\\/:*?"<>|]','_')
  $target = Join-Path $dir ("{0} - {1}{2}" -f $safePrefix, $safeAmt, $ext)
  if (Test-Path -LiteralPath $target) {
    try {
      $existing = Get-Item -LiteralPath $target -ErrorAction Stop
      $current = Get-Item -LiteralPath $Path -ErrorAction Stop
      if ($existing.Length -eq $current.Length) {
        Remove-Item -LiteralPath $Path -Force
        Write-Host ("Duplicate already exists: {0}" -f $target)
        return $target
      }
    } catch {}
  }
  $n=1; $cand=$target; while(Test-Path -LiteralPath $cand){ $cand = Join-Path $dir ("{0} - {1} ({2}){3}" -f $safePrefix, $safeAmt, $n, $ext); $n++ }
  Move-Item -LiteralPath $Path -Destination $cand
  Write-Host ("Renamed with amount: {0}" -f $cand)
  return $cand
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$workspaceRoot = Get-WorkspaceRoot -StartPath $scriptRoot
$locationPushed = $false
$scriptFailed = $false
try {
  Push-Location -LiteralPath $workspaceRoot
  $locationPushed = $true
  $selectedDate = $null
  if ($PSBoundParameters.ContainsKey('Date') -and $Date) {
    $formats = @('yyyyMMdd','yyyy-MM-dd','yyyy/MM/dd')
    foreach($fmt in $formats){ try { $selectedDate = [datetime]::ParseExact($Date,$fmt,$null) ; break } catch {} }
    if (-not $selectedDate) { try { $selectedDate = [datetime]$Date } catch { $selectedDate = (Get-Date) } }
  } else { $selectedDate = (Get-Date) }
  $selectedDate = $selectedDate.Date
  $dateFolder = $selectedDate.ToString('yyyy-MM-dd')

  if (-not $PSBoundParameters.ContainsKey('SaveRoot') -or [string]::IsNullOrWhiteSpace($SaveRoot)) {
    $defaultRoot = Join-Path (Join-Path $workspaceRoot '03-outputs') 'remittance-runner'
    $SaveRoot = Join-Path (Join-Path $defaultRoot $dateFolder) 'files'
  }
  $SaveRoot = [IO.Path]::GetFullPath($SaveRoot)
  if (-not (Test-Path -LiteralPath $SaveRoot)) {
    New-Item -ItemType Directory -Path $SaveRoot -Force -ErrorAction SilentlyContinue | Out-Null
  }
  $processedDir = Join-Path (Split-Path $SaveRoot -Parent) 'processed'
  New-Item -ItemType Directory -Path $processedDir -Force -ErrorAction SilentlyContinue | Out-Null
  $processedMaps = @{}
  foreach($storeName in $stores){
    $safeStore = Get-SafeName -Value $storeName
    if (-not $safeStore) { $safeStore = 'store' }
    $procFile = Join-Path $processedDir ("processed-{0}.txt" -f $safeStore)
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    if (Test-Path -LiteralPath $procFile) {
      try {
        foreach($line in Get-Content -LiteralPath $procFile -ErrorAction SilentlyContinue){
          $trim = $line.Trim()
          if ($trim) { [void]$set.Add($trim) }
        }
      } catch {}
    }
    $processedMaps[$storeName] = [PSCustomObject]@{ File = $procFile; Set = $set }
  }

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

  $subjectRegex = [regex]::new('remittance|payment\s*advice|remittance\s*advice|payment\s*remittance|funds\s*transfer|eft\s*remittance','IgnoreCase')
  $fileNameRegex = [regex]::new('(remit|remittance|payment[\s_-]*advice|remit[\s_-]*advice|remittance[\s_-]*advice)','IgnoreCase')
  $allowedExtRegex = [regex]::new('\.(pdf)$','IgnoreCase')
  $imageExtRegex = [regex]::new('\.(png|jpg|jpeg|gif|bmp|svg|webp)$','IgnoreCase')
  $negativeSubjectRegex = [regex]::new('\b(statement|stmt|supplier\s*form|form|purchase\s*order|order)\b','IgnoreCase')
  $negativeFileNameRegex = [regex]::new('(statement|\bstmt\b|supplier[\s_-]*form|\bform\b|purchase[\s_-]*order|purchaseorder|\border\b|\bpo\d*\b)','IgnoreCase')
  $blockedAddrs = @('NZ-AR@NOVABIO.COM','AU-AR@NOVABIO.COM','au-orders@novabio.com','azhao@novabio.com') | ForEach-Object { $_.ToLower() }
  $allowAddrs = $AllowSenders | ForEach-Object { $_.ToLower() }
  $seen = New-Object 'System.Collections.Generic.HashSet[string]'
  $savedMsgPaths = New-Object 'System.Collections.Generic.List[string]'

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

  $storeDirs = @{}
  foreach($storeName in $stores){
    $store = ($ns.Folders | Where-Object { $_.Name -eq $storeName })
    if (-not $store) { Write-Host "Store not found: $storeName"; continue }
    $rootFolder = $store.Folders.Item('Inbox')
    if (-not $rootFolder) { continue }
    $saveDir = Join-Path $SaveRoot $storeName
    $storeDirs[$storeName] = $saveDir
    New-Item -ItemType Directory -Path $saveDir -Force | Out-Null
    $processedInfo = if ($processedMaps.ContainsKey($storeName)) { $processedMaps[$storeName] } else { $null }
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
        $entryId = ''
        try { $entryId = [string]$item.EntryID } catch {}
        if (-not $entryId) { $entryId = [guid]::NewGuid().ToString() }
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
          $msgPath = Save-MailAsMsg -Mail $item -TargetDir $saveDir
          if ($msgPath) { $savedMsgPaths.Add($msgPath) | Out-Null }
          continue
        }
          # If it has PDF attachments, fall through to normal PDF save flow
        }
        # Only proceed if subject mentions remittance or at least one attachment filename does
        $subjectMatch = $false; try { $subjectMatch = $subjectRegex.IsMatch($subj) } catch { $subjectMatch = $false }
        $subjectLower = ''; try { $subjectLower = $subj.ToLower() } catch { $subjectLower = '' }
        $subjectNegative = $false; if ($subjectLower) { $subjectNegative = ($subjectLower -like '*form*' -or $subjectLower -like '*statement*' -or $subjectLower -like '*stmt*' -or $subjectLower -like '*purchase*order*' -or $subjectLower -like '*order*') }
        $atts = $item.Attachments
        if (-not $atts -or $atts.Count -le 0) { continue }
        $nameMatch = $false
        try {
          for($ti=1; $ti -le $atts.Count; $ti++){ $tfn = [string]$atts.Item($ti).FileName; if ($fileNameRegex.IsMatch($tfn)) { $nameMatch = $true; break } }
        } catch { $nameMatch = $false }
        if (-not $Broad -and -not $allowOverride) {
          if (-not ($subjectMatch -or $nameMatch)) { continue }
        }
        if ($Broad -and -not $allowOverride) {
          if ($subjectNegative -and -not $subjectMatch) { continue }
        }
        $storeDir = $saveDir
        for($i=1; $i -le $atts.Count; $i++){
          $att = $atts.Item($i); if(-not $att){ continue }
          $fn = [string]$att.FileName
          $procSet = if ($processedInfo) { $processedInfo.Set } else { $null }
          $entryKey = ("{0}|{1}" -f $entryId,$i)
          if ($procSet -and $procSet.Contains($entryKey)) { Write-Host "Skip duplicate attachment $entryKey for $storeName"; continue }
          $isInline = $false
          try { $cid = $att.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F'); if($cid){$isInline=$true} } catch {}
          if ($isInline -and $imageExtRegex.IsMatch($fn)) { continue }
          $fnLower = $fn.ToLower()
          if ($fnLower -like '*form*' -or $fnLower -like '*supplier*form*' -or $fnLower -like '*statement*' -or $fnLower -like '*stmt*' -or $fnLower -like '*purchase*order*' -or $fnLower -like '*purchaseorder*' -or $fnLower -like '* order *') { continue }
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
            if (Test-Path -LiteralPath $cand) {
              $n = 1
              $stem = [IO.Path]::GetFileNameWithoutExtension($cand)
              $extOnly = [IO.Path]::GetExtension($cand)
              while (Test-Path -LiteralPath $cand) {
                $cand = Join-Path $storeDir ("{0} ({1}){2}" -f $stem, $n, $extOnly)
                $n++
              }
            }
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
                Remember-ProcessedKey -Info $processedInfo -Key $entryKey
              } else {
                Write-Host "Saved: $cand"
                Remember-ProcessedKey -Info $processedInfo -Key $entryKey
              }
            } else {
              Write-Host "Saved: $cand"
              Remember-ProcessedKey -Info $processedInfo -Key $entryKey
            }
          } catch { Write-Warning "Failed to save attachment '$fn': $($_.Exception.Message)" }
        }
      }
    }
  }
  if ($savedMsgPaths.Count -gt 0) {
    try {
      $pythonCmd = Get-Command python -ErrorAction Stop
      $converter = Join-Path $scriptRoot 'convert_msg_to_pdf.py'
      if (Test-Path -LiteralPath $converter) {
        Write-Host "Converting remittance MSGs to PDF..."
        $args = @($converter, '--msgs') + $savedMsgPaths
        & $pythonCmd.Source @args
      } else {
        Write-Warning "convert_msg_to_pdf.py not found; skipping MSG-to-PDF conversion."
      }
    } catch {
      Write-Warning ("MSG-to-PDF conversion failed: {0}" -f $_.Exception.Message)
    }
  }
  if ($PSBoundParameters.ContainsKey('PruneOriginals') -and $PruneOriginals) {
    foreach($storeName in $stores){
      if ($storeDirs.ContainsKey($storeName)) {
        $storeDir = $storeDirs[$storeName]
        if (Test-Path -LiteralPath $storeDir) { Remove-OriginalsWithoutAmountSuffix -Dir $storeDir }
      }
    }
  }
  $msgFiles = @(Get-ChildItem -LiteralPath $SaveRoot -Recurse -Filter *.msg -ErrorAction SilentlyContinue)
  if ($msgFiles -and $msgFiles.Count -gt 0) {
    try {
      $pythonCmd = Get-Command python -ErrorAction Stop
      $pyScript = Join-Path $scriptRoot 'download_yourremittance.py'
      if (Test-Path -LiteralPath $pyScript) {
        $arguments = @($pyScript, '--date', $dateFolder, '--base-dir', $SaveRoot, '--stores')
        $arguments += $stores
        Write-Host "Triggering secure remittance fetcher..."
        & $pythonCmd.Source @arguments
      } else {
        Write-Warning "download_yourremittance.py not found; skipping secure fetch."
      }
    } catch {
      Write-Warning ("Secure remittance fetcher failed: {0}" -f $_.Exception.Message)
    }
  }
  Write-Host ("Saved files under: {0}" -f $SaveRoot)
} catch {
  $scriptFailed = $true
  Write-Error $_
} finally {
  if ($locationPushed) {
    try { Pop-Location | Out-Null } catch {}
  }
}
if ($scriptFailed) { exit 1 }
