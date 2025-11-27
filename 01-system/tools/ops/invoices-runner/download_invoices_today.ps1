#Requires -Version 5.1
param(
    [string]$SaveRoot,
    [string]$Mailbox = "",
    [string]$FolderPath = "Inbox",
    [string]$SubjectKeyword,
    [switch]$Recurse,
    [switch]$AllMailboxes,
    [int]$LookbackDays = 0,
    [switch]$GroupBySenderDomain,
    [ValidateSet('Invoices','Remittance')]
    [string]$Profile = 'Invoices',
    [Nullable[datetime]]$Date,
    [string]$PdfToTextPath,
    [switch]$RenameWithAmount,
    [string]$TimeZoneId = 'AUS Eastern Standard Time'
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

function Get-OutlookApp {
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    } catch {
        try {
            return New-Object -ComObject Outlook.Application
        } catch {
            throw "Unable to start or attach to Outlook. If you are using the new 'Outlook (new)' app, COM automation is not supported. Please use classic Outlook for Windows."
        }
    }
}

function Get-OutlookFolder {
    param(
        [Parameter(Mandatory)]$Namespace,
        [string]$Mailbox,
        [string]$FolderPath
    )
    if ([string]::IsNullOrWhiteSpace($Mailbox)) {
        $root = $Namespace.GetDefaultFolder(6) # olFolderInbox
    } else {
        $mailRoot = $Namespace.Folders | Where-Object { $_.Name -eq $Mailbox }
        if (-not $mailRoot) { throw "Mailbox '$Mailbox' not found." }
        $root = $mailRoot
        if ($FolderPath -eq 'Inbox') {
            $root = $mailRoot.Folders.Item('Inbox')
        }
    }

    if ([string]::IsNullOrWhiteSpace($FolderPath) -or $FolderPath -eq 'Inbox') {
        return $root
    }

    $current = if ($FolderPath -eq 'Inbox') { $root } else { $root }
    if ($FolderPath -ne 'Inbox') {
        $parts = $FolderPath -split '[\\/]' | Where-Object { $_ -and $_ -ne 'Inbox' }
        foreach ($p in $parts) {
            $current = $current.Folders.Item($p)
            if (-not $current) { throw "Folder path '$FolderPath' not found." }
        }
    }
    return $current
}

function Get-SafeSegment {
    param([Parameter(Mandatory)][string]$Text, [int]$MaxLength = 80)
    $t = ($Text).Trim() -replace '[\\/:*?"<>|]','_'
    if ($t.Length -gt $MaxLength) { $t = $t.Substring(0,$MaxLength) }
    if ([string]::IsNullOrWhiteSpace($t)) { return '_' } else { return $t }
}

function Get-SenderSmtp {
    param([Parameter(Mandatory)]$Mail)
    try {
        if ($Mail.SenderEmailType -eq 'EX') {
            $ex = $Mail.Sender.GetExchangeUser()
            if ($ex -and $ex.PrimarySmtpAddress) { return $ex.PrimarySmtpAddress }
        }
    } catch {}
    try {
        $v = $Mail.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001E')
        if ($v) { return [string]$v }
    } catch {}
    try { return [string]$Mail.SenderEmailAddress } catch { return '' }
}

function Get-PdfToTextPath {
    param([string]$Provided)
    if ($Provided) { return $Provided }
    $candidates = @('pdftotext.exe','pdftotext')
    foreach ($c in $candidates) {
        $cmd = (Get-Command $c -ErrorAction SilentlyContinue)
        if ($cmd) { return $cmd.Source }
    }
    $here = (Get-Location).Path
    $common = @(
        (Join-Path $here 'tools\poppler\Library\bin\pdftotext.exe'),
        (Join-Path $here 'tools\poppler\bin\pdftotext.exe'),
        "C:\\Program Files\\poppler\\bin\\pdftotext.exe",
        "C:\\Program Files (x86)\\poppler\\bin\\pdftotext.exe",
        "$env:ProgramFiles\\poppler\\bin\\pdftotext.exe",
        "$env:ProgramFiles(x86)\\poppler\\bin\\pdftotext.exe"
    )
    foreach ($p in $common) { if (Test-Path -LiteralPath $p) { return $p } }
    $pop = Join-Path $here 'tools\poppler'
    try {
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
    foreach ($p in $patterns) {
        $m = [regex]::Match($Text, $p)
        if ($m.Success) {
            $raw = $m.Groups[1].Value.Trim()
            $raw = $raw -replace '\s',''
            return $raw
        }
    }
    return ''
}

function Extract-AmountFromText {
    param([Parameter(Mandatory)][string]$Text)
    $patterns = @(
        '(?im)\b(?:grand\s+total|total\s+amount|amount\s+paid|total\s+paid|net\s+total|invoice\s+total)\D*([\$£€]?\s?-?\d{1,3}(?:,\d{3})*(?:\.\d{2})?)',
        '(?im)\b(?:AUD|NZD)\s*([\$£€]?\s?-?\d{1,3}(?:,\d{3})*(?:\.\d{2})?)',
        '(?im)\btotal\b\D*([\$£€]?\s?-?\d{1,3}(?:,\d{3})*(?:\.\d{2})?)'
    )
    foreach ($p in $patterns) {
        $m = [regex]::Match($Text, $p)
        if ($m.Success) {
            $raw = $m.Groups[1].Value.Trim()
            $raw = $raw -replace '\s',''
            return $raw
        }
    }
    return ''
}

function Try-RenameWithAmount {
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$PdfToTextPath
    )
    if (-not (Test-Path -LiteralPath $Path)) { return $Path }
    if ([System.IO.Path]::GetExtension($Path) -notmatch '^\.pdf$' -and [System.IO.Path]::GetExtension($Path) -notmatch '^\.PDF$') { return $Path }
    $tool = Get-PdfToTextPath -Provided $PdfToTextPath
    try {
        $text = ''
        if ($tool) {
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $tool
            $psi.Arguments = ('-layout -nopgbrk -q -f 1 -l 6 -enc UTF-8 "{0}" -' -f $Path)
            $psi.UseShellExecute = $false
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $p = [System.Diagnostics.Process]::Start($psi)
            $text = $p.StandardOutput.ReadToEnd()
            $err  = $p.StandardError.ReadToEnd()
            $p.WaitForExit()
            if ($p.ExitCode -ne 0) { Write-Warning ("pdftotext failed ({0}): {1}" -f $p.ExitCode,$err); $text = '' }
        }
        if (-not $text) {
            try {
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
            } catch {
                try { if ($doc) { $doc.Close($false) } } catch {}
                try { if ($word) { $word.Quit() } } catch {}
                Write-Warning "Word-based PDF text extraction failed for '$Path': $($_.Exception.Message)"
            }
        }
        $amt = if ($text) { Parse-AmountFromText -Text $text } else { '' }
        if (-not $amt) { return $Path }
        $dir = Split-Path -Parent $Path
        $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $ext = [System.IO.Path]::GetExtension($Path)
        $safeAmt = ($amt -replace '[\\/:*?"<>|]','_')
        $newName = ("{0} - {1}{2}" -f $name,$safeAmt,$ext)
        $newPath = Join-Path $dir $newName
        $newPath = Get-UniquePath -Path $newPath
        Move-Item -LiteralPath $Path -Destination $newPath
        Write-Host ("Renamed with amount: {0}" -f $newPath)
        return $newPath
    } catch {
        Write-Warning "Failed to extract amount for '$Path': $($_.Exception.Message)"
        return $Path
    }
}

function Get-TargetFolders {
    param(
        [Parameter(Mandatory)]$Namespace,
        [string]$Mailbox,
        [string]$FolderPath,
        [switch]$AllMailboxes,
        [switch]$Recurse,
        [string[]]$IncludeStoreNames
    )

    $targets = New-Object System.Collections.ArrayList

    function AddFolderAndChildren {
        param($folder)
        if (-not $folder) { return }
        [void]$targets.Add($folder)
        if ($Recurse) {
            foreach ($sub in $folder.Folders) {
                AddFolderAndChildren -folder $sub
            }
        }
    }

    if ($AllMailboxes) {
        foreach ($store in $Namespace.Folders) {
            $storeDisplay = try { [string]$store.DisplayName } catch { [string]$store.Name }
            if ($IncludeStoreNames -and @($IncludeStoreNames | Where-Object { $_ -ieq $storeDisplay }).Count -eq 0) { continue }
            try {
                $target = $null
                if ([string]::IsNullOrWhiteSpace($FolderPath) -or $FolderPath -eq 'Inbox') {
                    $target = $store.Folders.Item('Inbox')
                } else {
                    $target = $store.Folders.Item('Inbox')
                    if ($target -and $FolderPath -match 'Inbox[\\/](.+)') {
                        $parts = $Matches[1] -split '[\\/]' | Where-Object { $_ }
                        foreach ($p in $parts) {
                            $target = $target.Folders.Item($p)
                            if (-not $target) { break }
                        }
                    }
                }
                if ($target) { AddFolderAndChildren -folder $target }
            } catch { }
        }
    } else {
        $folder = Get-OutlookFolder -Namespace $Namespace -Mailbox $Mailbox -FolderPath $FolderPath
        AddFolderAndChildren -folder $folder
    }

    return ,$targets
}

try {
    if (-not $SaveRoot) {
        $SaveRoot = Join-Path -Path (Get-Location) -ChildPath "Inv&Remit_Today"
    }
    $dateFolder = (Get-Date).ToString('yyyy-MM-dd')
    $saveDir = Join-Path -Path $SaveRoot -ChildPath $dateFolder
    New-Item -ItemType Directory -Path $saveDir -Force | Out-Null

    $outlook = Get-OutlookApp
    $ns = $outlook.GetNamespace('MAPI')
    # Choose stores based on profile
    $includeStores = @()
    if ($Profile -eq 'Invoices') {
        $includeStores = @('AZhao@novabio.com')
    } elseif ($Profile -eq 'Remittance') {
        $includeStores = @('Australia AR','New Zealand AR')
    }

    $scanAll = $true
    if (-not $includeStores -or $includeStores.Count -eq 0) { $scanAll = $AllMailboxes } else { $scanAll = $true }

    $targetFolders = Get-TargetFolders -Namespace $ns -Mailbox $Mailbox -FolderPath $FolderPath -AllMailboxes:$scanAll -Recurse:$Recurse -IncludeStoreNames $includeStores

    # Build time window in specified timezone, then convert to local for MAPI Restrict
    try { $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId) } catch { $tz = [System.TimeZoneInfo]::Local }
    if ($PSBoundParameters.ContainsKey('Date') -and $Date) {
        $startTZ = $Date.Date
        $endTZ   = $startTZ.AddDays(1)
    } else {
        $nowTZ = [System.TimeZoneInfo]::ConvertTime([datetime]::UtcNow, [System.TimeZoneInfo]::Utc, $tz)
        $startTZ = $nowTZ.Date.AddDays(-1 * [Math]::Max(0,$LookbackDays))
        $endTZ   = $nowTZ.Date.AddDays(1)
    }
    $start = [System.TimeZoneInfo]::ConvertTime($startTZ, $tz, [System.TimeZoneInfo]::Local)
    $end   = [System.TimeZoneInfo]::ConvertTime($endTZ,   $tz, [System.TimeZoneInfo]::Local)
    $fStart = $start.ToString('MM/dd/yyyy hh:mm tt')
    $fEnd   = $end.ToString('MM/dd/yyyy hh:mm tt')
    $restriction = "[ReceivedTime] >= '$fStart' AND [ReceivedTime] < '$fEnd'"

    # Profile-driven defaults
    if (-not $PSBoundParameters.ContainsKey('SubjectKeyword')) {
        if ($Profile -eq 'Invoices') { $SubjectKeyword = 'invoice' }
        elseif ($Profile -eq 'Remittance') { $SubjectKeyword = 'remittance' }
    }

    $subjectRegex = if ([string]::IsNullOrWhiteSpace($SubjectKeyword)) { $null } else { [regex]::new($SubjectKeyword, 'IgnoreCase') }
    $allowedExtRegex = [regex]::new('\.(pdf)$','IgnoreCase')
    $imageExtRegex = [regex]::new('\.(png|jpg|jpeg|gif|bmp|svg|webp)$','IgnoreCase')

    $msgCount = 0
    $msgMatched = 0
    $attSaved = 0

    $indexPath = Join-Path $saveDir '_index.csv'
    if (-not (Test-Path -LiteralPath $indexPath)) {
        'SavedPath,Store,Folder,SenderName,SenderAddress,Subject,ReceivedTime,AttachmentName,AttachmentIndex,EntryID' | Out-File -FilePath $indexPath -Encoding UTF8
    }

    foreach ($folder in $targetFolders) {
        try {
            $storeName = try { [string]$folder.Store.DisplayName } catch { '' }
            if (-not $storeName) {
                try {
                    $fp = [string]$folder.FolderPath
                    if ($fp -match '^[\\]{2}([^\\]+)') { $storeName = $Matches[1] } else { $storeName = [string]$folder.Name }
                } catch { $storeName = [string]$folder.Name }
            }
            $storeSeg = Get-SafeSegment -Text $storeName

            $folderPathRaw = try { [string]$folder.FolderPath } catch { [string]$folder.Name }
            $relFolder = $folderPathRaw
            if ($folderPathRaw -match '^[\\]{2}[^\\]+\\(.+)$') { $relFolder = $Matches[1] }
            $relFolder = [uri]::UnescapeDataString($relFolder)

            # Simplified structure: date\store only
            $baseDir = Join-Path $saveDir $storeSeg
            New-Item -ItemType Directory -Path $baseDir -Force | Out-Null

            $items = $folder.Items
            $items.IncludeRecurrences = $true
            $items.Sort('[ReceivedTime]')
            $itemsToday = $items.Restrict($restriction)

            $folderProcessed = 0
            $folderSaved = 0

            foreach ($item in $itemsToday) {
                $folderProcessed++
                $msgCount++
                # olObjectClass 43 = MailItem
                $isMail = $false
                try { if ($item -and $item.Class -eq 43) { $isMail = $true } } catch { $isMail = $false }
                if (-not $isMail) { continue }

                $subject = ''
                try { $subject = [string]$item.Subject } catch { $subject = '' }
                $subjectLooksInvoice = $false
                if ($subjectRegex) { $subjectLooksInvoice = $subjectRegex.IsMatch($subject) }

                $hasAttachments = $false
                try { if ($item.Attachments -and $item.Attachments.Count -gt 0) { $hasAttachments = $true } } catch { $hasAttachments = $false }
                if (-not $hasAttachments) { continue }

                $senderName = try { [string]$item.SenderName } catch { '' }
                $senderAddress = Get-SenderSmtp -Mail $item
                $senderDomain = ''
                if ($senderAddress -and $senderAddress -match '@(.+)$') { $senderDomain = $Matches[1] }
                # Blocklist for Remittance profile
                $blockedSenders = @('NZ-AR@NOVABIO.COM','AU-AR@NOVABIO.COM','au-orders@novabio.com','azhao@novabio.com')
                $blockedLower = $blockedSenders | ForEach-Object { $_.ToLower() }
                if ($Profile -eq 'Remittance') {
                    $senderLower = ([string]$senderAddress).ToLower()
                    if ($blockedLower -contains $senderLower) { continue }
                    # Block entire novabio.com domain for remittance
                    try {
                        $domainLower = ($senderLower -split '@')[-1]
                        if ($domainLower -eq 'novabio.com') { continue }
                    } catch { }
                }

                # Save directly under date\store
                $mailSaveDir = $baseDir
                if (-not (Test-Path -LiteralPath $mailSaveDir)) { New-Item -ItemType Directory -Path $mailSaveDir -Force | Out-Null }

                $savedOneFromThisMail = $false
                for ($i = 1; $i -le $item.Attachments.Count; $i++) {
                    $att = $item.Attachments.Item($i)
                    if (-not $att) { continue }
                    $fileName = [string]$att.FileName

                    # Try to detect inline (embedded) attachments by content-id; if image and inline, skip
                    $isInline = $false
                    try {
                        $cid = $att.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F')
                        if ($cid) { $isInline = $true }
                    } catch { }

                    $nameMatchesKeyword = $false
                    if ($subjectRegex) { $nameMatchesKeyword = $subjectRegex.IsMatch($fileName) }
                    $extLooksDoc = $allowedExtRegex.IsMatch($fileName)
                    $isInlineImage = $isInline -and $imageExtRegex.IsMatch($fileName)

                    if (($subjectLooksInvoice -or $nameMatchesKeyword -or $extLooksDoc) -and -not $isInlineImage) {
                        $safeName = Get-SafeSegment -Text $fileName -MaxLength 120
                        $targetPath = Join-Path $mailSaveDir $safeName
                        $targetPath = Get-UniquePath -Path $targetPath
                        try {
                            $originalPath = $targetPath
                            $att.SaveAsFile($targetPath)
                            # Optional rename with amount for remittance PDFs
                            $doRename = $false
                            if ($Profile -eq 'Remittance') {
                                if ($PSBoundParameters.ContainsKey('RenameWithAmount')) { $doRename = [bool]$RenameWithAmount } else { $doRename = $true }
                                if ($doRename) { $targetPath = Try-RenameWithAmount -Path $targetPath -PdfToTextPath $PdfToTextPath }
                            }
                            if ($targetPath -ne $originalPath -and (Test-Path -LiteralPath $originalPath)) { Remove-Item -LiteralPath $originalPath -Force }
                            $attSaved++
                            $folderSaved++
                            $savedOneFromThisMail = $true
                            Write-Host "Saved: $targetPath"
                            # Append index row
                            $receivedStr = ''
                            try { $receivedStr = ($item.ReceivedTime).ToString('s') } catch {}
                            $entryId = ''
                            try { $entryId = [string]$item.EntryID } catch {}
                            $row = '{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}' -f ($targetPath -replace ',',' '), ($storeName -replace ',',' '), ($relFolder -replace ',',' '), ($senderName -replace ',',' '), ($senderAddress -replace ',',' '), ($subject -replace ',',' '), $receivedStr, ($fileName -replace ',',' '), $i, $entryId
                            Add-Content -LiteralPath $indexPath -Value $row -Encoding UTF8
                        } catch {
                            Write-Warning "Failed to save attachment '$fileName': $($_.Exception.Message)"
                        }
                    }
                }
                if ($savedOneFromThisMail) { $msgMatched++ }
            }

            $fPath = try { $folder.FolderPath } catch { $folder.Name }
            Write-Host ("Folder processed: {0} | Items today: {1} | Attachments saved: {2}" -f $fPath, $folderProcessed, $folderSaved)
        } catch {
            Write-Warning "Skipping folder due to error: $($_.Exception.Message)"
        }
    }

    Write-Host "Processed messages today: $msgCount"
    Write-Host "Messages matched:        $msgMatched"
    Write-Host "Attachments saved:       $attSaved"
    Write-Host "Save folder:            $saveDir"

} catch {
    Write-Error $_
    exit 1
} finally {
    try { [GC]::Collect(); [GC]::WaitForPendingFinalizers() } catch { }
}
