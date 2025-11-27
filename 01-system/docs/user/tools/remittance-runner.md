# remittance-runner
**Category**: ops
**Version**: v0.3 (Updated: 2025-11-27)

## What it does
- Collects remittance advices from Outlook (AU/NZ stores) and saves PDFs.
- Scans Inbox (and subfolders when -Recurse) for matching subjects/attachments.
- Renames PDFs with detected amounts; auto-converts remittance .msg to HTML/PDF; triggers secure fetcher for portal links.
- Keeps intermediate HTML/non-amount PDFs/MSG sources under `03-outputs/remittance-runner/<date>/intermediate/msg-{html,pdf,src}/`; store folders keep final PDFs.
- Writes outputs under `03-outputs/remittance-runner/` by date.

## Parameters
- `Stores` (array): e.g., `Australia AR`,`New Zealand AR`.
- `Date` (string): `YYYY-MM-DD` or `YYYYMMDD`.
- `TimeZoneId` (string): default `AUS Eastern Standard Time`.
- `FastScan` (switch): scan last N items (set via `MaxItems`).
- `MaxItems` (int): default 400; use 800 for broad scans.
- `Recurse` (switch): include subfolders.
- `PruneOriginals` (switch): drop non-amount-suffixed duplicates.
- `Broad` (switch): looser subject/filename filters.
- `AllowSenders` (array): extra allowed sender addresses (defaults include payments@nzdf.mil.nz, payables@ap1.fpim.health.nz, and core AU senders).

## Usage
1) AU only fast scan:
```
powershell -NoProfile -File 01-system/tools/ops/remittance-runner/run.ps1 -Stores 'Australia AR' -Date 'YYYY-MM-DD' -FastScan -MaxItems 400 -PruneOriginals
```
2) AU+NZ broad/recurse:
```
powershell -NoProfile -File 01-system/tools/ops/remittance-runner/run.ps1 -Stores 'Australia AR','New Zealand AR' -Date 'YYYY-MM-DD' -Recurse -Broad -FastScan -MaxItems 800 -PruneOriginals
```
3) Add allowed senders:
```
powershell -NoProfile -File 01-system/tools/ops/remittance-runner/run.ps1 -Stores 'Australia AR','New Zealand AR' -Date 'YYYY-MM-DD' -AllowSenders 'payments@nzdf.mil.nz','payables@ap1.fpim.health.nz'
```

## Paths
- Input: Outlook Inbox folders for the specified stores.
- Output: `03-outputs/remittance-runner/<YYYY-MM-DD>/` (final PDFs), with intermediates in `intermediate/msg-{html,pdf,src}/`.

## Requirements
- Outlook with the store mailboxes, PowerShell 5.1+.
- Poppler/Acrobat/Word available for PDF text extraction (bundled Poppler auto-detected at `01-system/tools/runtimes/poppler/poppler-25.07.0/Library/bin/pdftotext.exe`).
-- For secure fetch: Playwright/Chromium (bundled) and mail access for OTP delivery.

## Tips
- If you re-run and need to regrab items, clear the corresponding processed file under `03-outputs/processed/`.

## Changelog
- v0.2 (2025-11-27): Documented intermediates, secure fetch, and bundled Poppler auto-detect.
- v0.1 (2025-11-13): Initial version.
