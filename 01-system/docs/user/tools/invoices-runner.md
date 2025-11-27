# invoices-runner
**Category**: ops
**Version**: v0.1 (Updated: 2025-11-13)

## What it does
- Pulls invoice PDFs from Outlook Inbox (AZhao@novabio.com) and saves to dated folders.
- Optional subfolder scan with `-Recurse`.
- Writes outputs under `03-outputs/invoices-runner/` by date; uses `supplier_map.json` to label senders.

## Parameters
- `Date` (string): `YYYY-MM-DD` or `YYYYMMDD`.
- `TimeZoneId` (string): default `AUS Eastern Standard Time`.
- `Recurse` (switch): include subfolders.

## Usage
1) Inbox only:
```
powershell -NoProfile -File 01-system/tools/ops/invoices-runner/run.ps1 -Date 'YYYY-MM-DD'
```
2) Include subfolders:
```
powershell -NoProfile -File 01-system/tools/ops/invoices-runner/run.ps1 -Date 'YYYY-MM-DD' -Recurse
```

## Paths
- Input: Outlook Inbox for AZhao@novabio.com (and subs if `-Recurse`).
- Output: `03-outputs/invoices-runner/<YYYY-MM-DD>/`.

## Requirements
- Outlook with mailbox access; PowerShell 5.1+.
- Supplier mapping file lives at `01-system/tools/ops/invoices-runner/supplier_map.json`.

## Changelog
- v0.1 (2025-11-13): Initial version.
