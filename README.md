# Gmail Confirmation Automation

Free local automation for generating balance confirmation documents, sending one Gmail email per party, and tracking sent/reply/bounce status in Excel.

## Quick Start

Validate and generate DOCX files without sending:

```powershell
python -m gmail_automation --master "Excel .xlsx" --mode preview --no-pdf
```

Generate PDFs as well, using Microsoft Word automation:

```powershell
python -m gmail_automation --master "Excel .xlsx" --mode preview
```

Create a sample Gmail config:

```powershell
python -m gmail_automation --init-config
```

Send emails only after `config.json` contains the sender Gmail address and app password:

```powershell
python -m gmail_automation --master "Excel .xlsx" --mode send
```

Track replies and bounces:

```powershell
python -m gmail_automation --master "Excel .xlsx" --mode track
```

## Workbooks

- `Excel .xlsx`: master user input workbook.
- `Tracking.xlsx`: execution state maintained by the app.

Generated files are written under `generated/<PartyId>/`. Logs are written under `logs/`.

## Build EXE

```powershell
.\scripts\build_exe.ps1
```
