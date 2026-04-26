# Gmail Confirmation Automation

Local Windows automation for balance confirmation work. It reads parties from a master Excel workbook, generates Word/PDF confirmation attachments from embedded templates, sends one Gmail email per party, and maintains a resumable tracking workbook.

## What It Does

- Generates authorisation letter, balance confirmation letter, and reply form for each party.
- Converts generated DOCX files to PDF using Microsoft Word automation.
- Sends Gmail messages through SMTP using a Gmail App Password.
- Tracks sent status, replies, and bounces through Gmail IMAP.
- Writes progress to `Tracking.xlsx` so reruns skip completed work and continue from the last successful step.

## Files

- `Excel .xlsx`: main user input workbook.
- `config.json`: Gmail and run settings.
- `Tracking.xlsx`: generated/updated automatically for execution state.
- `generated/<PartyId>/`: generated DOCX/PDF files for each party.
- `logs/`: run logs with validation, generation, send, and tracking results.
- `dist/GmailConfirmationAutomation.exe`: packaged executable.

The Word templates are embedded in the application resources, so normal users only need the master Excel file and `config.json`.

## Quick Start With EXE

Run full automation:

```powershell
.\dist\GmailConfirmationAutomation.exe --master "Excel .xlsx" --mode all
```

Validate and generate attachments without sending:

```powershell
.\dist\GmailConfirmationAutomation.exe --master "Excel .xlsx" --mode preview
```

Generate DOCX only, without PDF conversion:

```powershell
.\dist\GmailConfirmationAutomation.exe --master "Excel .xlsx" --mode preview --no-pdf
```

Track replies and bounces after emails are sent:

```powershell
.\dist\GmailConfirmationAutomation.exe --master "Excel .xlsx" --mode track
```

## Run Modes

- `validate`: checks workbook and embedded templates.
- `preview`: validates, generates documents, and verifies output; does not send email.
- `generate`: same document generation flow for preparing attachments.
- `send`: generates if needed, then sends eligible rows.
- `track`: checks Gmail inbox for replies and bounce messages.
- `all`: runs generation, sending, and tracking in one command.

## Config

Create a sample config:

```powershell
.\dist\GmailConfirmationAutomation.exe --init-config
```

Important fields in `config.json`:

- `sender_email`: Gmail address used as the sender.
- `app_password`: Gmail App Password, not the normal Gmail login password.
- `send_mode`: use `send` for actual sending; use `preview` for safe testing.
- `convert_to_pdf`: `true` to create PDFs using Microsoft Word.
- `batch_size`, `batch_delay_seconds`, `per_email_delay_seconds`, `daily_send_limit`: sending limits and delays.

Do not share or commit a real Gmail App Password. If it is exposed, revoke it in Google Account settings and create a new one.

## Master Excel Columns

The workbook should contain a `Confirmations` sheet. Key columns:

- `PartyId`: unique ID for each party, also used for generated folder names.
- `Name`: party/contact name.
- `To Email`: primary recipient email address.
- `Email`: fallback recipient if `To Email` is blank.
- `CC`: optional CC addresses separated by semicolon.
- `Subject`: email subject.
- `Balance`: balance amount.
- `BalanceNature`: receivable/payable or similar balance description.
- `CompanyName`: company name used in the letters.
- `Address`: party address.
- `Phone`: optional phone number.
- `BalanceAsOnDate`: date to insert in letters.
- `LetterDate`: letter date.
- `AuditorReplyEmail`: reply email shown in templates.
- `File Path Locations`: optional extra attachments, separated by semicolon or new line.
- `MailBodyOverride`: optional custom email body for that row.

## Tracking Workbook

`Tracking.xlsx` is maintained by the app. Important status fields:

- `AttachmentCreated`: generated attachment files were created.
- `GeneratedDocxPaths` and `GeneratedPdfPaths`: file paths written for the party.
- `ReadyToSend`: generated files passed verification and can be sent.
- `MainSent`: email was sent successfully.
- `SentDate`: date/time of send.
- `GmailMessageId`: generated message ID used for audit.
- `BounceReceived`: Gmail bounce message was detected.
- `ReplyReceived`: an actual reply email was detected; this is not read/open tracking.
- `ReplyReceivedDate`: date/time of detected reply.
- `LastCheckedAt`: last Gmail tracking check time.
- `Status` and `Error`: current row state and latest issue, if any.

## Successful Run Checklist

After a run, check:

- Latest file in `logs/` ends with `Run completed` and has no `[ERROR]` lines.
- Each party has files under `generated/<PartyId>/`.
- PDF count matches DOCX count when `convert_to_pdf` is enabled.
- Opening generated DOCX files in Word does not show a repair/unreadable-content prompt.
- `Tracking.xlsx` shows `AttachmentCreated = Y`, `ReadyToSend = Y`, and `VerificationStatus = Passed`.
- For sent rows, `MainSent = Y`, `SentDate` is filled, and `GmailMessageId` is filled.
- Gmail Sent Mail contains the sent messages with the expected recipients and attachments.
- After replies or bounces arrive, running `--mode track` updates `ReplyReceived`, `ReplyReceivedDate`, `BounceReceived`, and `BounceDate`.

## Python Development Commands

Run from source:

```powershell
python -m gmail_automation --master "Excel .xlsx" --mode preview
```

Build the executable:

```powershell
.\scripts\build_exe.ps1
```
