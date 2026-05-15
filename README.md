# Gmail Confirmation Automation

Local Windows automation for balance confirmation work. It reads parties from a multi-sheet master Excel workbook, generates Word/PDF confirmation attachments from bundled Word templates, sends one email per party, and maintains a resumable tracking workbook.

## What It Does

- Generates a balance confirmation letter and vendor reply form for each party.
- Attaches the static authorisation PDF to every email.
- Converts generated DOCX files to PDF using Microsoft Word automation.
- Sends mail using one of:
  - Webtel SMTP for `ngmks.in` style hosted mailboxes
  - Gmail SMTP
- Tracks sent status, replies, and bounces through Gmail IMAP.
- Writes progress to `Tracking.xlsx` so reruns skip completed work and continue from the last successful step.

## Files

- `Information for External Balance Confirmations (1).xlsx`: main user input workbook.
- `config.json`: optional compatibility settings for CLI mode.
- `Tracking.xlsx`: generated/updated automatically for execution state.
- `generated/<PartyId>/`: generated DOCX/PDF files for each party.
- `logs/`: run logs with validation, generation, send, and tracking results.
- `dist/GmailConfirmationAutomation.exe`: packaged executable.

The mail body is codified in Python. The generated document formatting comes from Word templates bundled inside the app, so normal users only need the master Excel file; the static authorisation PDF is also bundled with the app.

## Quick Start With EXE

Launch the desktop UI:

```powershell
.\dist\GmailConfirmationAutomation.exe --ui
```

Run full automation:

```powershell
.\dist\GmailConfirmationAutomation.exe --master "Information for External Balance Confirmations (1).xlsx" --mode all
```

Validate and generate attachments without sending:

```powershell
.\dist\GmailConfirmationAutomation.exe --master "Information for External Balance Confirmations (1).xlsx" --mode preview
```

Generate DOCX only, without PDF conversion:

```powershell
.\dist\GmailConfirmationAutomation.exe --master "Information for External Balance Confirmations (1).xlsx" --mode preview --no-pdf
```

Track replies and bounces after emails are sent:

```powershell
.\dist\GmailConfirmationAutomation.exe --master "Information for External Balance Confirmations (1).xlsx" --mode track
```

## Run Modes

- `validate`: checks workbook and embedded templates.
- `preview`: validates, generates documents, and verifies output; does not send email.
- `generate`: same document generation flow for preparing attachments.
- `send`: generates if needed, then sends eligible rows.
- `track`: checks Gmail inbox for replies and bounce messages.
- `all`: runs generation, sending, and tracking in one command.
- `--ui`: opens the desktop batch approval UI instead of running a command-line mode.

## Desktop Batch UI

The UI is intended for supervised sending:

- Choose the input Excel workbook with the file picker.
- Select a batch size and load the workbook.
- Generate and verify documents before sending.
- Click `Preview Next Batch` to see the next docs-ready customers with mail pending, then click `Run Preview Batch` to approve and send only that displayed batch.
- Use the `Results` tab to filter by batch, see per-customer document/batch/send status, and send or resend pending rows where documents are ready but mail is not sent.

Actual email sending still requires send mode `send` in Mail Settings. Keep send mode as `preview` when validating the workbook or generated documents without sending mail.

## Webtel SMTP

For Webtel-hosted domain mailboxes such as `ghanshyam@ngmks.in`, choose `Use Webtel SMTP` in Mail Setup and enter only:

- Sender email / username: full mailbox address.
- Webtel mailbox password.

The app uses the verified server settings internally:

- SMTP host: `connect.webtelconnect.com`
- Port: `465`
- Security: SSL/TLS

## Config

Create a sample config:

```powershell
.\dist\GmailConfirmationAutomation.exe --init-config
```

Important fields in `config.json` (CLI compatibility):

- `mail_provider`: `gmail_smtp` or `webtel_smtp`.
- `sender_email`: sender mailbox address.
- `app_password` / `smtp_password`: SMTP credential for SMTP providers.
- `send_mode`: use `send` for actual sending; use `preview` for safe testing.
- `convert_to_pdf`: `true` to create PDFs using Microsoft Word.
- `batch_size`, `batch_delay_seconds`, `per_email_delay_seconds`, `daily_send_limit`: sending limits and delays.

Do not share or commit a real Gmail App Password. If it is exposed, revoke it in Google Account settings and create a new one.

## Master Excel Columns

The workbook can contain multiple sheets. The app processes every non-blank sheet except `Banks`; blank sheets and rows are skipped.

Required columns on each processed sheet:

- `S.No.`: row serial number; combined with the sheet name to create the tracking ID.
- `Party Type`: category/type from the workbook.
- `Party Name`: vendor/customer name inserted into generated documents.
- `Email To(Address)`: recipient email address.
- `Balance`: amount inserted into the mail body and generated documents.

## Tracking Workbook

`Tracking.xlsx` is maintained by the app. Important status fields:

- `PartyId`: generated as `<SheetName>-<S.No.>` for resumable tracking.
- `SheetName`, `S.No.`, `Party Type`, `Party Name`, `Email To(Address)`, `Balance`: source workbook details.
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
- Each party has files under `generated/<PartyId>/`, and the static PDF is copied under `generated/_static/`.
- PDF count matches DOCX count when `convert_to_pdf` is enabled.
- Opening generated DOCX files in Word does not show a repair/unreadable-content prompt.
- `Tracking.xlsx` shows `AttachmentCreated = Y`, `ReadyToSend = Y`, and `VerificationStatus = Passed`.
- For sent rows, `MainSent = Y`, `SentDate` is filled, and `GmailMessageId` is filled.
- Gmail Sent Mail contains the sent messages with the expected recipients and attachments.
- After replies or bounces arrive, running `--mode track` updates `ReplyReceived`, `ReplyReceivedDate`, `BounceReceived`, and `BounceDate`.

## Python Development Commands

Run from source:

```powershell
python -m gmail_automation --master "Information for External Balance Confirmations (1).xlsx" --mode preview
```

Launch the desktop UI from source:

```powershell
python -m gmail_automation --ui
```

Launch the modern DB-backed compliance UI from source:

```powershell
python -m gmail_automation --modern-ui
```

Initialize the local SQLite compliance database without launching the UI:

```powershell
python -m gmail_automation --init-db
```

Build the executable:

```powershell
.\scripts\build_exe.ps1
```
