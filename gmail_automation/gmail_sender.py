from __future__ import annotations

import email.utils
import mimetypes
import smtplib
import uuid
from email.message import EmailMessage
from pathlib import Path

from .models import ConfirmationRow, DocumentResult, MailTemplate, SendConfig, SendResult, now_text


class GmailSender:
    SMTP_HOST = "smtp.gmail.com"
    SMTP_PORT = 587

    def __init__(self, config: SendConfig) -> None:
        self.config = config

    def send(self, row: ConfirmationRow, mail_template: MailTemplate, documents: DocumentResult) -> SendResult:
        if not self.config.sender_email or not self.config.app_password:
            raise ValueError("Gmail sender email and app password are required for sending")

        message_id = email.utils.make_msgid(idstring=f"{row.row_id}.{uuid.uuid4().hex}", domain="gmail-automation.local")
        message = EmailMessage()
        message["From"] = self.config.sender_email
        message["To"] = row.email
        if row.cc:
            message["Cc"] = row.cc
        message["Subject"] = row.subject
        message["Message-ID"] = message_id
        message.set_content(self._render_body(row, mail_template))

        attachments = (documents.pdf_paths or documents.docx_paths) + documents.extra_attachment_paths
        for path in attachments:
            self._attach(message, path)

        recipients = [row.email] + [part.strip() for part in row.cc.split(";") if part.strip()]
        with smtplib.SMTP(self.SMTP_HOST, self.SMTP_PORT, timeout=60) as smtp:
            smtp.starttls()
            smtp.login(self.config.sender_email, self.config.app_password)
            smtp.send_message(message, from_addr=self.config.sender_email, to_addrs=recipients)

        return SendResult(
            sent_at=now_text(),
            smtp_message_id=message_id,
            gmail_thread_id="",
            to_email=row.email,
            cc_email=row.cc,
            attachment_paths=attachments,
        )

    def _render_body(self, row: ConfirmationRow, mail_template: MailTemplate) -> str:
        body = row.mail_body_override or mail_template.body_template_text
        replacements = {
            "{party_name}": row.party_name,
            "{amount}": row.balance,
            "{{First Name}}": row.contact_first_name,
            "{{Last Name}}": row.contact_last_name,
            "March 31, 2025": row.balance_as_on_date,
            "31st March 2025": row.balance_as_on_date,
        }
        for old, new in replacements.items():
            body = body.replace(old, new)
        return body

    def _attach(self, message: EmailMessage, path: Path) -> None:
        if not path.exists():
            raise FileNotFoundError(f"Attachment not found: {path}")
        mime_type, _ = mimetypes.guess_type(path.name)
        if mime_type:
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"
        message.add_attachment(path.read_bytes(), maintype=maintype, subtype=subtype, filename=path.name)
