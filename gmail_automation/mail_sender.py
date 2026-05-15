from __future__ import annotations

import email.utils
import mimetypes
import smtplib
import uuid
from dataclasses import replace
from email.message import EmailMessage
from pathlib import Path

from .formatting import format_amount
from .models import ConfirmationRow, DocumentResult, MailTemplate, SendConfig, SendResult, now_text


SUPPORTED_MAIL_PROVIDERS = {"gmail_smtp", "webtel_smtp"}


class MailSender:
    def send(self, row: ConfirmationRow, mail_template: MailTemplate, documents: DocumentResult) -> SendResult:
        raise NotImplementedError

    def test_connection(self) -> str:
        raise NotImplementedError


class SmtpMailSender(MailSender):
    def __init__(self, config: SendConfig, host: str, port: int, starttls: bool = True, ssl_tls: bool = False) -> None:
        self.config = config
        self.host = host
        self.port = port
        self.starttls = starttls
        self.ssl_tls = ssl_tls

    def send(self, row: ConfirmationRow, mail_template: MailTemplate, documents: DocumentResult) -> SendResult:
        sender_email = self.config.sender_email.strip()
        password = self.config.smtp_password.strip() or self.config.app_password.strip()
        username = self.config.smtp_username.strip() or sender_email
        if not sender_email or not password:
            raise ValueError("Sender email and SMTP password are required for sending")

        message, message_id, attachments, recipients = build_email_message(row, mail_template, documents, sender_email)
        smtp_class = smtplib.SMTP_SSL if self.ssl_tls else smtplib.SMTP
        with smtp_class(self.host, self.port, timeout=60) as smtp:
            if self.starttls and not self.ssl_tls:
                smtp.starttls()
            smtp.login(username, password)
            smtp.send_message(message, from_addr=sender_email, to_addrs=recipients)

        return SendResult(
            sent_at=now_text(),
            smtp_message_id=message_id,
            gmail_thread_id="",
            to_email=row.email,
            cc_email=row.cc,
            attachment_paths=attachments,
        )

    def test_connection(self) -> str:
        sender_email = self.config.sender_email.strip()
        password = self.config.smtp_password.strip() or self.config.app_password.strip()
        username = self.config.smtp_username.strip() or sender_email
        if not sender_email or not password:
            raise ValueError("Sender email and SMTP password are required")
        smtp_class = smtplib.SMTP_SSL if self.ssl_tls else smtplib.SMTP
        with smtp_class(self.host, self.port, timeout=30) as smtp:
            if self.starttls and not self.ssl_tls:
                smtp.starttls()
            smtp.login(username, password)
        security = "SSL/TLS" if self.ssl_tls else ("STARTTLS" if self.starttls else "plain SMTP")
        return f"SMTP connection verified for {sender_email} via {self.host}:{self.port} ({security})."


def create_mail_sender(config: SendConfig) -> MailSender:
    provider = _normalize_provider(config.mail_provider)
    if provider == "webtel_smtp":
        return SmtpMailSender(config, host="connect.webtelconnect.com", port=465, starttls=False, ssl_tls=True)
    host = config.smtp_host.strip() or "smtp.gmail.com"
    port = config.smtp_port if config.smtp_port > 0 else 587
    return SmtpMailSender(config, host=host, port=port, starttls=config.smtp_use_starttls, ssl_tls=config.smtp_use_ssl)


def send_with_fallback(config: SendConfig, row: ConfirmationRow, mail_template: MailTemplate, documents: DocumentResult) -> SendResult:
    errors: list[str] = []
    for provider in provider_chain(config):
        provider_config = replace(config, mail_provider=provider)
        try:
            return create_mail_sender(provider_config).send(row, mail_template, documents)
        except Exception as exc:
            errors.append(f"{provider}: {exc}")
    raise ValueError("All configured providers failed. " + " | ".join(errors))


def provider_chain(config: SendConfig) -> list[str]:
    chain: list[str] = []
    primary = _normalize_provider(config.mail_provider)
    if primary:
        chain.append(primary)
    for provider in (config.fallback_providers or "").split(","):
        normalized = _normalize_provider(provider)
        if normalized and normalized not in chain:
            chain.append(normalized)
    if not chain:
        chain.append("gmail_smtp")
    return chain


def _normalize_provider(provider: object) -> str:
    normalized = str(provider or "").strip().lower()
    if normalized in SUPPORTED_MAIL_PROVIDERS:
        return normalized
    if "webtel" in normalized:
        return "webtel_smtp"
    if "gmail" in normalized:
        return "gmail_smtp"
    return "gmail_smtp"


def build_email_message(
    row: ConfirmationRow,
    mail_template: MailTemplate,
    documents: DocumentResult,
    sender_email: str,
) -> tuple[EmailMessage, str, list[Path], list[str]]:
    sender_domain = sender_email.rsplit("@", 1)[-1].strip() if "@" in sender_email else "localhost"
    message_id = email.utils.make_msgid(idstring=f"{row.row_id}.{uuid.uuid4().hex}", domain=sender_domain)
    message = EmailMessage()
    message["From"] = sender_email
    message["To"] = row.email
    if row.cc:
        message["Cc"] = row.cc
    message["Subject"] = row.subject
    message["Date"] = email.utils.formatdate(localtime=True)
    message["Message-ID"] = message_id
    message.set_content(render_mail_body(row, mail_template))

    attachments = (documents.pdf_paths or documents.docx_paths) + documents.extra_attachment_paths
    for path in attachments:
        attach_path(message, path)
    recipients = [row.email] + [part.strip() for part in row.cc.split(";") if part.strip()]
    return message, message_id, attachments, recipients


def render_mail_body(row: ConfirmationRow, mail_template: MailTemplate) -> str:
    body = row.mail_body_override or mail_template.body_template_text
    replacements = {
        "{party_name}": row.party_name,
        "{amount}": format_amount(row.balance),
        "{{First Name}}": row.contact_first_name,
        "{{Last Name}}": row.contact_last_name,
        "March 31, 2025": row.balance_as_on_date,
        "31st March 2025": row.balance_as_on_date,
    }
    for old, new in replacements.items():
        body = body.replace(old, new)
    return body


def attach_path(message: EmailMessage, path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(f"Attachment not found: {path}")
    mime_type, _ = mimetypes.guess_type(path.name)
    if mime_type:
        maintype, subtype = mime_type.split("/", 1)
    else:
        maintype, subtype = "application", "octet-stream"
    message.add_attachment(path.read_bytes(), maintype=maintype, subtype=subtype, filename=path.name)


