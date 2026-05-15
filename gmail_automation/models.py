from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path


@dataclass
class RowState:
    attachment_created: str = "N"
    generated_docx_paths: list[Path] = field(default_factory=list)
    generated_pdf_paths: list[Path] = field(default_factory=list)
    attachment_created_at: str = ""
    ready_to_send: str = "N"
    main_sent: str = "N"
    sent_date: str = ""
    gmail_message_id: str = ""
    gmail_thread_id: str = ""
    bounce_received: str = "N"
    bounce_date: str = ""
    reply_received: str = "N"
    reply_received_date: str = ""
    last_checked_at: str = ""
    attempt_count: int = 0
    last_attempt_at: str = ""
    last_successful_step: str = ""
    next_retry_at: str = ""
    retry_locked: str = "N"
    batch_id: str = ""
    batch_sequence: str = ""
    verification_status: str = ""
    verification_errors: str = ""
    last_log_file: str = ""
    status: str = "Pending"
    error: str = ""
    template_version: str = ""
    generated_attachment_hash: str = ""


@dataclass
class ConfirmationRow:
    row_id: str
    excel_row_number: int
    party: str
    party_name: str
    contact_name: str
    contact_first_name: str
    contact_last_name: str
    email: str
    cc: str
    subject: str
    balance: str
    balance_nature: str
    company_name: str
    address: str
    phone: str
    balance_as_on_date: str
    letter_date: str
    auditor_reply_email: str
    mail_body_override: str
    extra_attachment_paths: list[Path]
    state: RowState = field(default_factory=RowState)


@dataclass
class MailTemplate:
    subject_template: str
    body_template_text: str
    required_fields: set[str]


@dataclass
class DocumentResult:
    docx_paths: list[Path]
    pdf_paths: list[Path]
    extra_attachment_paths: list[Path]
    created_at: str
    template_version: str
    attachment_hash: str
    warnings: list[str] = field(default_factory=list)


@dataclass
class VerificationResult:
    passed: bool
    errors: list[str]


@dataclass
class SendResult:
    sent_at: str
    smtp_message_id: str
    gmail_thread_id: str
    to_email: str
    cc_email: str
    attachment_paths: list[Path]


@dataclass
class TrackingResult:
    bounce_received: bool = False
    bounce_date: str = ""
    reply_received: bool = False
    reply_received_date: str = ""
    last_checked_at: str = ""
    evidence_subject: str = ""


@dataclass
class ErrorInfo:
    category: str
    retryable: bool
    message: str


@dataclass
class RetryDecision:
    should_retry: bool
    next_retry_at: str
    locked: bool


@dataclass
class SendConfig:
    mail_provider: str = "gmail_smtp"
    fallback_providers: str = ""
    auth_type: str = "smtp_password"
    sender_email: str = ""
    app_password: str = ""
    smtp_password: str = ""
    smtp_host: str = ""
    smtp_port: int = 0
    smtp_use_starttls: bool = True
    smtp_use_ssl: bool = False
    smtp_username: str = ""
    batch_size: int = 20
    document_generation_workers: int = 4
    batch_delay_seconds: int = 60
    per_email_delay_seconds: int = 3
    daily_send_limit: int = 100
    dry_run_limit: int = 3
    convert_to_pdf: bool = True
    send_mode: str = "preview"


@dataclass
class SendBatch:
    batch_id: str
    rows: list[ConfirmationRow]


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
