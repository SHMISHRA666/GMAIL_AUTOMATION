from __future__ import annotations

import json
import os
from pathlib import Path

from .models import SendConfig


def _as_bool(value: object, default: bool) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "y", "on"}:
        return True
    if text in {"0", "false", "no", "n", "off"}:
        return False
    return default


def load_config(path: Path | None) -> SendConfig:
    data: dict = {}
    if path and path.exists():
        data = json.loads(path.read_text(encoding="utf-8"))
    smtp_port = int(data.get("smtp_port", os.environ.get("SMTP_PORT", "0") or "0"))
    return SendConfig(
        mail_provider=str(data.get("mail_provider") or os.environ.get("MAIL_PROVIDER", "gmail_smtp")),
        fallback_providers=str(data.get("fallback_providers") or os.environ.get("MAIL_FALLBACK_PROVIDERS", "")),
        auth_type=str(data.get("auth_type") or os.environ.get("MAIL_AUTH_TYPE", "smtp_password")),
        sender_email=data.get("sender_email") or os.environ.get("GMAIL_SENDER", ""),
        app_password=data.get("app_password") or os.environ.get("GMAIL_APP_PASSWORD", ""),
        smtp_password=data.get("smtp_password") or os.environ.get("SMTP_PASSWORD", ""),
        smtp_host=str(data.get("smtp_host") or os.environ.get("SMTP_HOST", "")),
        smtp_port=smtp_port,
        smtp_use_starttls=_as_bool(data.get("smtp_use_starttls", os.environ.get("SMTP_USE_STARTTLS")), True),
        smtp_use_ssl=_as_bool(data.get("smtp_use_ssl", os.environ.get("SMTP_USE_SSL")), False),
        smtp_username=str(data.get("smtp_username") or os.environ.get("SMTP_USERNAME", "")),
        batch_size=int(data.get("batch_size", 20)),
        document_generation_workers=max(1, int(data.get("document_generation_workers", 4))),
        batch_delay_seconds=int(data.get("batch_delay_seconds", 60)),
        per_email_delay_seconds=int(data.get("per_email_delay_seconds", 3)),
        daily_send_limit=int(data.get("daily_send_limit", 100)),
        dry_run_limit=int(data.get("dry_run_limit", 3)),
        convert_to_pdf=_as_bool(data.get("convert_to_pdf"), True),
        send_mode=str(data.get("send_mode", "preview")),
    )


def write_sample_config(path: Path) -> None:
    sample = {
        "mail_provider": "gmail_smtp",
        "fallback_providers": "webtel_smtp",
        "auth_type": "smtp_password",
        "sender_email": "your.gmail@gmail.com",
        "app_password": "gmail-app-password",
        "smtp_password": "",
        "smtp_host": "smtp.gmail.com",
        "smtp_port": 587,
        "smtp_use_starttls": True,
        "smtp_use_ssl": False,
        "smtp_username": "",
        "batch_size": 20,
        "document_generation_workers": 4,
        "batch_delay_seconds": 60,
        "per_email_delay_seconds": 3,
        "daily_send_limit": 100,
        "dry_run_limit": 3,
        "convert_to_pdf": True,
        "send_mode": "preview",
    }
    path.write_text(json.dumps(sample, indent=2), encoding="utf-8")
