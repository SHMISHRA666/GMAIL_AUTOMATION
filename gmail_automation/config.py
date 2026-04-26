from __future__ import annotations

import json
import os
from pathlib import Path

from .models import SendConfig


def load_config(path: Path | None) -> SendConfig:
    data: dict = {}
    if path and path.exists():
        data = json.loads(path.read_text(encoding="utf-8"))
    return SendConfig(
        sender_email=data.get("sender_email") or os.environ.get("GMAIL_SENDER", ""),
        app_password=data.get("app_password") or os.environ.get("GMAIL_APP_PASSWORD", ""),
        batch_size=int(data.get("batch_size", 20)),
        batch_delay_seconds=int(data.get("batch_delay_seconds", 60)),
        per_email_delay_seconds=int(data.get("per_email_delay_seconds", 3)),
        daily_send_limit=int(data.get("daily_send_limit", 100)),
        dry_run_limit=int(data.get("dry_run_limit", 3)),
        convert_to_pdf=bool(data.get("convert_to_pdf", True)),
        send_mode=str(data.get("send_mode", "preview")),
    )


def write_sample_config(path: Path) -> None:
    sample = {
        "sender_email": "your.gmail@gmail.com",
        "app_password": "gmail-app-password",
        "batch_size": 20,
        "batch_delay_seconds": 60,
        "per_email_delay_seconds": 3,
        "daily_send_limit": 100,
        "dry_run_limit": 3,
        "convert_to_pdf": True,
        "send_mode": "preview",
    }
    path.write_text(json.dumps(sample, indent=2), encoding="utf-8")
