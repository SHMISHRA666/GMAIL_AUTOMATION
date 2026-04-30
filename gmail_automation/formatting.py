from __future__ import annotations

from decimal import Decimal, InvalidOperation


def format_amount(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    normalized = text.replace(",", "")
    try:
        amount = Decimal(normalized)
    except InvalidOperation:
        return text
    return f"{amount:.2f}"
