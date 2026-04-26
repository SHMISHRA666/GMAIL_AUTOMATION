from __future__ import annotations

from datetime import datetime, timedelta

from .models import RetryDecision


class RetryPolicy:
    MAX_ATTEMPTS = {
        "document_generation": 2,
        "pdf_conversion": 3,
        "excel_save": 3,
        "gmail_send": 3,
        "gmail_tracking": 3,
    }

    DELAYS = {
        "document_generation": [5, 15],
        "pdf_conversion": [5, 15, 30],
        "excel_save": [5, 15, 30],
        "gmail_send": [30, 120, 300],
        "gmail_tracking": [30, 120, 300],
    }

    def next_attempt(self, step: str, attempt_count: int, retryable: bool) -> RetryDecision:
        if not retryable:
            return RetryDecision(False, "", True)
        max_attempts = self.MAX_ATTEMPTS.get(step, 1)
        if attempt_count >= max_attempts:
            return RetryDecision(False, "", True)
        delays = self.DELAYS.get(step, [30])
        delay = delays[min(attempt_count, len(delays) - 1)]
        next_retry_at = (datetime.now() + timedelta(seconds=delay)).strftime("%Y-%m-%d %H:%M:%S")
        return RetryDecision(True, next_retry_at, False)
