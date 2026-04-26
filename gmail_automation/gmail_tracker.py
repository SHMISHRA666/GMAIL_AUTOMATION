from __future__ import annotations

import imaplib
import email
from email.header import decode_header

from .errors import TrackingError
from .models import ConfirmationRow, SendConfig, TrackingResult, now_text


class GmailTracker:
    IMAP_HOST = "imap.gmail.com"

    def __init__(self, config: SendConfig) -> None:
        self.config = config

    def check(self, row: ConfirmationRow) -> TrackingResult:
        if not self.config.sender_email or not self.config.app_password:
            raise TrackingError("Gmail sender email and app password are required for tracking")

        result = TrackingResult(last_checked_at=now_text())
        try:
            with imaplib.IMAP4_SSL(self.IMAP_HOST) as imap:
                imap.login(self.config.sender_email, self.config.app_password)
                imap.select("INBOX")
                bounce = self._search_bounce(imap, row)
                if bounce:
                    result.bounce_received = True
                    result.bounce_date = bounce["date"]
                    result.evidence_subject = bounce["subject"]
                reply = self._search_reply(imap, row)
                if reply:
                    result.reply_received = True
                    result.reply_received_date = reply["date"]
                    result.evidence_subject = reply["subject"]
        except Exception as exc:
            raise TrackingError(str(exc)) from exc
        return result

    def _search_bounce(self, imap: imaplib.IMAP4_SSL, row: ConfirmationRow) -> dict[str, str] | None:
        criteria = '(OR FROM "mailer-daemon" FROM "Mail Delivery Subsystem")'
        return self._first_matching(imap, criteria, row)

    def _search_reply(self, imap: imaplib.IMAP4_SSL, row: ConfirmationRow) -> dict[str, str] | None:
        subject = row.subject.replace('"', "")
        criteria = f'(OR FROM "{row.email}" SUBJECT "{subject}")'
        return self._first_matching(imap, criteria, row)

    def _first_matching(self, imap: imaplib.IMAP4_SSL, criteria: str, row: ConfirmationRow) -> dict[str, str] | None:
        status, data = imap.search(None, criteria)
        if status != "OK" or not data or not data[0]:
            return None
        for msg_id in reversed(data[0].split()[-20:]):
            status, msg_data = imap.fetch(msg_id, "(RFC822)")
            if status != "OK" or not msg_data:
                continue
            for part in msg_data:
                if not isinstance(part, tuple):
                    continue
                msg = email.message_from_bytes(part[1])
                subject = self._decode_header(msg.get("Subject", ""))
                if row.subject.lower() in subject.lower() or row.email.lower() in str(msg.get("From", "")).lower():
                    return {"date": msg.get("Date", now_text()), "subject": subject}
        return None

    def _decode_header(self, value: str) -> str:
        parts = decode_header(value)
        output = []
        for text, encoding in parts:
            if isinstance(text, bytes):
                output.append(text.decode(encoding or "utf-8", errors="replace"))
            else:
                output.append(text)
        return "".join(output)
