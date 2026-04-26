from __future__ import annotations

import smtplib
import socket

from .models import ErrorInfo


class ValidationError(Exception):
    pass


class TemplateError(Exception):
    pass


class DocumentGenerationError(Exception):
    pass


class FileLockError(Exception):
    pass


class TrackingError(Exception):
    pass


class ErrorClassifier:
    def classify(self, step: str, error: Exception) -> ErrorInfo:
        message = str(error)
        if isinstance(error, ValidationError):
            return ErrorInfo("ValidationError", False, message)
        if isinstance(error, TemplateError):
            return ErrorInfo("TemplateError", False, message)
        if isinstance(error, FileLockError) or "Permission denied" in message:
            return ErrorInfo("FileLockError", True, message)
        if isinstance(error, DocumentGenerationError):
            return ErrorInfo("DocumentGenerationError", True, message)
        if isinstance(error, smtplib.SMTPAuthenticationError):
            return ErrorInfo("GmailPermanentError", False, "Gmail authentication failed")
        if isinstance(error, smtplib.SMTPRecipientsRefused):
            return ErrorInfo("GmailPermanentError", False, message)
        if isinstance(error, (smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError, socket.timeout, TimeoutError)):
            return ErrorInfo("GmailTemporaryError", True, message)
        if isinstance(error, TrackingError):
            return ErrorInfo("TrackingError", True, message)
        if step.lower().startswith("validation"):
            return ErrorInfo("ValidationError", False, message)
        return ErrorInfo("UnexpectedError", False, message)
