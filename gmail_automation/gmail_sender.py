from __future__ import annotations

from .mail_sender import SmtpMailSender
from .models import SendConfig


class GmailSender(SmtpMailSender):
    SMTP_HOST = "smtp.gmail.com"
    SMTP_PORT = 587

    def __init__(self, config: SendConfig) -> None:
        config.mail_provider = "gmail_smtp"
        if not config.smtp_host:
            config.smtp_host = self.SMTP_HOST
        if not config.smtp_port:
            config.smtp_port = self.SMTP_PORT
        if not config.smtp_password:
            config.smtp_password = config.app_password
        super().__init__(config=config, host=config.smtp_host, port=config.smtp_port, starttls=config.smtp_use_starttls, ssl_tls=config.smtp_use_ssl)
