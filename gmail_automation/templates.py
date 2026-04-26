from __future__ import annotations

import hashlib
from importlib import resources

from .docx_utils import extract_docx_text, normalize_mail_template_text
from .models import MailTemplate

RESOURCE_PACKAGE = "gmail_automation.resources"

TEMPLATE_FILES = {
    "authorisation_letter": "authorisation_letter.docx",
    "balance_confirmation_letter": "balance_confirmation_letter.docx",
    "reply_form": "reply_form.docx",
    "mail_template": "mail_template.docx",
}

RENDERER_VERSION = "renderer-v4"


class TemplateRepository:
    def __init__(self) -> None:
        self._cache: dict[str, bytes] = {}

    def version(self) -> str:
        digest = hashlib.sha256()
        digest.update(RENDERER_VERSION.encode("utf-8"))
        for name in sorted(TEMPLATE_FILES):
            digest.update(name.encode("utf-8"))
            digest.update(self.load_docx_template(name))
        return digest.hexdigest()[:16]

    def load_docx_template(self, template_name: str) -> bytes:
        if template_name not in TEMPLATE_FILES:
            raise ValueError(f"Unknown template: {template_name}")
        if template_name not in self._cache:
            resource = resources.files(RESOURCE_PACKAGE).joinpath(TEMPLATE_FILES[template_name])
            self._cache[template_name] = resource.read_bytes()
        return self._cache[template_name]

    def load_mail_template(self) -> MailTemplate:
        raw_text = extract_docx_text(self.load_docx_template("mail_template"))
        body = normalize_mail_template_text(raw_text)
        return MailTemplate(
            subject_template="{subject}",
            body_template_text=body,
            required_fields={"ContactFirstName", "ContactLastName", "BalanceAsOnDate"},
        )

    def validate_placeholders(self, available_fields: set[str]) -> list[str]:
        missing: list[str] = []
        mail_template = self.load_mail_template()
        missing.extend(sorted(mail_template.required_fields - available_fields))
        for template_name in ("authorisation_letter", "balance_confirmation_letter", "reply_form"):
            text = extract_docx_text(self.load_docx_template(template_name))
            if not text.strip():
                missing.append(f"{template_name}: no readable text")
        return missing
