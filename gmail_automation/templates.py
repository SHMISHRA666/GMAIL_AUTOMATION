from __future__ import annotations

import hashlib
from importlib import resources
from pathlib import Path

from .docx_utils import extract_docx_text
from .models import MailTemplate

RESOURCE_PACKAGE = "gmail_automation.resources"

STATIC_AUTHORISATION_PDF = "Authorisation for Direct Balance Confirmation.pdf"
DOCX_TEMPLATE_FILES = {
    "balance_confirmation_letter": "Balance confirmation letter.docx",
    "vendor_reply_form": "On Vendor letter.docx",
}

RENDERER_VERSION = "renderer-v8"
SUBJECT_TEMPLATE = "Confirmation of balance outstanding as on 31st March 2026"
BALANCE_AS_ON_DATE = "31st March 2026"
AUDITOR_REPLY_EMAIL = "ghanshyam@ngmks.in"

MAIL_BODY_TEMPLATE = """Dear Sir/Madam,

We, NGMKS & Associates, Chartered Accountants, Statutory Auditors of Purple United Sales Limited for the period ended 31st March 2026. The Company records show that your account at the close of business on 31st March 2026 has an outstanding receivable /payable balance from/to you of INR {amount} receivable/payable from you

We would be obliged if you could compare the above balance with your records and reply/confirm directly to us, within 7 days from the date of receipt of this letter at the below mentioned address after completing the attached confirmation form, as applicable.

Kind Attn: Nitin Goyal/Ghanshyam Kumar
Postal Address:
NGMKS & Associates
Chartered Accountants
811, 8th Floor, Wave Silver Tower
Sector-18, Noida - 201301, Uttar Pradesh
OR
Email Address: ghanshyam@ngmks.in

This confirmation also need to be sent through e-mail/postal address by attaching scan copy of Attached Form duly signed and stamped,

Thanks in advance for your cooperation in this regards, if such balance confirmation not received with stipulated time limit, we presumed above given balance are true and correct.

Thanks


CA NITIN GOYAL
Partner
NGMKS & Associates
D-65, Flatted Factory Complex
Jhandewalan
New Delhi, India-110055
011-45652955
+91-8800227843"""


class TemplateRepository:
    def __init__(self) -> None:
        self._cache: dict[str, bytes] = {}

    def version(self) -> str:
        digest = hashlib.sha256()
        digest.update(RENDERER_VERSION.encode("utf-8"))
        digest.update(SUBJECT_TEMPLATE.encode("utf-8"))
        digest.update(MAIL_BODY_TEMPLATE.encode("utf-8"))
        for name in sorted(DOCX_TEMPLATE_FILES):
            digest.update(name.encode("utf-8"))
            digest.update(self.load_docx_template(name))
        try:
            digest.update(self.load_static_authorisation_pdf_bytes())
        except FileNotFoundError:
            pass
        return digest.hexdigest()[:16]

    def load_docx_template(self, template_name: str) -> bytes:
        if template_name not in DOCX_TEMPLATE_FILES:
            raise ValueError(f"Unknown template: {template_name}")
        if template_name not in self._cache:
            self._cache[template_name] = resources.files(RESOURCE_PACKAGE).joinpath(DOCX_TEMPLATE_FILES[template_name]).read_bytes()
        return self._cache[template_name]

    def load_static_authorisation_pdf_bytes(self) -> bytes:
        return resources.files(RESOURCE_PACKAGE).joinpath(STATIC_AUTHORISATION_PDF).read_bytes()

    def materialize_static_authorisation_pdf(self, work_dir: Path) -> Path:
        static_dir = work_dir / "generated" / "_static"
        static_dir.mkdir(parents=True, exist_ok=True)
        path = static_dir / STATIC_AUTHORISATION_PDF
        data = self.load_static_authorisation_pdf_bytes()
        if not path.exists() or path.read_bytes() != data:
            path.write_bytes(data)
        return path

    def load_mail_template(self) -> MailTemplate:
        return MailTemplate(
            subject_template=SUBJECT_TEMPLATE,
            body_template_text=MAIL_BODY_TEMPLATE,
            required_fields={"Party Name", "Email To(Address)", "Balance"},
        )

    def validate_placeholders(self, available_fields: set[str]) -> list[str]:
        missing: list[str] = []
        mail_template = self.load_mail_template()
        missing.extend(sorted(mail_template.required_fields - available_fields))
        try:
            self.load_static_authorisation_pdf_bytes()
        except FileNotFoundError:
            missing.append(f"static PDF not found: {STATIC_AUTHORISATION_PDF}")
        for template_name, filename in DOCX_TEMPLATE_FILES.items():
            try:
                text = extract_docx_text(self.load_docx_template(template_name))
            except FileNotFoundError:
                missing.append(f"docx template not found: {filename}")
                continue
            if not text.strip():
                missing.append(f"{template_name}: no template text")
        return missing
