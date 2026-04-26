from __future__ import annotations

from pathlib import Path

from .docx_utils import extract_docx_text
from .models import ConfirmationRow, DocumentResult, VerificationResult


class DocumentVerifier:
    UNRESOLVED_MARKERS = [
        "{{First Name}}",
        "{{Last Name}}",
        "Name of Vendor/customer",
        "Address of Vendor/customer",
        "INR\ufffd\ufffd\ufffd\ufffd\ufffd..",
        "INR\u2026\u2026\u2026\u2026\u2026..",
        "INR ---------------------------",
        "INR --------------",
        "INR -------------------",
        "March 31, 2025",
        "31st March 2025",
    ]

    def verify(self, row: ConfirmationRow, documents: DocumentResult) -> VerificationResult:
        errors: list[str] = []
        expected = documents.docx_paths + documents.pdf_paths
        for path in expected:
            if not path.exists():
                errors.append(f"Missing generated file: {path}")
            elif path.stat().st_size == 0:
                errors.append(f"Generated file is empty: {path}")

        for extra in documents.extra_attachment_paths:
            if not extra.exists():
                errors.append(f"Extra attachment not found: {extra}")

        combined_text = "\n".join(self._safe_extract(path) for path in documents.docx_paths)
        for marker in self.UNRESOLVED_MARKERS:
            if marker in combined_text:
                errors.append(f"Unresolved placeholder remains: {marker}")

        required_values = [
            row.party_name,
            row.address,
            row.balance_as_on_date,
            row.company_name,
            row.auditor_reply_email,
        ]
        if row.balance:
            required_values.append(str(row.balance))
        for value in required_values:
            if value and value not in combined_text:
                errors.append(f"Expected value not found in generated DOCX text: {value}")

        if documents.pdf_paths and len(documents.pdf_paths) != len(documents.docx_paths):
            errors.append("PDF count does not match DOCX count")

        return VerificationResult(passed=not errors, errors=errors)

    def _safe_extract(self, path: Path) -> str:
        try:
            return extract_docx_text(path)
        except Exception:
            return ""
