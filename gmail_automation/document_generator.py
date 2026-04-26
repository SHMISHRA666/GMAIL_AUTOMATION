from __future__ import annotations

import hashlib
from pathlib import Path

from .docx_utils import replace_docx_text
from .errors import DocumentGenerationError
from .models import ConfirmationRow, DocumentResult, now_text
from .templates import TemplateRepository


class DocumentGenerator:
    DOC_TEMPLATES = [
        ("authorisation_letter", "Authorisation_Letter"),
        ("balance_confirmation_letter", "Balance_Confirmation_Letter"),
        ("reply_form", "Reply_Form"),
    ]

    def __init__(self, work_dir: Path, convert_to_pdf: bool = True) -> None:
        self.work_dir = work_dir
        self.generated_dir = work_dir / "generated"
        self.convert_to_pdf = convert_to_pdf

    def generate(self, row: ConfirmationRow, templates: TemplateRepository) -> DocumentResult:
        party_dir = self.generated_dir / _safe_name(row.row_id)
        party_dir.mkdir(parents=True, exist_ok=True)
        replacements = self._replacements(row)
        docx_paths: list[Path] = []
        pdf_paths: list[Path] = []
        warnings: list[str] = []

        try:
            for template_name, output_name in self.DOC_TEMPLATES:
                template_bytes = templates.load_docx_template(template_name)
                rendered = replace_docx_text(template_bytes, replacements)
                docx_path = party_dir / f"{_safe_name(row.row_id)}_{output_name}.docx"
                docx_path.write_bytes(rendered)
                docx_paths.append(docx_path)
                if self.convert_to_pdf:
                    pdf_path = docx_path.with_suffix(".pdf")
                    self._convert_to_pdf(docx_path, pdf_path)
                    pdf_paths.append(pdf_path)
        except Exception as exc:
            raise DocumentGenerationError(str(exc)) from exc

        hash_value = self._hash_files(pdf_paths or docx_paths)
        return DocumentResult(
            docx_paths=docx_paths,
            pdf_paths=pdf_paths,
            extra_attachment_paths=row.extra_attachment_paths,
            created_at=now_text(),
            template_version=templates.version(),
            attachment_hash=hash_value,
            warnings=warnings,
        )

    def _replacements(self, row: ConfirmationRow) -> dict[str, str]:
        balance_sentence = f"{row.balance_nature} of INR {row.balance}".strip()
        return {
            "Date:": f"Date: {row.letter_date}",
            "Name of Vendor/customer": row.party_name,
            "Address of Vendor/customer": row.address,
            "31st March 2026": row.balance_as_on_date,
            "31st March,2026": row.balance_as_on_date,
            "31st March 2025": row.balance_as_on_date,
            "March 31, 2025": row.balance_as_on_date,
            "Purple United Sales Limited": row.company_name,
            "ghanshyam@ngmks.in": row.auditor_reply_email,
            "INR\ufffd\ufffd\ufffd\ufffd\ufffd..": f"INR {row.balance} {row.balance_nature}",
            "INR\u2026\u2026\u2026\u2026\u2026..": f"INR {row.balance} {row.balance_nature}",
            "receivable /payable balance from/to you of INR\ufffd\ufffd\ufffd\ufffd\ufffd..": balance_sentence,
            "receivable /payable balance from/to you of INR\u2026\u2026\u2026\u2026\u2026..": balance_sentence,
        }

    def _convert_to_pdf(self, docx_path: Path, pdf_path: Path) -> None:
        try:
            import win32com.client  # type: ignore
        except Exception as exc:
            raise DocumentGenerationError("Microsoft Word automation is unavailable for PDF conversion") from exc

        word = None
        doc = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(docx_path))
            doc.SaveAs(str(pdf_path), FileFormat=17)
        finally:
            if doc is not None:
                doc.Close(False)
            if word is not None:
                word.Quit()

    def _hash_files(self, paths: list[Path]) -> str:
        digest = hashlib.sha256()
        for path in paths:
            digest.update(path.name.encode("utf-8"))
            digest.update(path.read_bytes())
        return digest.hexdigest()


def _safe_name(value: str) -> str:
    allowed = []
    for char in value.strip():
        allowed.append(char if char.isalnum() or char in ("-", "_") else "_")
    return "".join(allowed) or "row"
