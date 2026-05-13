from __future__ import annotations

import hashlib
from pathlib import Path

from .docx_utils import replace_docx_text
from .errors import DocumentGenerationError
from .formatting import format_amount
from .models import ConfirmationRow, DocumentResult, now_text
from .templates import TemplateRepository

PARTY_NAME_MARKER = "\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026..(Name of Vendor/Customer)"


class DocumentGenerator:
    DOC_TEMPLATES = [
        ("balance_confirmation_letter", "Balance confirmation letter.docx"),
        ("vendor_reply_form", "On Vendor letter.docx"),
    ]

    def __init__(self, work_dir: Path, convert_to_pdf: bool = True) -> None:
        self.work_dir = work_dir
        self.generated_dir = work_dir / "generated"
        self.convert_to_pdf = convert_to_pdf

    def generate(self, row: ConfirmationRow, templates: TemplateRepository) -> DocumentResult:
        party_dir = self.generated_dir / _safe_name(row.row_id)
        party_dir.mkdir(parents=True, exist_ok=True)
        docx_paths: list[Path] = []
        pdf_paths: list[Path] = []
        warnings: list[str] = []

        try:
            for template_name, output_name in self.DOC_TEMPLATES:
                rendered = replace_docx_text(
                    templates.load_docx_template(template_name),
                    self._replacements(row),
                    black_replacements={PARTY_NAME_MARKER},
                )
                docx_path = party_dir / output_name
                docx_path.write_bytes(rendered)
                docx_paths.append(docx_path)
                if self.convert_to_pdf:
                    pdf_path = docx_path.with_suffix(".pdf")
                    self._convert_to_pdf(docx_path, pdf_path)
                    pdf_paths.append(pdf_path)
        except Exception as exc:
            raise DocumentGenerationError(str(exc)) from exc

        static_pdf_path = templates.materialize_static_authorisation_pdf(self.work_dir)
        hash_value = self._hash_files((pdf_paths or docx_paths) + [static_pdf_path])
        return DocumentResult(
            docx_paths=docx_paths,
            pdf_paths=pdf_paths,
            extra_attachment_paths=[static_pdf_path],
            created_at=now_text(),
            template_version=templates.version(),
            attachment_hash=hash_value,
            warnings=warnings,
        )

    def _replacements(self, row: ConfirmationRow) -> dict[str, str]:
        return {
            PARTY_NAME_MARKER: row.party_name,
            "8,75,000.00": format_amount(row.balance),
        }

    def _convert_to_pdf(self, docx_path: Path, pdf_path: Path) -> None:
        pythoncom = None
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except Exception as exc:
            raise DocumentGenerationError("Microsoft Word automation is unavailable for PDF conversion") from exc

        word = None
        doc = None
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(docx_path))
            doc.SaveAs(str(pdf_path), FileFormat=17)
        finally:
            if doc is not None:
                doc.Close(False)
            if word is not None:
                word.Quit()
            if com_initialized and pythoncom is not None:
                pythoncom.CoUninitialize()

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
