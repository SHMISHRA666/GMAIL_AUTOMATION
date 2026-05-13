from __future__ import annotations

import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path

from .batch import BatchPlanner
from .config import load_config
from .document_generator import DocumentGenerator
from .errors import ErrorClassifier, ValidationError
from .excel_store import ExcelStateStore
from .gmail_sender import GmailSender
from .logging_utils import AppLogger
from .models import ConfirmationRow, DocumentResult, SendBatch, SendConfig, VerificationResult
from .retry import RetryPolicy
from .templates import TemplateRepository
from .verifier import DocumentVerifier


@dataclass
class CustomerStatus:
    row_id: str
    party_name: str
    email: str
    batch_id: str
    batch_sequence: str
    docs_created: str
    batch_selected: str
    mail_sent: str
    status: str
    error: str


@dataclass
class WorkflowSummary:
    total_rows: int
    ready_to_send: int
    sent_rows: int
    failed_rows: int


@dataclass
class BatchRunResult:
    batch_id: str
    attempted: int
    sent: int
    failed: int


@dataclass
class DocumentGenerationOutcome:
    row: ConfirmationRow
    result: DocumentResult | None = None
    verification: VerificationResult | None = None
    skipped: bool = False


class GmailAutomationWorkflow:
    def __init__(self, config_path: Path | None = None) -> None:
        self.config_path = config_path
        self.master_path: Path | None = None
        self.work_dir: Path | None = None
        self.config: SendConfig = load_config(config_path if config_path and config_path.exists() else None)
        self.store: ExcelStateStore | None = None
        self.templates = TemplateRepository()
        self.error_classifier = ErrorClassifier()
        self.retry_policy = RetryPolicy()
        self.verifier = DocumentVerifier()
        self.logger: AppLogger | None = None
        self._sent_in_session = 0

    def load_input_file(self, master_path: Path, batch_size: int) -> WorkflowSummary:
        self.master_path = master_path
        self.work_dir = master_path.parent
        self.config.batch_size = batch_size
        self.store = ExcelStateStore(master_path)
        self.logger = AppLogger(self.work_dir)
        self.logger.start_run("ui")
        self._startup_validation()
        return self.get_summary()

    def set_batch_size(self, batch_size: int) -> None:
        self.config.batch_size = batch_size

    def generate_documents(self) -> WorkflowSummary:
        self._require_loaded()
        assert self.store is not None
        assert self.work_dir is not None
        rows = self.store.load_rows()
        self.templates.version()
        self.templates.materialize_static_authorisation_pdf(self.work_dir)

        for batch in self._row_batches(rows, self.config.batch_size):
            for row in batch:
                self.store.mark_log_file(row.row_id, self._log_file_text())

            workers = self._document_generation_worker_count(len(batch))
            with ThreadPoolExecutor(max_workers=workers) as executor:
                futures = {executor.submit(self._generate_document_outcome, row): row for row in batch}
                for future in as_completed(futures):
                    row = futures[future]
                    try:
                        outcome = future.result()
                        self._record_document_outcome(outcome)
                    except Exception as exc:
                        self._handle_error(row, "document_generation", exc)
        return self.get_summary()

    def preview_next_batch(self) -> SendBatch | None:
        self._require_loaded()
        assert self.store is not None
        batches = BatchPlanner().plan(self.store.load_rows(), self.config)
        return batches[0] if batches else None

    def run_next_batch(self) -> BatchRunResult:
        batch = self.preview_next_batch()
        if batch is None:
            return BatchRunResult(batch_id="", attempted=0, sent=0, failed=0)
        return self._send_rows(batch.rows, batch.batch_id)

    def resend_failed(self, batch_id: str | None = None, row_ids: set[str] | None = None) -> BatchRunResult:
        return self.send_pending(batch_id=batch_id, row_ids=row_ids)

    def send_pending(self, batch_id: str | None = None, row_ids: set[str] | None = None) -> BatchRunResult:
        self._require_loaded()
        assert self.store is not None
        rows = [row for row in self.store.load_rows() if self._can_send_pending(row)]
        if batch_id:
            rows = [row for row in rows if row.state.batch_id == batch_id]
        if row_ids:
            rows = [row for row in rows if row.row_id in row_ids]
        send_batch_id = batch_id or f"manual_{int(time.time())}"
        return self._send_rows(rows, send_batch_id)

    def get_results(self) -> list[CustomerStatus]:
        self._require_loaded()
        assert self.store is not None
        return [self._status_from_row(row) for row in self.store.load_rows()]

    def get_batch_ids(self) -> list[str]:
        return sorted({row.batch_id for row in self.get_results() if row.batch_id})

    def get_summary(self) -> WorkflowSummary:
        self._require_loaded()
        assert self.store is not None
        rows = self.store.load_rows()
        return WorkflowSummary(
            total_rows=len(rows),
            ready_to_send=sum(1 for row in rows if row.state.ready_to_send == "Y"),
            sent_rows=sum(1 for row in rows if row.state.main_sent == "Y"),
            failed_rows=sum(1 for row in rows if row.state.error),
        )

    def _send_rows(self, rows: list[ConfirmationRow], batch_id: str) -> BatchRunResult:
        self._require_loaded()
        assert self.store is not None
        if not rows:
            return BatchRunResult(batch_id=batch_id, attempted=0, sent=0, failed=0)
        if self.config.send_mode != "send":
            raise ValidationError("Set send_mode to 'send' in config.json before sending emails.")

        sender = GmailSender(self.config)
        mail_template = self.templates.load_mail_template()
        sent = 0
        failed = 0
        for sequence, row in enumerate(rows, start=1):
            self.store.mark_batch(row.row_id, batch_id, sequence)
            self.store.mark_log_file(row.row_id, self._log_file_text())
            try:
                if self._sent_in_session >= self.config.daily_send_limit:
                    raise ValidationError(f"Daily send limit reached: {self.config.daily_send_limit}")
                self._validate_row(row)
                if row.state.main_sent == "Y":
                    continue
                result = sender.send(row, mail_template, self._documents_from_state(row))
                self.store.mark_send_success(row.row_id, result)
                self._log_info(row.row_id, "gmail_send", "Email sent", {"to": row.email, "message_id": result.smtp_message_id})
                sent += 1
                self._sent_in_session += 1
                time.sleep(self.config.per_email_delay_seconds)
            except Exception as exc:
                failed += 1
                self._handle_error(row, "gmail_send", exc)
        return BatchRunResult(batch_id=batch_id, attempted=len(rows), sent=sent, failed=failed)

    def _generate_document_outcome(self, row: ConfirmationRow) -> DocumentGenerationOutcome:
        self._validate_row(row)
        if self._can_skip_generation(row):
            return DocumentGenerationOutcome(row=row, skipped=True)

        generator = DocumentGenerator(self._require_work_dir(), convert_to_pdf=self.config.convert_to_pdf)
        templates = TemplateRepository()
        result = generator.generate(row, templates)
        verification = DocumentVerifier().verify(row, result)
        return DocumentGenerationOutcome(row=row, result=result, verification=verification)

    def _record_document_outcome(self, outcome: DocumentGenerationOutcome) -> None:
        assert self.store is not None
        if outcome.skipped:
            self._log_info(outcome.row.row_id, "document_generation", "Skipping existing generated documents", {})
            return

        if outcome.result is None or outcome.verification is None:
            raise ValidationError("Document generation completed without a result")

        if outcome.verification.passed:
            self.store.mark_documents_created(outcome.row.row_id, outcome.result)
            self._log_info(outcome.row.row_id, "document_generation", "Documents generated and verified", {})
        else:
            errors = "; ".join(outcome.verification.errors)
            self.store.mark_documents_created(outcome.row.row_id, outcome.result, "Failed", errors)
            self._log_error(
                outcome.row.row_id,
                "document_verification",
                "Generated documents failed verification",
                {"errors": outcome.verification.errors},
            )

    def _row_batches(self, rows: list[ConfirmationRow], batch_size: int) -> list[list[ConfirmationRow]]:
        size = max(1, batch_size)
        return [rows[index : index + size] for index in range(0, len(rows), size)]

    def _document_generation_worker_count(self, row_count: int) -> int:
        return max(1, min(row_count, self.config.document_generation_workers))

    def _startup_validation(self) -> None:
        assert self.store is not None
        errors = self.store.validate()
        errors.extend(
            self.templates.validate_placeholders(
                {
                    "S.No.",
                    "Party Type",
                    "Party Name",
                    "Email To(Address)",
                    "Balance",
                }
            )
        )
        if errors:
            raise ValidationError("; ".join(errors))
        self.store.ensure_workbooks()
        self._log_info(None, "startup_validation", "Startup validation passed", {"template_version": self.templates.version()})

    def _validate_row(self, row: ConfirmationRow) -> None:
        missing = []
        if not row.row_id:
            missing.append("PartyId")
        if not row.email:
            missing.append("Email To(Address)")
        if not row.subject:
            missing.append("Subject")
        if not row.party_name:
            missing.append("Party Name")
        if not row.balance:
            missing.append("Balance")
        if missing:
            raise ValidationError(f"Missing required fields: {', '.join(missing)}")
        if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", row.email):
            raise ValidationError(f"Invalid email address: {row.email}")

    def _can_skip_generation(self, row: ConfirmationRow) -> bool:
        state = row.state
        if state.verification_status != "Passed" or state.ready_to_send != "Y":
            return False
        if state.attachment_created != "Y" or state.template_version != self.templates.version():
            return False
        paths = state.generated_docx_paths + state.generated_pdf_paths
        return bool(paths) and all(path.exists() and path.stat().st_size > 0 for path in paths)

    def _can_send_pending(self, row: ConfirmationRow) -> bool:
        return row.state.ready_to_send == "Y" and row.state.verification_status == "Passed" and row.state.main_sent != "Y"

    def _documents_from_state(self, row: ConfirmationRow):
        from .models import DocumentResult

        return DocumentResult(
            docx_paths=row.state.generated_docx_paths,
            pdf_paths=row.state.generated_pdf_paths,
            extra_attachment_paths=[self.templates.materialize_static_authorisation_pdf(self._require_work_dir())],
            created_at=row.state.attachment_created_at,
            template_version=row.state.template_version,
            attachment_hash=row.state.generated_attachment_hash,
        )

    def _status_from_row(self, row: ConfirmationRow) -> CustomerStatus:
        docs_created = "Passed" if row.state.attachment_created == "Y" and row.state.verification_status == "Passed" else row.state.status
        batch_selected = f"{row.state.batch_id} #{row.state.batch_sequence}" if row.state.batch_id else "Not selected"
        mail_sent = "Sent" if row.state.main_sent == "Y" else ("Failed" if row.state.error else "Not sent")
        return CustomerStatus(
            row_id=row.row_id,
            party_name=row.party_name,
            email=row.email,
            batch_id=row.state.batch_id,
            batch_sequence=row.state.batch_sequence,
            docs_created=docs_created,
            batch_selected=batch_selected,
            mail_sent=mail_sent,
            status=row.state.status,
            error=row.state.error,
        )

    def _handle_error(self, row: ConfirmationRow, step: str, exc: Exception) -> None:
        assert self.store is not None
        info = self.error_classifier.classify(step, exc)
        decision = self.retry_policy.next_attempt(step, row.state.attempt_count + 1, info.retryable)
        if info.retryable:
            self.store.mark_retryable_error(row.row_id, step, info.message, decision.next_retry_at, decision.locked)
        else:
            self.store.mark_error(row.row_id, step, info.message)
        self._log_error(row.row_id, step, info.message, {"category": info.category, "retryable": info.retryable, "locked": decision.locked})

    def _require_loaded(self) -> None:
        if self.store is None or self.master_path is None:
            raise ValidationError("Choose an Excel workbook before running automation.")

    def _require_work_dir(self) -> Path:
        if self.work_dir is None:
            raise ValidationError("Choose an Excel workbook before running automation.")
        return self.work_dir

    def _log_file_text(self) -> Path:
        if self.logger and self.logger.log_file:
            return self.logger.log_file
        return self._require_work_dir() / "logs" / "ui.log"

    def _log_info(self, row_id: str | None, step: str, message: str, details: dict) -> None:
        if self.logger:
            self.logger.info(row_id, step, message, details)

    def _log_error(self, row_id: str | None, step: str, message: str, details: dict) -> None:
        if self.logger:
            self.logger.error(row_id, step, message, details)
