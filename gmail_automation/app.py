from __future__ import annotations

import argparse
import re
import subprocess
import sys
import time
from pathlib import Path

from .batch import BatchPlanner
from .config import load_config, write_sample_config
from .document_generator import DocumentGenerator
from .errors import ErrorClassifier, ValidationError
from .excel_store import ExcelStateStore
from .gmail_tracker import GmailTracker
from .logging_utils import AppLogger
from .mail_sender import send_with_fallback
from .models import ConfirmationRow, SendConfig
from .retry import RetryPolicy
from .templates import TemplateRepository
from .verifier import DocumentVerifier


class GmailConfirmationApp:
    def __init__(self, master_path: Path, config: SendConfig, mode: str, tracking_path: Path | None = None) -> None:
        self.master_path = master_path
        self.work_dir = master_path.parent
        self.config = config
        self.mode = mode
        self.store = ExcelStateStore(master_path, tracking_path)
        self.templates = TemplateRepository()
        self.logger = AppLogger(self.work_dir)
        self.error_classifier = ErrorClassifier()
        self.retry_policy = RetryPolicy()
        self.verifier = DocumentVerifier()

    def run(self) -> int:
        self.logger.start_run(self.mode)
        self.logger.info(None, "startup", "Starting Gmail confirmation automation", {"mode": self.mode})
        try:
            self._startup_validation()
            if self.mode in {"preview", "generate", "all", "send"}:
                self._generate_and_verify()
            if self.mode in {"send", "all"}:
                self._send_batches()
            if self.mode in {"track", "all"}:
                self._track_replies()
        except Exception as exc:
            info = self.error_classifier.classify("startup", exc)
            self.logger.error(None, "startup", info.message, {"category": info.category})
            print(f"ERROR: {info.message}")
            return 1
        self.logger.info(None, "complete", "Run completed", {})
        print(f"Completed. Log: {self.logger.log_file}")
        return 0

    def _startup_validation(self) -> None:
        errors = self.store.validate()
        template_errors = self.templates.validate_placeholders(
            {
                "S.No.",
                "Party Type",
                "Party Name",
                "Email To(Address)",
                "Balance",
            }
        )
        errors.extend(template_errors)
        if errors:
            raise ValidationError("; ".join(errors))
        self.store.ensure_workbooks()
        self.logger.info(None, "startup_validation", "Startup validation passed", {"template_version": self.templates.version()})

    def _generate_and_verify(self) -> None:
        generator = DocumentGenerator(self.work_dir, convert_to_pdf=self.config.convert_to_pdf)
        rows = self.store.load_rows()
        for row in rows:
            self.store.mark_log_file(row.row_id, self.logger.log_file)
            try:
                self._validate_row(row)
                if self._can_skip_generation(row):
                    self.logger.info(row.row_id, "document_generation", "Skipping existing generated documents", {})
                    continue
                result = generator.generate(row, self.templates)
                verification = self.verifier.verify(row, result)
                if verification.passed:
                    self.store.mark_documents_created(row.row_id, result)
                    self.logger.info(row.row_id, "document_generation", "Documents generated and verified", {"docx": len(result.docx_paths), "pdf": len(result.pdf_paths)})
                else:
                    self.store.mark_documents_created(row.row_id, result, "Failed", "; ".join(verification.errors))
                    self.logger.error(row.row_id, "document_verification", "Generated documents failed verification", {"errors": verification.errors})
            except Exception as exc:
                self._handle_error(row, "document_generation", exc)

    def _send_batches(self) -> None:
        if self.config.send_mode != "send":
            self.logger.warning(None, "mail_send", "Send mode is not enabled in config; skipping actual send", {"send_mode": self.config.send_mode})
            return
        rows = self.store.load_rows()
        mail_template = self.templates.load_mail_template()
        batches = BatchPlanner().plan(rows, self.config)
        sent_count = 0
        for batch in batches:
            for sequence, row in enumerate(batch.rows, start=1):
                self.store.mark_batch(row.row_id, batch.batch_id, sequence)
                try:
                    self._validate_row(row)
                    if row.state.main_sent == "Y":
                        continue
                    documents = self._documents_from_state(row)
                    result = send_with_fallback(self.config, row, mail_template, documents)
                    self.store.mark_send_success(row.row_id, result)
                    self.logger.info(row.row_id, "mail_send", "Email sent", {"to": row.email, "message_id": result.smtp_message_id})
                    sent_count += 1
                    if sent_count >= self.config.daily_send_limit:
                        self.logger.warning(None, "mail_send", "Daily send limit reached", {"limit": self.config.daily_send_limit})
                        return
                    time.sleep(self.config.per_email_delay_seconds)
                except Exception as exc:
                    self._handle_error(row, "mail_send", exc)
            if batch is not batches[-1]:
                time.sleep(self.config.batch_delay_seconds)

    def _track_replies(self) -> None:
        tracker = GmailTracker(self.config)
        rows = [row for row in self.store.load_rows() if row.state.main_sent == "Y"]
        for row in rows:
            try:
                result = tracker.check(row)
                self.store.mark_tracking_update(row.row_id, result)
                self.logger.info(row.row_id, "gmail_tracking", "Tracking updated", {"reply": result.reply_received, "bounce": result.bounce_received})
            except Exception as exc:
                self._handle_error(row, "gmail_tracking", exc)

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

    def _documents_from_state(self, row: ConfirmationRow):
        from .models import DocumentResult

        return DocumentResult(
            docx_paths=row.state.generated_docx_paths,
            pdf_paths=row.state.generated_pdf_paths,
            extra_attachment_paths=[self.templates.materialize_static_authorisation_pdf(self.work_dir)],
            created_at=row.state.attachment_created_at,
            template_version=row.state.template_version,
            attachment_hash=row.state.generated_attachment_hash,
        )

    def _handle_error(self, row: ConfirmationRow, step: str, exc: Exception) -> None:
        info = self.error_classifier.classify(step, exc)
        decision = self.retry_policy.next_attempt(step, row.state.attempt_count + 1, info.retryable)
        if info.retryable:
            self.store.mark_retryable_error(row.row_id, step, info.message, decision.next_retry_at, decision.locked)
        else:
            self.store.mark_error(row.row_id, step, info.message)
        self.logger.error(row.row_id, step, info.message, {"category": info.category, "retryable": info.retryable, "locked": decision.locked})


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Gmail balance confirmation automation")
    parser.add_argument("--master", default="Information for External Balance Confirmations (1).xlsx", help="Path to master Excel workbook")
    parser.add_argument("--tracking", default="", help="Optional tracking workbook path")
    parser.add_argument("--config", default="config.json", help="Path to config JSON")
    parser.add_argument("--mode", choices=["validate", "preview", "generate", "send", "track", "all"], default="preview")
    parser.add_argument("--no-pdf", action="store_true", help="Generate DOCX only and skip PDF conversion")
    parser.add_argument("--init-config", action="store_true", help="Write a sample config.json and exit")
    parser.add_argument("--ui", action="store_true", help="Launch the desktop batch approval UI")
    parser.add_argument("--modern-ui", action="store_true", help="Launch the modern DB-backed compliance UI")
    parser.add_argument("--console", action="store_true", help="Relaunch in console mode for CLI output")
    parser.add_argument("--init-db", action="store_true", help="Initialize the local SQLite compliance database and exit")
    return parser


def _relaunch_console_mode(current_args: list[str]) -> int:
    if not getattr(sys, "frozen", False):
        print("Console mode handoff is only available from the packaged executable.")
        return 1
    current_exe = Path(sys.executable).resolve()
    if current_exe.stem.endswith("Console"):
        return 0
    console_exe = current_exe.with_name(f"{current_exe.stem}Console.exe")
    if not console_exe.exists():
        print(f"Console executable not found: {console_exe}")
        return 1
    forward_args = [arg for arg in current_args if arg != "--console"]
    subprocess.Popen([str(console_exe), *forward_args], cwd=str(current_exe.parent))
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    parsed_argv = list(argv) if argv is not None else None
    args = parser.parse_args(parsed_argv)
    current_args = parsed_argv if parsed_argv is not None else list(sys.argv[1:])
    launched_without_args = (len(parsed_argv) == 0) if parsed_argv is not None else (len(sys.argv) <= 1)
    if args.console:
        return _relaunch_console_mode(current_args)
    config_path = Path(args.config).resolve()
    if args.init_db:
        from .db import default_database_path, init_db

        db_path = default_database_path()
        init_db(db_path)
        print(f"Initialized database: {db_path}")
        return 0
    if args.modern_ui or launched_without_args:
        from .modern_ui import launch_modern_ui

        return launch_modern_ui()
    if args.ui:
        from .ui import launch_ui

        return launch_ui(config_path)
    if args.init_config:
        write_sample_config(config_path)
        print(f"Wrote sample config: {config_path}")
        return 0
    config = load_config(config_path if config_path.exists() else None)
    if args.no_pdf:
        config.convert_to_pdf = False
    config.send_mode = "send" if args.mode == "send" else config.send_mode
    master_path = Path(args.master).resolve()
    tracking_path = Path(args.tracking).resolve() if args.tracking else None
    return GmailConfirmationApp(master_path, config, args.mode, tracking_path).run()
