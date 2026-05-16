from __future__ import annotations

import threading
from pathlib import Path

from sqlmodel import Session, select

from .config import load_config
from .dao import ClientDAO
from .db import default_database_path, init_db
from .db_models import Client, ClientQuarter, Counterparty, DocumentJob, EmailMessage, ExcelImport, GeneratedDocument, Template, TemplateVariable
from .docx_utils import extract_docx_text
from .errors import ValidationError
from .mail_sender import create_mail_sender
from .models import SendConfig
from .services import ClientService, ComplianceService, DashboardService, ImportService, STATUS_COLUMNS, TemplateService, WorkflowService
from .settings_store import SettingsService


class ModernComplianceController:
    def __init__(self, db_path: Path | None = None) -> None:
        self.db_path = db_path or default_database_path()
        self.engine = init_db(self.db_path)

    def close(self) -> None:
        self.engine.dispose()

    def start_background_compliance_reconcile(self) -> None:
        def worker() -> None:
            try:
                with Session(self.engine) as session:
                    ComplianceService(session).reconcile_all()
            except Exception:
                # Reconciliation also runs during dashboard reads and workflow updates.
                pass

        threading.Thread(target=worker, daemon=True).start()

    def snapshot(self, selected_quarter_id: int | None = None) -> dict:
        with Session(self.engine) as session:
            clients = ClientDAO(session).list_clients()
            quarters = ClientDAO(session).list_quarters()
            current_quarter = _latest_quarter(quarters)
            selected_quarter = session.get(ClientQuarter, selected_quarter_id) if selected_quarter_id else current_quarter
            summary = DashboardService(session).compliance_summary(selected_quarter.id) if selected_quarter and selected_quarter.id else None
            client_cards = _client_cards(clients, quarters, DashboardService(session))
            imports = []
            if selected_quarter and selected_quarter.id:
                imports = list(session.exec(select(ExcelImport).where(ExcelImport.client_quarter_id == selected_quarter.id)))
            return {
                "clients": clients,
                "quarters": quarters,
                "client_cards": client_cards,
                "current_quarter": current_quarter,
                "selected_quarter": selected_quarter,
                "summary": summary,
                "imports": imports,
                "db_path": self.db_path,
            }

    def current_quarter_id_for_client(self, client_id: int) -> int:
        with Session(self.engine) as session:
            quarters = ClientDAO(session).list_quarters(client_id)
            current_quarter = _latest_quarter(quarters)
            if current_quarter is None or current_quarter.id is None:
                raise ValueError("Create a quarter for this client first.")
            return current_quarter.id

    def client_cards(self) -> list[dict]:
        with Session(self.engine) as session:
            return _client_cards(ClientDAO(session).list_clients(), ClientDAO(session).list_quarters(), DashboardService(session))

    def client_quarters(self, client_id: int) -> list[ClientQuarter]:
        with Session(self.engine) as session:
            return ClientDAO(session).list_quarters(client_id)

    def create_client_record(self, name: str, client_type: str = "listed_org") -> dict:
        if not (name or "").strip():
            raise ValueError("Enter a client name.")
        with Session(self.engine) as session:
            client = ClientService(session).create_client(name=name.strip(), client_type=(client_type or "").strip() or "listed_org")
            return {"id": client.id, "name": client.name, "message": f"Created client: {client.name}."}

    def create_client_and_quarter(self, name: str, financial_year: str, quarter: str, client_type: str = "listed_org") -> None:
        with Session(self.engine) as session:
            service = ClientService(session)
            client = service.create_client(name=(name or "").strip(), client_type=(client_type or "").strip() or "listed_org")
            assert client.id is not None
            service.create_quarter(client.id, (financial_year or "").strip(), (quarter or "").strip(), current=True)

    def create_client(self, name: str, client_type: str = "listed_org") -> str:
        return str(self.create_client_record(name, client_type)["message"])

    def create_quarter(self, client_id: int, financial_year: str, quarter: str, current: bool = True) -> str:
        if not (financial_year or "").strip():
            raise ValueError("Enter a financial year.")
        with Session(self.engine) as session:
            created = ClientService(session).create_quarter(client_id, financial_year.strip(), (quarter or "").strip(), current=current)
            return f"Created quarter: {created.financial_year} {created.quarter}."

    def import_workbook(self, client_quarter_id: int, workbook_path: str, client_id: int | None = None) -> str:
        if not (workbook_path or "").strip():
            raise ValueError("Enter the Excel workbook path.")
        path = _normalize_user_path(workbook_path)
        if not path.exists():
            raise ValueError(f"Workbook not found: {path}")
        if not path.is_file():
            raise ValueError(f"Not a file: {path}")
        with Session(self.engine) as session:
            client_quarter = session.get(ClientQuarter, client_quarter_id)
            if client_quarter is None:
                raise ValueError("Select a quarter before importing.")
            if client_id is not None and client_quarter.client_id != client_id:
                raise ValueError("Selected quarter does not belong to the client you are configuring.")
            try:
                summary = ImportService(session).import_excel(client_quarter.client_id, client_quarter_id, path, replace_existing=True)
            except ValidationError as exc:
                raise ValueError(str(exc)) from exc
            legacy_status_columns = [column for column in summary.detected_columns if column in STATUS_COLUMNS]
            status_note = (
                f" Legacy status columns wired: {', '.join(legacy_status_columns)}."
                if legacy_status_columns
                else ""
            )
            return f"Imported {summary.rows_imported} row(s) from {len(summary.sheet_counts)} sheet(s).{status_note}"

    def save_mail_templates(self, client_quarter_id: int, subject: str, body: str) -> str:
        with Session(self.engine) as session:
            if session.get(ClientQuarter, client_quarter_id) is None:
                raise ValueError("Select a quarter before saving mail templates.")
            service = TemplateService(session)
            service.save_text_template(client_quarter_id, "mail_subject", "Mail Subject", subject or "")
            service.save_text_template(client_quarter_id, "mail_body", "Mail Body", body or "")
            required = service.required_variables(client_quarter_id)
            return "Variables detected: " + (", ".join(sorted(required)) or "none")

    def preview_mail_templates(self, client_quarter_id: int) -> dict[str, str]:
        with Session(self.engine) as session:
            if session.get(ClientQuarter, client_quarter_id) is None:
                raise ValueError("Select a quarter before previewing mail templates.")
            templates = list(
                session.exec(
                    select(Template).where(
                        Template.client_quarter_id == client_quarter_id,
                        Template.template_type.in_({"mail_subject", "mail_body"}),
                        Template.is_active == True,  # noqa: E712
                    ).order_by(Template.created_at.desc(), Template.id.desc())
                )
            )
            latest_by_type: dict[str, Template] = {}
            for template in templates:
                if template.template_type not in latest_by_type:
                    latest_by_type[template.template_type] = template
            subject = latest_by_type.get("mail_subject")
            body = latest_by_type.get("mail_body")
            return {
                "subject": subject.content_text if subject else "",
                "body": body.content_text if body else "",
            }

    def save_document_templates(self, client_quarter_id: int, template_paths: str) -> str:
        paths = _parse_template_paths(template_paths)
        if not paths:
            raise ValueError("Enter at least one document template path.")
        with Session(self.engine) as session:
            client_quarter = session.get(ClientQuarter, client_quarter_id)
            if client_quarter is None:
                raise ValueError("Select a quarter before saving document templates.")
            service = TemplateService(session)
            templates = service.save_document_templates(
                client_quarter_id=client_quarter_id,
                file_paths=paths,
                storage_dir=self.db_path.parent / "templates" / str(client_quarter_id) / "documents",
                client_id=client_quarter.client_id,
            )
            required = service.required_variables(client_quarter_id)
            names = ", ".join(template.name for template in templates)
            variables = ", ".join(sorted(required)) or "none"
            columns = set()
            imports = list(session.exec(select(ExcelImport).where(ExcelImport.client_quarter_id == client_quarter_id)))
            if imports:
                columns = set(imports[-1].detected_columns)
            mapping_errors = service.validate_mappings(client_quarter_id, columns) if required else []
            if mapping_errors:
                return (
                    f"Saved {len(templates)} document template(s): {names}. "
                    f"Variables detected: {variables}. Validation issues: {'; '.join(mapping_errors)}"
                )
            return f"Saved {len(templates)} document template(s): {names}. Configured and validated. Variables detected: {variables}"

    def add_document_templates(self, client_quarter_id: int, template_paths: str) -> str:
        paths = _parse_template_paths(template_paths)
        if not paths:
            raise ValueError("Enter at least one document template path.")
        with Session(self.engine) as session:
            client_quarter = session.get(ClientQuarter, client_quarter_id)
            if client_quarter is None:
                raise ValueError("Select a quarter before adding document templates.")
            service = TemplateService(session)
            templates = [
                service.save_file_template(
                    client_quarter_id=client_quarter_id,
                    template_type="document",
                    name=path.stem,
                    file_path=path,
                    storage_dir=self.db_path.parent / "templates" / str(client_quarter_id) / "documents",
                    client_id=client_quarter.client_id,
                )
                for path in paths
            ]
            names = ", ".join(template.name for template in templates)
            required = service.required_variables(client_quarter_id)
            variables = ", ".join(sorted(required)) or "none"
            return f"Added {len(templates)} document template(s): {names}. Variables detected: {variables}"

    def list_document_templates(self, client_quarter_id: int) -> list[dict]:
        with Session(self.engine) as session:
            templates = list(
                session.exec(
                    select(Template).where(
                        Template.client_quarter_id == client_quarter_id,
                        Template.template_type == "document",
                        Template.is_active == True,  # noqa: E712
                    )
                )
            )
            rows = []
            for template in templates:
                variables = {
                    row.variable_name
                    for row in session.exec(select(TemplateVariable).where(TemplateVariable.template_id == template.id))
                }
                rows.append(
                    {
                        "id": template.id,
                        "name": template.name,
                        "file_path": template.file_path,
                        "variables": sorted(variables),
                    }
                )
            return rows

    def preview_document_template(self, template_id: int) -> str:
        with Session(self.engine) as session:
            template = session.get(Template, template_id)
            if template is None or template.template_type != "document":
                raise ValueError("Select a configured document template first.")
            if not template.file_path:
                return (template.content_text or "")[:2000]
            path = Path(template.file_path)
            if not path.exists():
                raise ValueError(f"Configured template file is missing: {path}")
            suffix = path.suffix.lower()
            if suffix == ".docx":
                return extract_docx_text(path)[:2000]
            if suffix in {".txt", ".html", ".htm", ".md"}:
                return path.read_text(encoding="utf-8", errors="ignore")[:2000]
            if suffix == ".pdf":
                content = path.read_bytes().decode("latin-1", errors="ignore")
                return content[:2000]
            return f"Preview is not supported for {suffix} files."

    def update_document_template(self, template_id: int, template_path: str) -> str:
        new_path = Path((template_path or "").strip().strip('"'))
        if not new_path.exists():
            raise ValueError("Choose an existing replacement document template file.")
        with Session(self.engine) as session:
            existing = session.get(Template, template_id)
            if existing is None or existing.template_type != "document" or existing.client_quarter_id is None:
                raise ValueError("Select a configured document template first.")
            client_quarter = session.get(ClientQuarter, existing.client_quarter_id)
            if client_quarter is None:
                raise ValueError("Configured document is missing its client quarter.")
            existing.is_active = False
            session.add(existing)
            session.commit()
            service = TemplateService(session)
            updated = service.save_file_template(
                client_quarter_id=existing.client_quarter_id,
                template_type="document",
                name=new_path.stem,
                file_path=new_path,
                storage_dir=self.db_path.parent / "templates" / str(existing.client_quarter_id) / "documents",
                client_id=client_quarter.client_id,
            )
            return f"Updated document template: {updated.name}."

    def delete_document_template(self, template_id: int) -> str:
        with Session(self.engine) as session:
            template = session.get(Template, template_id)
            if template is None or template.template_type != "document":
                raise ValueError("Select a configured document template first.")
            template.is_active = False
            session.add(template)
            session.commit()
            return f"Deleted document template: {template.name}."

    def reset_quarter_after_config_change(self, client_quarter_id: int, reset_documents: bool = False, reset_mail: bool = False) -> str:
        with Session(self.engine) as session:
            changed = ComplianceService(session).reset_quarter_status(
                client_quarter_id,
                reset_documents=reset_documents,
                reset_mail=reset_mail,
            )
            auto_mapped = self._autofill_missing_template_mappings(session, client_quarter_id)
            mapping_note = f" Auto-mapped {auto_mapped} missing variable(s)." if auto_mapped else ""
            if reset_documents:
                return (
                    f"Reset {changed} counterparty row(s): compliance is non-compliant and generated docs/mail state was cleared."
                    f"{mapping_note}"
                )
            if reset_mail:
                return f"Reset {changed} counterparty row(s): compliance is non-compliant and mail state was cleared.{mapping_note}"
            return "No reset requested."

    def _autofill_missing_template_mappings(self, session: Session, client_quarter_id: int) -> int:
        template_service = TemplateService(session)
        imports = list(session.exec(select(ExcelImport).where(ExcelImport.client_quarter_id == client_quarter_id)))
        columns = set(imports[-1].detected_columns if imports else [])
        required_variables = template_service.required_variables(client_quarter_id)
        existing_mappings = template_service.templates.mappings_for_quarter(client_quarter_id)
        merged_mappings = {
            variable: (mapping.source_type, mapping.source_key, mapping.constant_value)
            for variable, mapping in existing_mappings.items()
        }
        added = 0
        for variable in sorted(required_variables):
            if variable in existing_mappings:
                continue
            candidate = _best_column_match(variable, columns)
            if candidate:
                merged_mappings[variable] = ("excel_column", candidate, "")
                added += 1
                continue
            if variable.startswith("row."):
                merged_mappings[variable] = ("counterparty_field", variable.split(".", 1)[1], "")
                added += 1
        if added:
            template_service.save_mappings(client_quarter_id, merged_mappings)
        return added

    def auto_map_variables(self, client_quarter_id: int) -> str:
        with Session(self.engine) as session:
            if session.get(ClientQuarter, client_quarter_id) is None:
                raise ValueError("Select a quarter before mapping variables.")
            service = TemplateService(session)
            imports = list(session.exec(select(ExcelImport).where(ExcelImport.client_quarter_id == client_quarter_id)))
            columns = set(imports[-1].detected_columns if imports else [])
            mappings = {}
            for variable in service.required_variables(client_quarter_id):
                candidate = _best_column_match(variable, columns)
                if candidate:
                    mappings[variable] = ("excel_column", candidate, "")
                elif variable.startswith("row."):
                    mappings[variable] = ("counterparty_field", variable.split(".", 1)[1], "")
            service.save_mappings(client_quarter_id, mappings)
            errors = service.validate_mappings(client_quarter_id, columns)
            return "Mappings valid." if not errors else "Mapping issues: " + "; ".join(errors)

    def quarter_configuration_status(self, client_quarter_id: int) -> dict:
        with Session(self.engine) as session:
            if session.get(ClientQuarter, client_quarter_id) is None:
                raise ValueError("Select a quarter before checking configuration.")
            service = TemplateService(session)
            imports = list(session.exec(select(ExcelImport).where(ExcelImport.client_quarter_id == client_quarter_id)))
            columns = set(imports[-1].detected_columns if imports else [])
            variables = service.required_variables(client_quarter_id)
            return {
                "import_count": len(imports),
                "latest_columns": sorted(columns),
                "variables": sorted(variables),
                "mapping_errors": service.validate_mappings(client_quarter_id, columns) if variables else [],
            }

    def quarter_workflow_readiness(self, client_quarter_id: int) -> dict:
        with Session(self.engine) as session:
            if session.get(ClientQuarter, client_quarter_id) is None:
                raise ValueError("Select a quarter before checking workflow readiness.")
            counterparties = [
                counterparty
                for counterparty in session.exec(select(Counterparty).where(Counterparty.client_quarter_id == client_quarter_id))
                if counterparty.party_name or counterparty.email or counterparty.balance
            ]
            imports = list(session.exec(select(ExcelImport).where(ExcelImport.client_quarter_id == client_quarter_id)))
            templates = list(
                session.exec(select(Template).where(Template.client_quarter_id == client_quarter_id, Template.is_active == True))  # noqa: E712
            )
            document_templates = [template for template in templates if template.template_type == "document"]
            mail_templates = [template for template in templates if template.template_type in {"mail_subject", "mail_body"}]
            columns = set(imports[-1].detected_columns if imports else [])
            template_service = TemplateService(session)
            template_variables: dict[int, set[str]] = {
                template.id: {
                    row.variable_name
                    for row in session.exec(select(TemplateVariable).where(TemplateVariable.template_id == template.id))
                }
                for template in templates
                if template.id is not None
            }
            document_variables = sorted(
                {
                    variable
                    for template in document_templates
                    for variable in template_variables.get(template.id, set())
                }
            )
            mail_variables = sorted(
                {
                    variable
                    for template in mail_templates
                    for variable in template_variables.get(template.id, set())
                }
            )
            mappings = template_service.templates.mappings_for_quarter(client_quarter_id)
            document_mapping_errors = _mapping_errors_for_variables(document_variables, mappings, columns)
            mail_mapping_errors = _mapping_errors_for_variables(mail_variables, mappings, columns)
            send_config = self._load_send_config()
            smtp_ready = (
                send_config.mail_provider in {"gmail_smtp", "webtel_smtp"}
                and bool(send_config.sender_email)
                and bool(send_config.smtp_password or send_config.app_password)
            )
            send_enabled = send_config.send_mode == "send" and smtp_ready
            return {
                "counterparty_count": len(counterparties),
                "import_count": len(imports),
                "document_template_count": len(document_templates),
                "mail_template_count": len(mail_templates),
                "document_variables": document_variables,
                "mail_variables": mail_variables,
                "document_mapping_errors": document_mapping_errors,
                "mail_mapping_errors": mail_mapping_errors,
                "can_generate_documents": bool(counterparties) and bool(document_templates) and not document_mapping_errors,
                "can_send_mail": bool(counterparties) and len(mail_templates) >= 2 and not mail_mapping_errors and send_enabled,
                "send_mode": send_config.send_mode,
                "send_enabled": send_enabled,
                "mail_provider": send_config.mail_provider,
                "connected_email": send_config.sender_email,
            }

    def get_client_compliance_summary(self, client_id: int) -> dict:
        with Session(self.engine) as session:
            dao = ClientDAO(session)
            client = dao.get_client(client_id)
            quarters = dao.list_quarters(client_id)
            current_quarter = _latest_quarter(quarters)
            summary = DashboardService(session).compliance_summary(current_quarter.id) if current_quarter and current_quarter.id else None
            imports = []
            if current_quarter and current_quarter.id:
                imports = list(session.exec(select(ExcelImport).where(ExcelImport.client_quarter_id == current_quarter.id)))
            return {
                "client_id": client.id,
                "client_name": client.name,
                "quarter_id": current_quarter.id if current_quarter else None,
                "quarter_label": f"{current_quarter.financial_year} {current_quarter.quarter}" if current_quarter else "No quarter",
                "summary": summary,
                "imports": imports,
            }

    def quarter_counterparty_statuses(self, client_quarter_id: int) -> list[dict]:
        with Session(self.engine) as session:
            if session.get(ClientQuarter, client_quarter_id) is None:
                raise ValueError("Select a quarter before viewing row-level compliance.")
            ComplianceService(session).reconcile_quarter(client_quarter_id)
            counterparties = list(
                session.exec(
                    select(Counterparty)
                    .where(Counterparty.client_quarter_id == client_quarter_id)
                    .order_by(Counterparty.source_sheet, Counterparty.source_row_number)
                )
            )
            jobs_by_counterparty: dict[int, list[DocumentJob]] = {}
            for job in session.exec(select(DocumentJob).where(DocumentJob.client_quarter_id == client_quarter_id)):
                jobs_by_counterparty.setdefault(job.counterparty_id, []).append(job)
            quarter_job_ids = {
                job.id
                for jobs in jobs_by_counterparty.values()
                for job in jobs
                if job.id is not None
            }
            documents_by_counterparty: dict[int, list[GeneratedDocument]] = {}
            if quarter_job_ids:
                for document in session.exec(select(GeneratedDocument).where(GeneratedDocument.document_job_id.in_(quarter_job_ids))):
                    documents_by_counterparty.setdefault(document.counterparty_id, []).append(document)
            emails_by_counterparty: dict[int, list[EmailMessage]] = {}
            for message in session.exec(select(EmailMessage).where(EmailMessage.client_quarter_id == client_quarter_id)):
                emails_by_counterparty.setdefault(message.counterparty_id, []).append(message)

            rows = []
            for counterparty in counterparties:
                assert counterparty.id is not None
                documents = documents_by_counterparty.get(counterparty.id, [])
                jobs = jobs_by_counterparty.get(counterparty.id, [])
                messages = emails_by_counterparty.get(counterparty.id, [])
                latest_message = max(messages, key=lambda item: item.updated_at or item.created_at, default=None)
                rows.append(
                    {
                        "row": counterparty.source_row_number,
                        "counterparty_id": counterparty.id,
                        "sheet": counterparty.source_sheet,
                        "party_type": counterparty.party_type or "Unspecified",
                        "party_name": counterparty.party_name,
                        "email": counterparty.email,
                        "balance": counterparty.balance,
                        "compliance_status": counterparty.status,
                        "document_status": _document_status(jobs, documents),
                        "document_count": len(documents),
                        "documents": [_generated_document_summary(document) for document in documents if document.id is not None],
                        "mail_status": latest_message.status if latest_message else "not generated",
                        "mail_sent": "Yes" if latest_message and latest_message.status == "sent" else "No",
                    }
                )
            return rows

    def run_generation(self, client_quarter_id: int) -> str:
        with Session(self.engine) as session:
            workflow = WorkflowService(session, output_root=self.db_path.parent / "generated")
            workflow.enqueue_document_generation(client_quarter_id)
            results = workflow.generate_pending_documents(client_quarter_id)
            generated = sum(1 for result in results if result.status == "generated")
            failed = len(results) - generated
            return f"Generated {generated} document job(s); {failed} need attention."

    def regenerate_documents(self, client_quarter_id: int, counterparty_ids: set[int]) -> str:
        if not counterparty_ids:
            raise ValueError("Select at least one Excel row before regenerating documents.")
        readiness = self.quarter_workflow_readiness(client_quarter_id)
        if not readiness["document_template_count"]:
            raise ValueError("Save document templates before regenerating documents.")
        if readiness["document_mapping_errors"]:
            raise ValueError("Fix document mappings before regenerating documents: " + "; ".join(readiness["document_mapping_errors"]))
        with Session(self.engine) as session:
            workflow = WorkflowService(session, output_root=self.db_path.parent / "generated")
            results = workflow.regenerate_documents(client_quarter_id, counterparty_ids)
            generated = sum(1 for result in results if result.status == "generated")
            failed = len(results) - generated
            return f"Regenerated {generated} selected document job(s); {failed} need attention."

    def _saved_mail_templates_for_quarter(self, client_quarter_id: int) -> tuple[str, str]:
        preview = self.preview_mail_templates(client_quarter_id)
        subject = (preview.get("subject") or "").strip()
        body = preview.get("body") or ""
        if not subject or not body.strip():
            raise ValueError("Save mail subject and body for this quarter before queueing or sending.")
        return subject, body

    def queue_and_preview_send(self, client_quarter_id: int) -> str:
        subject, body = self._saved_mail_templates_for_quarter(client_quarter_id)
        with Session(self.engine) as session:
            workflow = WorkflowService(session, output_root=self.db_path.parent / "generated")
            messages = workflow.queue_email_messages(client_quarter_id, subject, body)
            sent = workflow.mark_preview_batch_sent(client_quarter_id)
            return f"Queued {len(messages)} email(s); marked {sent} sent in preview mode."

    def queue_and_preview_send_selected(self, client_quarter_id: int, counterparty_ids: set[int]) -> str:
        if not counterparty_ids:
            raise ValueError("Select at least one Excel row before queueing mail.")
        subject, body = self._saved_mail_templates_for_quarter(client_quarter_id)
        readiness = self.quarter_workflow_readiness(client_quarter_id)
        if readiness["mail_template_count"] < 2:
            raise ValueError("Save mail templates before queueing mail.")
        if readiness["mail_mapping_errors"]:
            raise ValueError("Fix mail mappings before queueing mail: " + "; ".join(readiness["mail_mapping_errors"]))
        generated_ids = self._counterparty_ids_with_generated_documents(client_quarter_id)
        missing_docs = counterparty_ids - generated_ids
        if missing_docs:
            raise ValueError("Generate documents for all selected rows before queueing mail.")
        with Session(self.engine) as session:
            workflow = WorkflowService(session, output_root=self.db_path.parent / "generated")
            messages = workflow.queue_email_messages(client_quarter_id, subject, body, counterparty_ids=counterparty_ids)
            sent = workflow.mark_preview_batch_sent(client_quarter_id, counterparty_ids=counterparty_ids)
            return f"Queued {len(messages)} selected email(s); marked {sent} sent in preview mode."

    def send_mail_selected(self, client_quarter_id: int, counterparty_ids: set[int]) -> str:
        if not counterparty_ids:
            raise ValueError("Select at least one Excel row before sending mail.")
        subject, body = self._saved_mail_templates_for_quarter(client_quarter_id)
        readiness = self.quarter_workflow_readiness(client_quarter_id)
        if readiness["mail_template_count"] < 2:
            raise ValueError("Save mail templates before sending mail.")
        if readiness["mail_mapping_errors"]:
            raise ValueError("Fix mail mappings before sending: " + "; ".join(readiness["mail_mapping_errors"]))
        if not readiness["send_enabled"]:
            raise ValueError("Enable send mode and configure Gmail SMTP or Webtel SMTP before sending live mail.")
        generated_ids = self._counterparty_ids_with_generated_documents(client_quarter_id)
        missing_docs = counterparty_ids - generated_ids
        if missing_docs:
            raise ValueError("Generate documents for all selected rows before sending mail.")
        send_config = self._load_send_config()
        with Session(self.engine) as session:
            workflow = WorkflowService(session, config=send_config, output_root=self.db_path.parent / "generated")
            messages = workflow.queue_email_messages(client_quarter_id, subject, body, counterparty_ids=counterparty_ids)
            sent = workflow.send_queued_emails(client_quarter_id, preview=False, counterparty_ids=counterparty_ids)
            failed = len(messages) - sent
            return f"Queued {len(messages)} selected email(s); sent {sent}, failed {failed}."

    def get_mail_settings(self) -> dict:
        with Session(self.engine) as session:
            settings = SettingsService(session)
            provider = settings.get_value("mail.provider", "")
            workspace_config = Path.cwd() / "config.json"
            local_config = self.db_path.parent / "config.json"
            fallback = load_config(workspace_config if workspace_config.exists() else (local_config if local_config.exists() else None))
            config = self._load_send_config()
            effective_provider = normalized_mail_provider_value(provider or config.mail_provider or fallback.mail_provider)
            return {
                "mail_provider": effective_provider,
                "fallback_providers": config.fallback_providers or fallback.fallback_providers,
                "send_mode": config.send_mode,
                "sender_email": config.sender_email,
                "daily_send_limit": str(config.daily_send_limit),
                "per_email_delay_seconds": str(config.per_email_delay_seconds),
                "smtp_username": config.smtp_username or config.sender_email,
                "smtp_password_saved": bool(config.smtp_password or config.app_password),
                "connected_email": config.sender_email,
            }

    def save_mail_settings(
        self,
        mail_provider: str,
        send_mode: str,
        sender_email: str,
        fallback_providers: str,
        daily_send_limit: int,
        per_email_delay_seconds: int,
        smtp_host: str,
        smtp_port: int,
        smtp_username: str,
        smtp_password: str = "",
    ) -> str:
        provider = normalized_mail_provider_value(mail_provider)
        mode = (send_mode or "").strip() or "preview"
        with Session(self.engine) as session:
            settings = SettingsService(session)
            settings.set_value("mail.provider", provider)
            settings.set_value("mail.fallback_providers", (fallback_providers or "").strip())
            settings.set_value("mail.send_mode", mode)
            settings.set_value("mail.sender_email", (sender_email or "").strip())
            settings.set_value("mail.daily_send_limit", str(max(1, int(daily_send_limit))))
            settings.set_value("mail.per_email_delay_seconds", str(max(0, int(per_email_delay_seconds))))
            settings.set_value("mail.smtp_host", (smtp_host or "smtp.gmail.com").strip())
            settings.set_value("mail.smtp_port", str(max(1, int(smtp_port))))
            settings.set_value("mail.smtp_use_ssl", "true" if provider == "webtel_smtp" or int(smtp_port) == 465 else "false")
            settings.set_value("mail.smtp_username", (smtp_username or sender_email or "").strip())
            if smtp_password.strip():
                _save_smtp_password(settings, provider, sender_email, smtp_password.strip())
        return "Mail settings saved."

    def test_mail_connection(self) -> str:
        config = self._load_send_config()
        if config.mail_provider in {"gmail_smtp", "webtel_smtp"} and not config.smtp_username:
            config.smtp_username = config.sender_email
        sender = create_mail_sender(config)
        return sender.test_connection()

    def send_test_email(self, to_email: str) -> str:
        config = self._load_send_config()
        if config.mail_provider in {"gmail_smtp", "webtel_smtp"} and not config.smtp_username:
            config.smtp_username = config.sender_email
        if not to_email.strip():
            raise ValueError("Enter a test recipient email.")
        dummy_row = Counterparty(
            id=0,
            client_id=0,
            client_quarter_id=0,
            source_sheet="test",
            source_row_number=0,
            source_row_key="test",
            party_type="Test",
            party_name="Test Party",
            email=to_email.strip(),
            cc="",
            balance="0",
        )
        row_model = _counterparty_to_confirmation_row(dummy_row, subject="Mail configuration test", body="This is a test email from configuration.")
        from .models import DocumentResult, MailTemplate

        result = create_mail_sender(config).send(
            row_model,
            MailTemplate(subject_template=row_model.subject, body_template_text=row_model.mail_body_override, required_fields=set()),
            DocumentResult(docx_paths=[], pdf_paths=[], extra_attachment_paths=[], created_at="", template_version="test", attachment_hash=""),
        )
        return f"Test email sent to {to_email.strip()} (message id: {result.smtp_message_id})."

    def _load_send_config(self) -> SendConfig:
        workspace_config = Path.cwd() / "config.json"
        local_config = self.db_path.parent / "config.json"
        config = load_config(workspace_config if workspace_config.exists() else (local_config if local_config.exists() else None))
        with Session(self.engine) as session:
            settings = SettingsService(session)
            config.mail_provider = normalized_mail_provider_value(settings.get_value("mail.provider", config.mail_provider))
            config.fallback_providers = settings.get_value("mail.fallback_providers", config.fallback_providers)
            config.send_mode = settings.get_value("mail.send_mode", config.send_mode)
            config.sender_email = settings.get_value("mail.sender_email", config.sender_email)
            config.daily_send_limit = int(settings.get_value("mail.daily_send_limit", str(config.daily_send_limit)) or config.daily_send_limit)
            config.per_email_delay_seconds = int(
                settings.get_value("mail.per_email_delay_seconds", str(config.per_email_delay_seconds)) or config.per_email_delay_seconds
            )
            if config.mail_provider == "gmail_smtp":
                config.smtp_host = "smtp.gmail.com"
                config.smtp_port = 587
                config.smtp_use_starttls = True
                config.smtp_use_ssl = False
            elif config.mail_provider == "webtel_smtp":
                config.smtp_host = "connect.webtelconnect.com"
                config.smtp_port = 465
                config.smtp_use_starttls = False
                config.smtp_use_ssl = True
            config.smtp_username = settings.get_value("mail.smtp_username", config.smtp_username or config.sender_email)
            config.smtp_password = _load_smtp_password(settings, config.mail_provider, config.sender_email, config.smtp_password or config.app_password)
        return config

    def _counterparty_ids_with_generated_documents(self, client_quarter_id: int) -> set[int]:
        with Session(self.engine) as session:
            rows = session.exec(
                select(DocumentJob).where(
                    DocumentJob.client_quarter_id == client_quarter_id,
                    DocumentJob.status == "generated",
                )
            )
            return {row.counterparty_id for row in rows}

    def generated_documents_for_counterparty(self, client_quarter_id: int, counterparty_id: int) -> list[dict]:
        with Session(self.engine) as session:
            counterparty = session.get(Counterparty, counterparty_id)
            if counterparty is None or counterparty.client_quarter_id != client_quarter_id:
                raise ValueError("Selected row does not belong to the active quarter.")
            jobs = list(
                session.exec(select(DocumentJob).where(DocumentJob.client_quarter_id == client_quarter_id, DocumentJob.counterparty_id == counterparty_id))
            )
            job_ids = [job.id for job in jobs if job.id is not None]
            if not job_ids:
                return []
            generated_documents = list(
                session.exec(
                    select(GeneratedDocument)
                    .where(GeneratedDocument.document_job_id.in_(job_ids))
                    .order_by(GeneratedDocument.created_at)
                )
            )
            return [
                {
                    "id": document.id,
                    "file_path": document.file_path,
                    "file_type": document.file_type,
                    "created_at": document.created_at,
                }
                for document in generated_documents
                if document.id is not None
            ]

    def preview_generated_document(self, generated_document_id: int) -> str:
        with Session(self.engine) as session:
            document = session.get(GeneratedDocument, generated_document_id)
            if document is None:
                raise ValueError("Select a generated document to preview.")
            path = Path(document.file_path)
            if not path.exists():
                raise ValueError(f"Generated document file is missing: {path}")
            suffix = path.suffix.lower()
            if suffix == ".docx":
                return extract_docx_text(path)[:3000]
            if suffix in {".txt", ".html", ".htm", ".md"}:
                return path.read_text(encoding="utf-8", errors="ignore")[:3000]
            if suffix == ".pdf":
                return path.read_bytes().decode("latin-1", errors="ignore")[:3000]
            return f"Preview is not supported for {suffix} files.\nFile: {path}"


def launch_modern_ui(db_path: Path | None = None) -> int:
    try:
        import flet as ft
    except ImportError as exc:
        raise RuntimeError("Flet is required for the modern UI. Install requirements.txt first.") from exc

    controller = ModernComplianceController(db_path)
    controller.start_background_compliance_reconcile()

    def main(page: ft.Page) -> None:
        page.title = "Gmail Compliance Automation"
        page.theme_mode = ft.ThemeMode.LIGHT
        page.window_width = 1280
        page.window_height = 820
        page.padding = 24
        page.scroll = ft.ScrollMode.AUTO
        ui = build_modern_ui_controls(ft, controller, page)
        page.add(*ui["controls"])
        ui["refresh"]()

    ft.app(target=main)
    return 0


def build_modern_ui_controls(ft, controller: ModernComplianceController, page=None) -> dict:
    ui_guard = {"busy": False}
    status = ft.Text("Ready.", selectable=True)
    selected_quarter = ft.Dropdown(label="Selected client quarter", expand=True)
    current_summary = ft.Text("No current quarter yet.", size=16)
    dashboard_cards = ft.Row(wrap=True, spacing=12)
    workflow_counts = ft.Column(spacing=4)

    client_name = ft.TextField(label="New client name", expand=True)
    client_type = ft.Dropdown(label="Client type", value="listed_org", options=[ft.dropdown.Option("listed_org", "Listed organization")])
    financial_year = ft.TextField(label="Financial year", value="2026-27", width=160)
    quarter = ft.Dropdown(label="Quarter", value="Q1", width=120, options=[ft.dropdown.Option(value) for value in ("Q1", "Q2", "Q3", "Q4")])
    configure_client = ft.Dropdown(label="Client", expand=True)
    configure_quarter = ft.Dropdown(label="Client quarter", expand=True)
    clients_list = ft.Column(spacing=4)
    compliance_clients = ft.Column(spacing=8)
    configure_client_id = {"value": None}
    compliance_detail_visible = {"value": False}
    settings_title = ft.Text("Select a client from Clients to configure inputs and mapping.", selectable=True)
    configure_empty = ft.Text("Create a client or use Configure on an existing client to continue.", selectable=True)
    mail_settings_title = ft.Text("Select a client from Clients to configure mail setup.", selectable=True)
    mail_setup_empty = ft.Text("Create/select a client before configuring mail setup.", selectable=True)
    workflow_context = ft.Text("Select a client quarter before running workflow actions.", selectable=True)
    compliance_detail_header = ft.Text("Select a client's current quarter.", size=24, weight=ft.FontWeight.BOLD, selectable=True)
    compliance_detail = ft.Text("Open a client's current quarter to see its compliance details.", selectable=True)
    compliance_actions_allowed = ft.Text("Actions become available after opening a current quarter.", selectable=True)
    compliance_empty_detail = ft.Text("Choose Current Quarter on an organization card to open the detailed row workflow view.", selectable=True)
    compliance_intro = ft.Text("Organizations are shown first. Open Current Quarter to view detailed row actions and statuses.")
    compliance_row_details = ft.Column(spacing=8)
    selected_counterparty_ids: set[int] = set()
    visible_counterparty_rows: list[dict] = []
    row_page_index = {"value": 0}
    row_page_size = ft.Dropdown(
        label="Rows per page",
        value="20",
        width=150,
        options=[ft.dropdown.Option(value) for value in ("10", "20", "50", "100")],
    )
    row_page_status = ft.Text("No rows loaded.", selectable=True)
    workflow_batch_size = ft.TextField(label="Batch size", value="20", width=110)
    workflow_action_hint = ft.Text("Open a current quarter and select rows to enable workflow actions.", selectable=True)
    generated_doc_hint = ft.Text("Select exactly one row with generated docs to preview output.", selectable=True)
    generated_doc_dropdown = ft.Dropdown(label="Generated document", expand=True)
    generated_doc_preview = ft.TextField(label="Generated document preview", multiline=True, min_lines=8, read_only=True, height=360)
    generated_doc_preview_title = ft.Text("Generated Docs Preview", size=18, weight=ft.FontWeight.BOLD, selectable=True)
    generated_docs_rows_panel = ft.Column(spacing=8)
    generated_docs_rows_stack = ft.Stack()
    generated_doc_preview_dismiss = ft.Container(visible=False)
    generated_doc_preview_bubble = ft.Container(visible=False)

    workbook_path = ft.TextField(
        label="Master Excel path",
        expand=True,
        hint_text="Browse to choose a file, or paste a path",
    )
    excel_picker = _attach_file_picker(ft, page)
    document_template_paths = ft.TextField(
        label="Document template / attachment path(s)",
        hint_text="Browse to choose DOCX/TXT templates or static PDF attachments; one path per line",
        multiline=True,
        min_lines=3,
    )
    selected_document_paths: list[str] = []
    document_paths_list = ft.Column(spacing=4)
    configured_document_template = ft.Dropdown(label="Configured document template", expand=True)
    document_validation_status = ft.Text("No document templates configured yet.", selectable=True)
    document_preview = ft.TextField(label="Template preview", multiline=True, min_lines=8, read_only=True)
    configured_documents_list = ft.Column(spacing=8)
    selected_document_preview_title = ft.Text("Configured document", size=18, weight=ft.FontWeight.BOLD, selectable=True)
    selected_document_update_path = ft.TextField(label="Replacement document template path", expand=True)
    mail_subject = ft.TextField(label="Mail subject Liquid template", value="Balance confirmation for {{ party_name }}", expand=True)
    mail_body = ft.TextField(
        label="Mail body Liquid template",
        multiline=True,
        min_lines=6,
        value="Dear {{ party_name }},\n\nPlease confirm the balance of {{ balance }} for {{ quarter.name }}.\n\nRegards",
    )
    mail_preview_subject = ft.TextField(label="Mail subject", expand=True)
    mail_preview_body = ft.TextField(label="Mail body", multiline=True, min_lines=8)
    mail_provider = ft.Dropdown(
        label="Mail provider",
        value="gmail_smtp",
        options=[
            ft.dropdown.Option(key="webtel_smtp", text="Webtel SMTP"),
            ft.dropdown.Option(key="gmail_smtp", text="Gmail SMTP"),
        ],
        visible=False,
    )
    send_mode = ft.Dropdown(label="Send mode", value="preview", options=[ft.dropdown.Option("preview"), ft.dropdown.Option("send")], width=160)
    sender_email_field = ft.TextField(label="Sender email", width=420)
    provider_choice_status = ft.Text("Selected provider: Gmail SMTP", selectable=True)
    fallback_providers_field = ft.TextField(
        label="Fallback providers (comma separated)",
        hint_text="Example: webtel_smtp,gmail_smtp",
        width=720,
    )
    daily_send_limit_field = ft.TextField(label="Daily send limit", value="500", width=140)
    per_email_delay_field = ft.TextField(label="Per-email delay (seconds)", value="3", width=180)
    smtp_username_field = ft.TextField(label="SMTP username", width=320)
    smtp_password_field = ft.TextField(label="SMTP password/app password", password=True, can_reveal_password=True, width=720)
    smtp_defaults_text = ft.Text("SMTP server defaults are selected automatically by the app.", selectable=True)
    test_recipient_field = ft.TextField(label="Test recipient email", width=720)
    mail_setup_next_steps = ft.Text(selectable=True)
    mail_settings_status = ft.Text("Mail setup: not configured.", selectable=True)
    inputs_status = ft.Text("Save templates, import Excel, then auto-map variables.", selectable=True)

    page_cards: list = []
    nav_buttons: list = []
    active_page = {"name": "Home"}
    selected_mail_provider = {"value": "gmail_smtp"}

    def safe_update() -> None:
        if page is not None:
            page.update()

    def show_generated_doc_bubble(title: str = "Generated Docs Preview") -> None:
        generated_doc_preview_title.value = title
        generated_doc_preview_bubble.width = _preview_bubble_width(page)
        generated_doc_preview_dismiss.visible = True
        generated_doc_preview_bubble.visible = True

    def hide_generated_doc_bubble() -> None:
        generated_doc_preview_dismiss.visible = False
        generated_doc_preview_bubble.visible = False

    def show_feedback(message: str, is_error: bool = False) -> None:
        status.value = message
        if page is None:
            return
        snack = ft.SnackBar(
            content=ft.Text(message),
            bgcolor=ft.Colors.RED_100 if is_error else ft.Colors.GREEN_100,
        )
        try:
            page.snack_bar = snack
            page.snack_bar.open = True
            page.update()
        except Exception:
            try:
                page.open(snack)
                page.update()
            except Exception:
                pass

    def selected_quarter_id() -> int:
        if not selected_quarter.value:
            raise ValueError("Create or select a quarter first.")
        return int(selected_quarter.value)

    def configured_client_id() -> int:
        if configure_client_id["value"] is None:
            raise ValueError("Select a client from Clients first.")
        return int(configure_client_id["value"])

    def configured_quarter_id() -> int:
        if not configure_quarter.value:
            raise ValueError("Create or select a quarter for this client first.")
        return int(configure_quarter.value)

    def run_action(action):
        try:
            show_feedback(str(action()))
            refresh()
        except Exception as exc:
            show_feedback(f"Error: {exc}", is_error=True)
            safe_update()

    def prompt_counterparty_reset(action, reset_documents: bool = False, reset_mail: bool = False) -> None:
        def run_with_reset(reset: bool) -> str:
            message = str(action())
            if reset:
                reset_message = controller.reset_quarter_after_config_change(
                    configured_quarter_id(),
                    reset_documents=reset_documents,
                    reset_mail=reset_mail,
                )
                return f"{message} {reset_message}"
            return message

        if page is None:
            run_action(lambda: run_with_reset(False))
            return

        reset_scope = "generated docs, mail state, and compliance status" if reset_documents else "mail state and compliance status"

        def close_dialog() -> None:
            try:
                page.pop_dialog()
            except Exception:
                try:
                    dialog.open = False
                    dialog.update()
                except Exception:
                    pass
            try:
                page.update()
            except Exception:
                pass

        def choose(reset: bool) -> None:
            close_dialog()
            run_action(lambda: run_with_reset(reset))

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("Reset counterparty status?"),
            content=ft.Text(
                f"This configuration change may affect existing counterparties. Reset {reset_scope} for this quarter?"
            ),
            actions=[
                ft.TextButton("No, retain status", on_click=lambda _event: choose(False)),
                ft.FilledButton("Yes, reset status", on_click=lambda _event: choose(True)),
            ],
        )
        try:
            page.show_dialog(dialog)
        except RuntimeError as exc:
            if "already opened" in str(exc).lower():
                return
            try:
                page.dialog = dialog
                dialog.open = True
                page.update()
            except Exception:
                show_feedback(str(exc), is_error=True)
                safe_update()
        except Exception as exc:
            try:
                page.dialog = dialog
                dialog.open = True
                page.update()
            except Exception:
                show_feedback(f"Could not open confirmation dialog: {exc}", is_error=True)
                safe_update()

    def run_mail_action(action):
        try:
            message = str(action())
            mail_settings_status.value = message
            show_feedback(message)
            refresh()
        except Exception as exc:
            message = f"Mail setup error: {exc}"
            mail_settings_status.value = message
            show_feedback(message, is_error=True)
            safe_update()

    def sync_document_paths_from_field() -> None:
        selected_document_paths[:] = _parse_template_path_strings(document_template_paths.value or "")
        render_selected_document_paths()

    def render_selected_document_paths() -> None:
        if selected_document_paths:
            document_paths_list.controls = [
                ft.Row(
                    [
                        ft.Text(path, expand=True, selectable=True),
                        ft.IconButton(
                            icon=ft.Icons.DELETE_OUTLINE,
                            tooltip="Remove",
                            data=path,
                            on_click=lambda event: remove_document_path(event.control.data),
                        ),
                    ]
                )
                for path in selected_document_paths
            ]
        else:
            document_paths_list.controls = [ft.Text("No document files selected.", selectable=True)]
        document_template_paths.value = "\n".join(selected_document_paths)

    def remove_document_path(path: str) -> None:
        selected_document_paths[:] = [value for value in selected_document_paths if value != path]
        render_selected_document_paths()
        safe_update()

    def clear_document_paths() -> str:
        selected_document_paths.clear()
        render_selected_document_paths()
        return "Cleared selected document paths."

    def refresh_document_templates() -> None:
        if not configure_quarter.value:
            configured_document_template.options = []
            configured_document_template.value = None
            document_validation_status.value = "Create/select a quarter first."
            document_preview.value = ""
            configured_documents_list.controls = [ft.Text("Create/select a quarter first.", selectable=True)]
            return
        rows = controller.list_document_templates(int(configure_quarter.value))
        configured_document_template.options = [
            ft.dropdown.Option(str(row["id"]), f"{row['name']} ({Path(row['file_path']).name if row['file_path'] else 'inline'})")
            for row in rows
        ]
        if configured_document_template.value and not any(str(row["id"]) == configured_document_template.value for row in rows):
            configured_document_template.value = None
        if not configured_document_template.value and rows:
            configured_document_template.value = str(rows[0]["id"])
        validation = controller.quarter_configuration_status(int(configure_quarter.value))
        errors = validation["mapping_errors"]
        if rows and not errors:
            document_validation_status.value = f"{len(rows)} document template(s) configured and validated."
        elif rows:
            document_validation_status.value = "Validation issues: " + "; ".join(errors)
        else:
            document_validation_status.value = "No document templates configured yet."
        configured_documents_list.controls = _configured_document_rows(ft, rows, open_document_preview, request_delete_document)

    def save_document_template_files() -> str:
        sync_document_paths_from_field()
        message = controller.save_document_templates(configured_quarter_id(), "\n".join(selected_document_paths))
        refresh_document_templates()
        return message

    def add_document_template_files() -> str:
        sync_document_paths_from_field()
        message = controller.add_document_templates(configured_quarter_id(), "\n".join(selected_document_paths))
        refresh_document_templates()
        return message

    def refresh_documents_for_ui() -> str:
        refresh_document_templates()
        return "Refreshed configured documents."

    def preview_selected_document_template() -> str:
        if not configured_document_template.value:
            raise ValueError("Select a configured document template first.")
        document_preview.value = controller.preview_document_template(int(configured_document_template.value))
        return "Loaded template preview."

    def open_document_preview(template_id: int, label: str) -> str:
        configured_document_template.value = str(template_id)
        selected_document_preview_title.value = label
        selected_document_update_path.value = ""
        document_preview.value = controller.preview_document_template(template_id)
        show_page("Document Preview")
        return "Opened configured document."

    def open_document_setup() -> str:
        show_page("Document Setup")
        return "Opened document setup."

    def open_mail_input_setup() -> str:
        quarter_id = configured_quarter_id()
        preview = controller.preview_mail_templates(quarter_id)
        mail_subject.value = preview["subject"] or ""
        mail_body.value = preview["body"] or ""
        show_page("Mail Input Setup")
        return "Opened mail input setup."

    def open_mail_preview() -> str:
        preview = controller.preview_mail_templates(configured_quarter_id())
        mail_preview_subject.value = preview["subject"] or "(No saved mail subject yet.)"
        mail_preview_body.value = preview["body"] or "No saved mail body yet."
        show_page("Mail Preview")
        return "Opened configured mail preview."

    def request_delete_document(template_id: int) -> None:
        try:
            configured_document_template.value = str(template_id)
            prompt_counterparty_reset(delete_selected_document_template, reset_documents=True)
        except Exception as exc:
            show_feedback(f"Error: {exc}", is_error=True)
            safe_update()

    def edit_configured_mail() -> str:
        mail_subject.value = "" if mail_preview_subject.value == "(No saved mail subject yet.)" else (mail_preview_subject.value or "")
        mail_body.value = "" if mail_preview_body.value == "No saved mail body yet." else (mail_preview_body.value or "")
        show_page("Mail Input Setup")
        return "Opened mail editor."

    def update_configured_mail() -> str:
        mail_subject.value = "" if mail_preview_subject.value == "(No saved mail subject yet.)" else (mail_preview_subject.value or "")
        mail_body.value = "" if mail_preview_body.value == "No saved mail body yet." else (mail_preview_body.value or "")
        return save_mail_template()

    def update_selected_document_template() -> str:
        if not configured_document_template.value:
            raise ValueError("Select a configured document first.")
        message = controller.update_document_template(int(configured_document_template.value), selected_document_update_path.value)
        refresh_document_templates()
        if configured_document_template.value:
            selected_label = next(
                (
                    getattr(option, "text", None) or str(getattr(option, "key", "Configured document"))
                    for option in configured_document_template.options
                    if str(getattr(option, "key", "")) == str(configured_document_template.value)
                ),
                "Configured document",
            )
            selected_document_preview_title.value = selected_label
            document_preview.value = controller.preview_document_template(int(configured_document_template.value))
        selected_document_update_path.value = ""
        return message

    def delete_selected_document_template() -> str:
        if not configured_document_template.value:
            raise ValueError("Select a configured document first.")
        message = controller.delete_document_template(int(configured_document_template.value))
        configured_document_template.value = None
        document_preview.value = ""
        selected_document_preview_title.value = "Configured document"
        selected_document_update_path.value = ""
        refresh_document_templates()
        return message

    def back_to_configure() -> str:
        show_page("Configure")
        return "Returned to Configure."

    def show_page(name: str, keep_compliance_detail: bool = False) -> None:
        active_page["name"] = name
        if name == "Compliance" and not keep_compliance_detail:
            compliance_detail_visible["value"] = False
            selected_counterparty_ids.clear()
            hide_generated_doc_bubble()
            compliance_intro.visible = True
            compliance_clients.visible = True
            compliance_empty_detail.visible = True
            compliance_detail_container.visible = False
        for card in page_cards:
            card.visible = card.data == name
        for button in nav_buttons:
            button.disabled = button.data == name
        safe_update()

    def selected_workflow_ids() -> set[int]:
        if not selected_counterparty_ids:
            raise ValueError("Select rows first, or use Select Current Page / Select All / Select Next Batch.")
        return set(selected_counterparty_ids)

    def current_row_page_size() -> int:
        try:
            return max(1, int(row_page_size.value or "20"))
        except ValueError:
            row_page_size.value = "20"
            return 20

    def paged_counterparty_rows() -> list[dict]:
        total = len(visible_counterparty_rows)
        page_size = current_row_page_size()
        max_page_index = max(0, (total - 1) // page_size) if total else 0
        row_page_index["value"] = min(max(row_page_index["value"], 0), max_page_index)
        start = row_page_index["value"] * page_size
        end = min(start + page_size, total)
        previous_row_page_button.disabled = row_page_index["value"] <= 0 or total == 0
        next_row_page_button.disabled = row_page_index["value"] >= max_page_index or total == 0
        if total:
            row_page_status.value = f"Showing rows {start + 1}-{end} of {total}; {len(selected_counterparty_ids)} selected."
        else:
            row_page_status.value = "No rows loaded."
        return visible_counterparty_rows[start:end]

    def current_page_counterparty_rows() -> list[dict]:
        total = len(visible_counterparty_rows)
        if not total:
            return []
        page_size = current_row_page_size()
        max_page_index = max(0, (total - 1) // page_size)
        row_page_index["value"] = min(max(row_page_index["value"], 0), max_page_index)
        start = row_page_index["value"] * page_size
        end = min(start + page_size, total)
        return visible_counterparty_rows[start:end]

    def render_compliance_rows() -> None:
        compliance_row_details.controls = [
            ft.Row([row_page_size, previous_row_page_button, next_row_page_button, row_page_status], wrap=True),
            *_counterparty_detail_controls(
                ft,
                paged_counterparty_rows(),
                selected_counterparty_ids,
                toggle_counterparty_selection,
                preview_generated_document_from_row,
            ),
        ]
        if page is None:
            return

    def reset_row_page_size() -> None:
        row_page_index["value"] = 0
        render_compliance_rows()
        update_workflow_button_state()
        safe_update()

    def previous_row_page() -> str:
        row_page_index["value"] = max(0, row_page_index["value"] - 1)
        render_compliance_rows()
        return "Moved to previous row page."

    def next_row_page() -> str:
        row_page_index["value"] += 1
        render_compliance_rows()
        return "Moved to next row page."

    def update_workflow_button_state() -> None:
        has_rows = bool(visible_counterparty_rows)
        has_selection = bool(selected_counterparty_ids)
        readiness = None
        if selected_quarter.value:
            try:
                readiness = controller.quarter_workflow_readiness(int(selected_quarter.value))
            except Exception:
                readiness = None
        selected_rows = [row for row in visible_counterparty_rows if row["counterparty_id"] in selected_counterparty_ids]
        selected_have_docs = bool(selected_rows) and all(row["document_count"] > 0 for row in selected_rows)
        selected_single_with_docs = len(selected_rows) == 1 and selected_rows[0]["document_count"] > 0

        select_all_button.disabled = not has_rows
        select_current_page_button.disabled = not has_rows
        select_batch_button.disabled = not has_rows
        clear_selection_button.disabled = not has_selection
        regenerate_selected_button.disabled = not (
            has_selection
            and readiness is not None
            and readiness["can_generate_documents"]
        )
        queue_selected_button.disabled = not (
            has_selection
            and selected_have_docs
            and readiness is not None
            and readiness["can_send_mail"]
        )
        load_generated_docs_button.disabled = not selected_single_with_docs
        preview_generated_doc_button.disabled = not bool(generated_doc_dropdown.value)

        if not has_rows:
            workflow_action_hint.value = "Persist Excel rows for this quarter before workflow retriggers are available."
        elif not has_selection:
            workflow_action_hint.value = "Select one or more rows, Select All, or Select Next Batch to enable workflow actions."
        elif readiness is not None and readiness["document_template_count"] == 0:
            workflow_action_hint.value = "Save document templates before regenerating documents."
        elif readiness is not None and readiness["document_mapping_errors"]:
            workflow_action_hint.value = "Run Auto Map + Validate and fix document mapping issues before regenerating docs."
        elif not selected_have_docs:
            workflow_action_hint.value = "Regenerate documents for selected rows before sending mail."
        elif readiness is not None and readiness["mail_template_count"] < 2:
            workflow_action_hint.value = "Save mail templates before sending mail."
        elif readiness is not None and readiness["mail_mapping_errors"]:
            workflow_action_hint.value = "Fix mail mappings before sending mail."
        elif readiness is not None and not readiness["send_enabled"]:
            workflow_action_hint.value = "Enable send mode and configure Gmail SMTP or Webtel SMTP to send live mail."
        else:
            workflow_action_hint.value = "Selected rows are ready for document regeneration and live mail send."

    def refresh_compliance_rows(quarter_id: int | None = None) -> None:
        if quarter_id is None:
            quarter_id = selected_quarter_id()
        visible_counterparty_rows[:] = controller.quarter_counterparty_statuses(quarter_id)
        visible_ids = {row["counterparty_id"] for row in visible_counterparty_rows}
        selected_counterparty_ids.intersection_update(visible_ids)
        render_compliance_rows()
        refresh_generated_document_options()
        update_workflow_button_state()

    def refresh_generated_document_options() -> None:
        selected_rows = [row for row in visible_counterparty_rows if row["counterparty_id"] in selected_counterparty_ids]
        if len(selected_rows) != 1:
            generated_doc_dropdown.options = []
            generated_doc_dropdown.value = None
            generated_doc_preview.value = ""
            generated_doc_hint.value = "Select exactly one row with generated docs to preview output."
            return
        row = selected_rows[0]
        if row["document_count"] <= 0:
            generated_doc_dropdown.options = []
            generated_doc_dropdown.value = None
            generated_doc_preview.value = ""
            generated_doc_hint.value = "Selected row has no generated documents yet."
            return
        generated_doc_hint.value = f"Selected: {row['party_name']}. Click Load Generated Docs to fetch preview options."

    def load_generated_documents_for_selected_row() -> str:
        selected_rows = [row for row in visible_counterparty_rows if row["counterparty_id"] in selected_counterparty_ids]
        if len(selected_rows) != 1:
            raise ValueError("Select exactly one row to load generated documents.")
        row = selected_rows[0]
        docs = controller.generated_documents_for_counterparty(selected_quarter_id(), int(row["counterparty_id"]))
        if not docs:
            generated_doc_dropdown.options = []
            generated_doc_dropdown.value = None
            generated_doc_preview.value = ""
            generated_doc_hint.value = "No generated docs found for selected row."
            return "No generated docs found for selected row."
        generated_doc_dropdown.options = [
            ft.dropdown.Option(str(doc["id"]), f"{Path(doc['file_path']).name} ({doc['created_at']})")
            for doc in docs
        ]
        generated_doc_dropdown.value = str(docs[-1]["id"])
        generated_doc_preview.value = controller.preview_generated_document(int(generated_doc_dropdown.value))
        generated_doc_hint.value = f"Loaded {len(docs)} generated document(s)."
        show_generated_doc_bubble(f"Preview: {Path(docs[-1]['file_path']).name}")
        update_workflow_button_state()
        return f"Loaded {len(docs)} generated document(s) for preview."

    def preview_selected_generated_document() -> str:
        if not generated_doc_dropdown.value:
            raise ValueError("Load and select a generated document first.")
        generated_doc_preview.value = controller.preview_generated_document(int(generated_doc_dropdown.value))
        selected_label = next(
            (
                getattr(option, "text", None) or str(getattr(option, "key", "Generated Docs Preview"))
                for option in generated_doc_dropdown.options
                if str(getattr(option, "key", "")) == str(generated_doc_dropdown.value)
            ),
            "Generated Docs Preview",
        )
        show_generated_doc_bubble(f"Preview: {selected_label}")
        return "Updated generated document preview."

    def preview_generated_document_from_row(counterparty_id: int, generated_document_id: int) -> None:
        try:
            docs = controller.generated_documents_for_counterparty(selected_quarter_id(), counterparty_id)
            if not docs:
                generated_doc_dropdown.options = []
                generated_doc_dropdown.value = None
                generated_doc_preview.value = ""
                generated_doc_hint.value = "No generated docs found for this row."
                safe_update()
                return
            generated_doc_dropdown.options = [
                ft.dropdown.Option(str(doc["id"]), f"{Path(doc['file_path']).name} ({doc['created_at']})")
                for doc in docs
            ]
            selected_doc = next((doc for doc in docs if doc["id"] == generated_document_id), docs[0])
            generated_doc_dropdown.value = str(selected_doc["id"])
            generated_doc_preview.value = controller.preview_generated_document(int(selected_doc["id"]))
            generated_doc_hint.value = f"Hover or click any generated doc chip in the rows tab. Loaded {len(docs)} doc(s) for this row."
            show_generated_doc_bubble(f"Preview: {Path(selected_doc['file_path']).name}")
            update_workflow_button_state()
        except Exception as exc:
            generated_doc_preview.value = f"Error: {exc}"
            generated_doc_hint.value = "Could not load generated document preview."
        safe_update()

    def toggle_counterparty_selection(counterparty_id: int, selected: bool) -> None:
        if selected:
            selected_counterparty_ids.add(counterparty_id)
        else:
            selected_counterparty_ids.discard(counterparty_id)
        refresh_compliance_rows()
        safe_update()

    def select_all_counterparties() -> str:
        if not visible_counterparty_rows:
            refresh_compliance_rows()
        selected_counterparty_ids.update(row["counterparty_id"] for row in visible_counterparty_rows)
        refresh_compliance_rows()
        return f"Selected {len(selected_counterparty_ids)} row(s)."

    def select_current_page_counterparties() -> str:
        if not visible_counterparty_rows:
            refresh_compliance_rows()
        page_rows = current_page_counterparty_rows()
        selected_counterparty_ids.update(row["counterparty_id"] for row in page_rows)
        refresh_compliance_rows()
        return f"Selected {len(page_rows)} row(s) on the current page."

    def clear_counterparty_selection() -> str:
        selected_counterparty_ids.clear()
        refresh_compliance_rows()
        return "Cleared row selection."

    def select_next_counterparty_batch() -> str:
        if not visible_counterparty_rows:
            refresh_compliance_rows()
        try:
            batch_size = max(1, int(workflow_batch_size.value or "20"))
        except ValueError as exc:
            raise ValueError("Enter a numeric batch size.") from exc
        remaining = [row["counterparty_id"] for row in visible_counterparty_rows if row["counterparty_id"] not in selected_counterparty_ids]
        selected_counterparty_ids.update(remaining[:batch_size])
        refresh_compliance_rows()
        return f"Selected {min(batch_size, len(remaining))} next row(s)."

    def regenerate_selected_documents() -> str:
        message = controller.regenerate_documents(selected_quarter_id(), selected_workflow_ids())
        refresh_compliance_rows()
        return message

    def send_selected_mail() -> str:
        message = controller.send_mail_selected(selected_quarter_id(), selected_workflow_ids())
        refresh_compliance_rows()
        return message

    def open_client_compliance(client_id: int) -> str:
        quarter_id = controller.current_quarter_id_for_client(client_id)
        # Keep global quarter picker + Configure client's quarter picker aligned with the org we opened.
        # Otherwise _refresh_impl can rewrite selected_quarter to a different client's "current quarter"
        # when configure_client_id still pointed at whoever was last edited in Configure — leaving
        # Compliance rows vs readiness vs selection out of sync and workflow buttons permanently disabled.
        configure_client_id["value"] = client_id
        selected_quarter.value = str(quarter_id)
        configure_quarter.value = str(quarter_id)
        compliance_detail_visible["value"] = True
        hide_generated_doc_bubble()
        show_page("Compliance", keep_compliance_detail=True)
        refresh_compliance_rows(quarter_id)
        return "Opened client compliance."

    def open_client_settings(client_id: int) -> str:
        configure_client_id["value"] = client_id
        try:
            quarter_id = controller.current_quarter_id_for_client(client_id)
            selected_quarter.value = str(quarter_id)
            configure_quarter.value = str(quarter_id)
        except ValueError:
            selected_quarter.value = None
            configure_quarter.value = None
        show_page("Configure")
        return "Opened client configuration."

    def refresh() -> None:
        if ui_guard["busy"]:
            return
        ui_guard["busy"] = True
        try:
            _refresh_impl()
        finally:
            ui_guard["busy"] = False

    def _refresh_impl() -> None:
        selected_id = int(selected_quarter.value) if selected_quarter.value else None
        snapshot = controller.snapshot(selected_id)
        clients = snapshot["clients"]
        quarters = snapshot["quarters"]
        client_lookup = {client.id: client.name for client in clients}

        selected_quarter.options = [
            ft.dropdown.Option(str(item.id), _quarter_label(item, client_lookup))
            for item in quarters
            if item.id is not None
        ]
        configure_client.options = [
            ft.dropdown.Option(str(client.id), client.name)
            for client in clients
            if client.id is not None
        ]
        current = snapshot["current_quarter"]
        if not selected_quarter.value and current and current.id is not None:
            selected_quarter.value = str(current.id)

        clients_list.controls = _client_card_controls(ft, snapshot["client_cards"], open_client_compliance, open_client_settings, run_action)
        compliance_clients.controls = _compliance_client_controls(ft, snapshot["client_cards"], open_client_compliance, run_action)

        selected = snapshot["selected_quarter"]
        summary = snapshot["summary"]
        detail_open = bool(compliance_detail_visible["value"] and selected and selected.id is not None)
        compliance_intro.visible = not detail_open
        compliance_clients.visible = not detail_open
        compliance_empty_detail.visible = not detail_open
        compliance_detail_container.visible = detail_open
        dashboard_cards.controls.clear()
        workflow_counts.controls.clear()
        if selected and summary:
            selected_label = next((_quarter_label(item, client_lookup) for item in quarters if str(item.id) == selected_quarter.value), "none")
            current_label = _quarter_label(current, client_lookup) if current else "none"
            selected_client_name = client_lookup.get(selected.client_id, "Selected organization")
            compliance_detail_header.value = f"{selected_client_name} - {selected.financial_year} {selected.quarter}"
            current_summary.value = (
                f"Current quarter running: {current_label}. "
                f"Compliance: {summary.fully_compliant}/{summary.total} counterparties compliant, "
                f"{summary.non_compliant} non-compliant. Client status: {_client_compliance_status(summary)}."
            )
            workflow_context.value = f"Workflow target: {selected_label}"
            dashboard_cards.controls.extend(
                [
                    _metric_card(ft, "Non Compliant", summary.non_compliant),
                    _metric_card(ft, "Compliant", summary.fully_compliant),
                    _metric_card(ft, "Counterparties", summary.total),
                ]
            )
            for label, values in (("Document jobs", summary.generation_counts), ("Emails", summary.email_counts)):
                workflow_counts.controls.append(ft.Text(f"{label}: {_count_summary(values)}"))
            workflow_counts.controls.append(ft.Text(f"Counterparties: {_count_summary(summary.party_type_counts)}"))
            compliance_detail.value = (
                f"Selected: {selected_label}. Client status: {_client_compliance_status(summary)}. "
                f"{summary.fully_compliant}/{summary.total} counterparties compliant, {summary.non_compliant} non-compliant. "
                f"Types: {_count_summary(summary.party_type_counts)}. "
                f"Docs: {_count_summary(summary.generation_counts)}. Emails: {_count_summary(summary.email_counts)}."
            )
            try:
                readiness = controller.quarter_workflow_readiness(int(selected.id))
                allowed = []
                if readiness["can_generate_documents"]:
                    allowed.append("regenerate documents")
                if readiness["can_send_mail"]:
                    allowed.append("send selected mail")
                compliance_actions_allowed.value = "Actions allowed: " + (", ".join(allowed) if allowed else "finish configuration before row actions are enabled")
            except Exception as exc:
                compliance_actions_allowed.value = f"Actions allowed: unavailable ({exc})"
        else:
            current_summary.value = "No current quarter yet. Create a client, then configure a quarter."
            workflow_context.value = "Create or select a client quarter before running workflow actions."
            compliance_detail_header.value = "Select a client's current quarter."
            compliance_detail.value = "Open a client's current quarter to see its compliance details."
            compliance_actions_allowed.value = "Actions become available after opening a current quarter."
            dashboard_cards.controls.append(_metric_card(ft, "No Current Quarter", 0))
            workflow_counts.controls.append(ft.Text("Workflow counts: none"))
        if configure_client_id["value"] is None and selected_quarter.value:
            try:
                selected_quarter_id_value = int(selected_quarter.value)
            except ValueError:
                selected_quarter_id_value = None
            if selected_quarter_id_value is not None:
                selected_quarter_obj = next((item for item in quarters if item.id == selected_quarter_id_value), None)
                if selected_quarter_obj is not None:
                    configure_client_id["value"] = selected_quarter_obj.client_id
        if configure_client_id["value"] is not None and not any(client.id == configure_client_id["value"] for client in clients):
            configure_client_id["value"] = None
        configure_client.value = str(configure_client_id["value"]) if configure_client_id["value"] is not None else None

        # Keep Configure client in sync with the header quarter only while actively viewing
        # Compliance detail. Doing this globally can override a newly created client with no
        # quarter yet (for example, "acme"), making it appear unconfigurable.
        if active_page["name"] == "Compliance" and compliance_detail_visible["value"] and selected_quarter.value and configure_client_id["value"] is not None:
            try:
                header_qid = int(selected_quarter.value)
            except ValueError:
                header_qid = None
            if header_qid is not None:
                header_quarter = next((item for item in quarters if item.id == header_qid), None)
                if header_quarter is not None and header_quarter.client_id != configure_client_id["value"]:
                    configure_client_id["value"] = header_quarter.client_id

        client_quarters = [item for item in quarters if item.client_id == configure_client_id["value"]]
        configure_quarter.options = [
            ft.dropdown.Option(str(item.id), f"{item.financial_year} {item.quarter} ({'current' if item.current_quarter else item.status})")
            for item in client_quarters
            if item.id is not None
        ]
        if configure_quarter.value and not any(str(item.id) == configure_quarter.value for item in client_quarters):
            configure_quarter.value = None
        if configure_client_id["value"] is not None and selected_quarter.value:
            try:
                sync_from_global = int(selected_quarter.value)
            except ValueError:
                sync_from_global = None
            if sync_from_global is not None and any(item.id == sync_from_global for item in client_quarters):
                configure_quarter.value = str(sync_from_global)
        if not configure_quarter.value:
            current_for_client = _latest_quarter(client_quarters)
            if current_for_client and current_for_client.id is not None:
                configure_quarter.value = str(current_for_client.id)
                selected_quarter.value = str(current_for_client.id)
        if configure_quarter.value:
            config_status = controller.quarter_configuration_status(int(configure_quarter.value))
            variables = ", ".join(config_status["variables"]) or "none"
            mapping = "valid" if not config_status["mapping_errors"] else "; ".join(config_status["mapping_errors"])
            inputs_status.value = (
                f"Imports: {config_status['import_count']}. "
                f"Detected columns: {len(config_status['latest_columns'])}. "
                f"Variables: {variables}. Mapping: {mapping}."
            )
            refresh_document_templates()
        elif configure_client_id["value"] is not None:
            inputs_status.value = "Create a quarter before importing Excel, saving templates, or mapping variables."
            configured_document_template.options = []
            configured_document_template.value = None
            document_validation_status.value = "Create/select a quarter first."
            document_preview.value = ""
        else:
            inputs_status.value = "Select a client to configure inputs and mapping."
            configured_document_template.options = []
            configured_document_template.value = None
            document_validation_status.value = "Create/select a quarter first."
            document_preview.value = ""
        settings_name = client_lookup.get(configure_client_id["value"], "")
        settings_title.value = f"Configure {settings_name}" if settings_name else "Select a client from Clients to configure."
        mail_settings_title.value = f"Mail setup for {settings_name}" if settings_name else "Select a client from Clients to configure mail setup."
        configure_empty.visible = configure_client_id["value"] is None
        configure_form.visible = configure_client_id["value"] is not None
        mail_setup_empty.visible = configure_client_id["value"] is None
        mail_setup_form.visible = configure_client_id["value"] is not None
        if active_page["name"] == "Compliance" and compliance_detail_visible["value"] and selected_quarter.value:
            refresh_compliance_rows(int(selected_quarter.value))
        settings = controller.get_mail_settings()
        selected_mail_provider["value"] = normalized_mail_provider(settings["mail_provider"])
        mail_provider.value = selected_mail_provider["value"]
        send_mode.value = settings["send_mode"]
        sender_email_field.value = settings["sender_email"]
        fallback_providers_field.value = settings["fallback_providers"]
        daily_send_limit_field.value = settings["daily_send_limit"]
        per_email_delay_field.value = settings["per_email_delay_seconds"]
        smtp_username_field.value = settings["smtp_username"]
        smtp_password_field.value = "saved" if settings["smtp_password_saved"] else ""
        mail_settings_status.value = (
            f"Mail setup: provider={settings['mail_provider']}, connected={settings['connected_email'] or 'not connected'}, "
            f"fallbacks={settings['fallback_providers'] or 'none'}, send_mode={settings['send_mode']}."
        )
        update_mail_setup_controls()
        safe_update()

    def on_selected_quarter_changed(_event) -> None:
        if ui_guard["busy"]:
            return
        refresh()

    def on_configure_quarter_changed(_event) -> None:
        if ui_guard["busy"]:
            return
        if configure_quarter.value:
            selected_quarter.value = configure_quarter.value
        refresh()

    def on_configure_client_changed(_event) -> None:
        if ui_guard["busy"]:
            return
        if configure_client.value:
            configure_client_id["value"] = int(configure_client.value)
            configure_quarter.value = None
        else:
            configure_client_id["value"] = None
            configure_quarter.value = None
        refresh()

    selected_quarter.on_change = on_selected_quarter_changed
    configure_quarter.on_change = on_configure_quarter_changed
    configure_client.on_change = on_configure_client_changed

    def on_mail_provider_changed(event) -> None:
        raw_provider = getattr(event, "data", None) or getattr(getattr(event, "control", None), "value", None) or mail_provider.value
        selected_mail_provider["value"] = normalized_mail_provider(raw_provider)
        update_mail_setup_controls()
        safe_update()

    mail_provider.on_change = on_mail_provider_changed

    async def pick_excel_path_only() -> None:
        # page.run_task(handler) invokes handler() with no ControlEvent argument.
        if excel_picker is None or page is None:
            return
        files = await excel_picker.pick_files(
            dialog_title="Select Excel workbook",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx", "xlsm"],
            allow_multiple=False,
        )
        if files and files[0].path:
            workbook_path.value = files[0].path
            page.update()

    async def pick_document_templates() -> None:
        if excel_picker is None or page is None:
            return
        files = await excel_picker.pick_files(
            dialog_title="Select document templates or attachments",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["docx", "txt", "html", "htm", "md", "pdf"],
            allow_multiple=True,
        )
        paths = [file.path for file in files or [] if file.path]
        if not paths:
            return
        selected_document_paths[:] = list(dict.fromkeys([*selected_document_paths, *paths]))
        render_selected_document_paths()
        page.update()

    async def pick_replacement_document_template() -> None:
        if excel_picker is None or page is None:
            return
        files = await excel_picker.pick_files(
            dialog_title="Select replacement document template",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["docx", "txt", "html", "htm", "md", "pdf"],
            allow_multiple=False,
        )
        if files and files[0].path:
            selected_document_update_path.value = files[0].path
            page.update()

    async def pick_mail_body_template() -> None:
        if excel_picker is None or page is None:
            return
        files = await excel_picker.pick_files(
            dialog_title="Select mail body template",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["txt", "html", "htm", "md"],
            allow_multiple=False,
        )
        if not files or not files[0].path:
            return
        path = Path(files[0].path)
        try:
            mail_body.value = path.read_text(encoding="utf-8")
            status.value = f"Loaded mail body template: {path.name}."
        except Exception as exc:
            status.value = f"Error: {exc}"
        page.update()

    def import_workbook_from_field() -> str:
        return controller.import_workbook(
            configured_quarter_id(),
            workbook_path.value,
            client_id=configured_client_id(),
        )

    def create_client() -> str:
        created = controller.create_client_record(client_name.value, client_type.value)
        configure_client_id["value"] = created["id"]
        client_name.value = ""
        show_page("Configure")
        return str(created["message"]) + " Configure the client to add quarters, imports, and templates."

    def create_configured_quarter() -> str:
        return controller.create_quarter(configured_client_id(), financial_year.value, quarter.value, current=True)

    def save_mail_template() -> str:
        quarter_id = configured_quarter_id()
        subject_to_save = mail_subject.value or ""
        body_to_save = mail_body.value or ""
        message = controller.save_mail_templates(quarter_id, subject_to_save, body_to_save)
        # Re-read persisted values so UI reflects DB state even after refreshes/dialog flows.
        preview = controller.preview_mail_templates(quarter_id)
        mail_subject.value = preview["subject"] or ""
        mail_body.value = preview["body"] or ""
        mail_preview_subject.value = preview["subject"] or "(No saved mail subject yet.)"
        mail_preview_body.value = preview["body"] or "No saved mail body yet."
        return message

    def save_mail_settings_action() -> str:
        provider = normalized_mail_provider()
        smtp_host, smtp_port = smtp_defaults_for_provider(provider)
        smtp_username = sender_email_field.value if provider == "webtel_smtp" else (smtp_username_field.value or "")
        return controller.save_mail_settings(
            mail_provider=provider,
            send_mode=send_mode.value or "preview",
            sender_email=sender_email_field.value or "",
            fallback_providers=fallback_providers_field.value or "",
            daily_send_limit=int(daily_send_limit_field.value or "500"),
            per_email_delay_seconds=int(per_email_delay_field.value or "3"),
            smtp_host=smtp_host,
            smtp_port=smtp_port,
            smtp_username=smtp_username,
            smtp_password="" if (smtp_password_field.value or "").strip() == "saved" else (smtp_password_field.value or ""),
        )

    def primary_mail_setup_action() -> str:
        return save_mail_settings_action()

    def test_connection_action() -> str:
        return controller.test_mail_connection()

    def send_test_email_action() -> str:
        recipient = (test_recipient_field.value or sender_email_field.value or "").strip()
        if not test_recipient_field.value and recipient:
            test_recipient_field.value = recipient
        return controller.send_test_email(recipient)

    def open_mail_setup() -> str:
        show_page("Mail Setup")
        return "Opened Mail Setup."

    def normalized_mail_provider(raw_provider: object | None = None) -> str:
        if raw_provider is None:
            raw_provider = selected_mail_provider.get("value") or mail_provider.value or "gmail_smtp"
        raw = str(raw_provider or "gmail_smtp").strip().lower()
        if raw in {"gmail_smtp", "webtel_smtp"}:
            return raw
        if "webtel" in raw:
            return "webtel_smtp"
        if "gmail" in raw:
            return "gmail_smtp"
        return "gmail_smtp"

    def smtp_defaults_for_provider(provider: str) -> tuple[str, int]:
        if provider == "webtel_smtp":
            return "connect.webtelconnect.com", 465
        return "smtp.gmail.com", 587

    def select_mail_provider(provider: str) -> None:
        selected = normalized_mail_provider(provider)
        selected_mail_provider["value"] = selected
        if selected == "webtel_smtp":
            fallback_providers_field.value = ""
        update_mail_setup_controls()
        safe_update()

    save_mail_settings_button = ft.FilledButton("Save Mail Settings", on_click=lambda _event: run_mail_action(primary_mail_setup_action))
    test_connection_button = ft.OutlinedButton("Test Connection", on_click=lambda _event: run_mail_action(test_connection_action))
    send_test_email_button = ft.OutlinedButton("Send Test Email", on_click=lambda _event: run_mail_action(send_test_email_action))
    select_webtel_smtp_button = ft.OutlinedButton("Use Webtel SMTP", on_click=lambda _event: select_mail_provider("webtel_smtp"))
    select_gmail_smtp_button = ft.OutlinedButton("Use Gmail SMTP", on_click=lambda _event: select_mail_provider("gmail_smtp"))
    provider_selector = ft.Column(
        [
            ft.Text("Choose how this sender account will authenticate:"),
            ft.Row([select_gmail_smtp_button, select_webtel_smtp_button], wrap=True),
            provider_choice_status,
        ]
    )
    provider_next_steps_section = ft.Container(
        padding=12,
        border_radius=12,
        bgcolor=ft.Colors.BLUE_GREY_50,
        content=mail_setup_next_steps,
    )
    smtp_settings_section = _section(
        ft,
        "SMTP Credentials",
        ft.Column(
            [
                ft.Text("Use this for a Gmail app password or Webtel mailbox password."),
                smtp_defaults_text,
                smtp_username_field,
                smtp_password_field,
            ]
        ),
    )
    delivery_settings_section = _section(
        ft,
        "Delivery Settings",
        ft.Column(
            [
                ft.Row([daily_send_limit_field, per_email_delay_field], wrap=True),
                fallback_providers_field,
                test_recipient_field,
                ft.Row([save_mail_settings_button, test_connection_button, send_test_email_button], wrap=True),
            ]
        ),
    )

    def update_mail_setup_controls() -> None:
        provider = normalized_mail_provider()
        selected_mail_provider["value"] = provider
        mail_provider.value = provider
        webtel_selected = provider == "webtel_smtp"

        select_webtel_smtp_button.disabled = webtel_selected
        select_gmail_smtp_button.disabled = provider == "gmail_smtp"
        smtp_username_field.visible = not webtel_selected

        if provider == "webtel_smtp":
            provider_choice_status.value = "Selected provider: Webtel SMTP. Enter only the sender email and mailbox password."
            save_mail_settings_button.text = "Save Webtel SMTP"
            sender_email_field.hint_text = "Example: ghanshyam@ngmks.in"
            smtp_username_field.value = sender_email_field.value
            smtp_defaults_text.value = "The app will use Webtel SMTP internally: connect.webtelconnect.com:465 with SSL/TLS."
            smtp_password_field.label = "Webtel mailbox password"
            fallback_providers_field.value = fallback_providers_field.value or ""
            mail_setup_next_steps.value = (
                "Next steps for Webtel SMTP:\n"
                "1) Enter the sender email, for example ghanshyam@ngmks.in.\n"
                "2) Enter the mailbox password in Webtel mailbox password.\n"
                "3) Click Save Webtel SMTP, then Test Connection.\n"
                "4) Send Test Email. The app uses connect.webtelconnect.com:465 SSL/TLS automatically."
            )
        else:
            provider_choice_status.value = "Selected provider: Gmail SMTP. Gmail app password is required for this path."
            save_mail_settings_button.text = "Save Mail Settings"
            sender_email_field.hint_text = "Gmail sender email"
            smtp_username_field.value = smtp_username_field.value or sender_email_field.value
            smtp_defaults_text.value = "The app will use Gmail SMTP internally: smtp.gmail.com:587 with STARTTLS."
            smtp_password_field.label = "SMTP password/app password"
            fallback_providers_field.value = fallback_providers_field.value or "webtel_smtp"
            mail_setup_next_steps.value = (
                "Next steps for Gmail SMTP:\n"
                "1) Enter Gmail sender email and app password.\n"
                "2) The app uses smtp.gmail.com:587 automatically.\n"
                "3) Click Test Connection, then Send Test Email.\n"
                "4) Save Mail Settings and use send mode 'send' for live mail."
            )

    def on_generated_doc_dropdown_changed(_event) -> None:
        if not generated_doc_dropdown.value:
            update_workflow_button_state()
            safe_update()
            return
        try:
            status.value = preview_selected_generated_document()
            update_workflow_button_state()
        except Exception as exc:
            show_feedback(f"Error: {exc}", is_error=True)
            return
        safe_update()

    row_page_size.on_change = lambda _event: reset_row_page_size()
    previous_row_page_button = ft.OutlinedButton("Previous Page", on_click=lambda _event: run_action(previous_row_page), disabled=True)
    next_row_page_button = ft.OutlinedButton("Next Page", on_click=lambda _event: run_action(next_row_page), disabled=True)
    select_all_button = ft.OutlinedButton("Select All", on_click=lambda _event: run_action(select_all_counterparties), disabled=True)
    select_current_page_button = ft.OutlinedButton("Select Current Page", on_click=lambda _event: run_action(select_current_page_counterparties), disabled=True)
    select_batch_button = ft.OutlinedButton("Select Next Batch", on_click=lambda _event: run_action(select_next_counterparty_batch), disabled=True)
    clear_selection_button = ft.OutlinedButton("Clear Selection", on_click=lambda _event: run_action(clear_counterparty_selection), disabled=True)
    regenerate_selected_button = ft.FilledButton("Regenerate Docs for Selection", on_click=lambda _event: run_action(regenerate_selected_documents), disabled=True)
    queue_selected_button = ft.FilledButton("Send Mail for Selection", on_click=lambda _event: run_action(send_selected_mail), disabled=True)
    load_generated_docs_button = ft.OutlinedButton("Load Generated Docs", on_click=lambda _event: run_action(load_generated_documents_for_selected_row), disabled=True)
    preview_generated_doc_button = ft.OutlinedButton("Preview Selected Generated Doc", on_click=lambda _event: run_action(preview_selected_generated_document), disabled=True)
    generated_doc_dropdown.on_change = on_generated_doc_dropdown_changed
    render_selected_document_paths()

    if page is not None and excel_picker is not None:
        excel_import_row = ft.Row(
            [
                ft.OutlinedButton(
                    "Choose Excel",
                    on_click=lambda _e: page.run_task(pick_excel_path_only),
                ),
                ft.FilledButton(
                    "Persist Excel for Quarter",
                    on_click=lambda _event: run_action(import_workbook_from_field),
                ),
            ]
        )
    else:
        excel_import_row = ft.Row(
            [ft.FilledButton("Persist Excel for Quarter", on_click=lambda _event: run_action(import_workbook_from_field))]
        )

    if page is not None and excel_picker is not None:
        document_template_row = ft.Row(
            [
                ft.OutlinedButton("Browse Docs…", on_click=lambda _event: page.run_task(pick_document_templates)),
                ft.FilledButton(
                    "Configure Docs (Replace)",
                    on_click=lambda _event: prompt_counterparty_reset(save_document_template_files, reset_documents=True),
                ),
                ft.OutlinedButton(
                    "Add New Doc",
                    on_click=lambda _event: prompt_counterparty_reset(add_document_template_files, reset_documents=True),
                ),
            ],
            wrap=True,
        )
        document_update_row = ft.Row(
            [
                ft.OutlinedButton("Browse Replacement…", on_click=lambda _event: page.run_task(pick_replacement_document_template)),
                ft.FilledButton(
                    "Update Document",
                    on_click=lambda _event: prompt_counterparty_reset(update_selected_document_template, reset_documents=True),
                ),
                ft.OutlinedButton(
                    "Delete Document",
                    on_click=lambda _event: prompt_counterparty_reset(delete_selected_document_template, reset_documents=True),
                ),
            ],
            wrap=True,
        )
        mail_template_row = ft.Row(
            [
                ft.OutlinedButton("Load Mail Body…", on_click=lambda _event: page.run_task(pick_mail_body_template)),
                ft.FilledButton(
                    "Update Mail Template",
                    on_click=lambda _event: prompt_counterparty_reset(save_mail_template, reset_mail=True),
                ),
            ]
        )
    else:
        document_template_row = ft.Row(
            [
                ft.FilledButton(
                    "Configure Docs (Replace)",
                    on_click=lambda _event: prompt_counterparty_reset(save_document_template_files, reset_documents=True),
                ),
                ft.OutlinedButton(
                    "Add New Doc",
                    on_click=lambda _event: prompt_counterparty_reset(add_document_template_files, reset_documents=True),
                ),
            ]
        )
        document_update_row = ft.Row(
            [
                ft.FilledButton(
                    "Update Document",
                    on_click=lambda _event: prompt_counterparty_reset(update_selected_document_template, reset_documents=True),
                ),
                ft.OutlinedButton(
                    "Delete Document",
                    on_click=lambda _event: prompt_counterparty_reset(delete_selected_document_template, reset_documents=True),
                ),
            ]
        )
        mail_template_row = ft.Row(
            [
                ft.FilledButton(
                    "Update Mail Template",
                    on_click=lambda _event: prompt_counterparty_reset(save_mail_template, reset_mail=True),
                )
            ]
        )

    home_page = _section(
        ft,
        "Home",
        ft.Column(
            [
                current_summary,
                ft.Row([ft.OutlinedButton("Refresh", on_click=lambda _event: refresh())]),
                dashboard_cards,
            ],
            scroll=ft.ScrollMode.AUTO,
        ),
    )
    configure_form = ft.Column(
        [
            settings_title,
            ft.Text("Configuration is available after selecting or creating a client. Open each setup page to edit inputs; this page only lists what is configured."),
            ft.Row([configure_client, configure_quarter, financial_year, quarter, ft.FilledButton("Create Quarter", on_click=lambda _event: run_action(create_configured_quarter))]),
            _section(
                ft,
                "Counterparty Input",
                ft.Column(
                    [
                        ft.Text("Choose the master Excel and persist counterparties for this quarter."),
                        workbook_path,
                        excel_import_row,
                    ]
                ),
            ),
            _section(
                ft,
                "Mail Input Configuration",
                ft.Column(
                    [
                        ft.Text("Configure the saved mail subject and body for this quarter."),
                        ft.Row(
                            [
                                ft.FilledButton("Configure Mail Input", on_click=lambda _event: run_action(open_mail_input_setup)),
                                ft.OutlinedButton("View Configured Mail", on_click=lambda _event: run_action(open_mail_preview)),
                                ft.OutlinedButton("SMTP Setup", on_click=lambda _event: run_action(open_mail_setup)),
                            ]
                        ),
                    ]
                ),
            ),
            _section(
                ft,
                "Document Configure Setup",
                ft.Column(
                    [
                        ft.Text("Configured document templates for this quarter."),
                        ft.Row(
                            [
                                ft.FilledButton("Configure Docs", on_click=lambda _event: run_action(open_document_setup)),
                                ft.OutlinedButton("Refresh Docs", on_click=lambda _event: run_action(refresh_documents_for_ui)),
                            ]
                        ),
                        configured_documents_list,
                    ]
                ),
            ),
        ]
    )

    clients_page = _section(
        ft,
        "Clients",
        ft.Column(
            [
                ft.Text("Create clients and open Configure or Compliance for existing clients."),
                ft.Row([client_name, client_type, ft.FilledButton("Create Client", on_click=lambda _event: run_action(create_client))]),
                _section(ft, "Clients", clients_list),
            ],
            scroll=ft.ScrollMode.AUTO,
        ),
    )
    configure_page = _section(
        ft,
        "Configure",
        ft.Column(
            [
                configure_empty,
                configure_form,
            ],
            scroll=ft.ScrollMode.AUTO,
        ),
    )
    mail_input_page = _section(
        ft,
        "Mail Input Setup",
        ft.Column(
            [
                ft.Row([ft.OutlinedButton("Back to Configure", on_click=lambda _event: run_action(back_to_configure))]),
                ft.Text("Configure only the mail input template for this quarter."),
                mail_subject,
                mail_body,
                mail_template_row,
            ],
            scroll=ft.ScrollMode.AUTO,
        ),
    )
    mail_preview_page = _section(
        ft,
        "Configured Mail",
        ft.Column(
            [
                ft.Row([ft.OutlinedButton("Back to Configure", on_click=lambda _event: run_action(back_to_configure))]),
                mail_preview_subject,
                mail_preview_body,
                ft.Row(
                    [
                        ft.OutlinedButton("Edit on Mail Input Page", on_click=lambda _event: run_action(edit_configured_mail)),
                        ft.FilledButton(
                            "Update Mail",
                            on_click=lambda _event: prompt_counterparty_reset(update_configured_mail, reset_mail=True),
                        ),
                    ]
                ),
            ],
            scroll=ft.ScrollMode.AUTO,
        ),
    )
    document_setup_page = _section(
        ft,
        "Document Setup",
        ft.Column(
            [
                ft.Row([ft.OutlinedButton("Back to Configure", on_click=lambda _event: run_action(back_to_configure))]),
                ft.Text("Manage document templates/static attachments only. Use Liquid variables such as {{ party_name }} or {{ row.balance }}."),
                document_template_paths,
                document_paths_list,
                document_template_row,
                ft.Row(
                    [
                        ft.OutlinedButton("Clear Docs", on_click=lambda _event: run_action(clear_document_paths)),
                        ft.OutlinedButton("Refresh Docs", on_click=lambda _event: run_action(refresh_documents_for_ui)),
                    ],
                    wrap=True,
                ),
                document_validation_status,
                ft.OutlinedButton("Auto Map + Validate", on_click=lambda _event: run_action(lambda: controller.auto_map_variables(configured_quarter_id()))),
                inputs_status,
            ],
            scroll=ft.ScrollMode.AUTO,
        ),
    )
    document_preview_page = _section(
        ft,
        "Configured Document",
        ft.Column(
            [
                ft.Row([ft.OutlinedButton("Back to Configure", on_click=lambda _event: run_action(back_to_configure))]),
                selected_document_preview_title,
                document_preview,
                ft.Text("To edit a configured document, choose a replacement file and update this document row."),
                ft.Row([selected_document_update_path], scroll=ft.ScrollMode.AUTO),
                document_update_row,
            ],
            scroll=ft.ScrollMode.AUTO,
        ),
    )
    mail_setup_form = ft.Column(
        [
            mail_settings_title,
            ft.Text("Choose Gmail SMTP or Webtel SMTP. The setup steps and required inputs below will change based on that choice."),
            provider_selector,
            ft.Row([send_mode, sender_email_field], wrap=True),
            provider_next_steps_section,
            smtp_settings_section,
            delivery_settings_section,
            mail_settings_status,
            ft.Text("Where inputs are saved: non-secret mail settings go to the local app database; passwords go through keyring when available."),
        ]
    )
    mail_setup_page = _section(
        ft,
        "Mail Setup",
        ft.Column(
            [
                mail_setup_empty,
                mail_setup_form,
                ft.Text("Supported sending providers: Gmail SMTP and Webtel SMTP."),
            ],
            scroll=ft.ScrollMode.AUTO,
        ),
    )
    generated_docs_rows_panel.controls = [
        ft.Text("Review row status and hover or click the Generated Docs column to preview outputs inline."),
        compliance_row_details,
    ]
    generated_doc_preview_dismiss = ft.Container(
        left=0,
        top=0,
        right=0,
        bottom=0,
        bgcolor="#01000000",
        visible=False,
        on_click=lambda _event: (hide_generated_doc_bubble(), safe_update()),
    )
    generated_doc_preview_bubble = ft.Container(
        width=620,
        padding=16,
        border_radius=18,
        bgcolor=ft.Colors.WHITE,
        opacity=0.94,
        border=_border_all(ft, 1, ft.Colors.BLUE_GREY_200),
        shadow=ft.BoxShadow(blur_radius=22, color=ft.Colors.BLUE_GREY_100),
        right=16,
        top=52,
        visible=False,
        content=ft.Column(
            [
                ft.Row(
                    [
                        generated_doc_preview_title,
                        ft.IconButton(ft.Icons.CLOSE, tooltip="Close preview", on_click=lambda _event: (hide_generated_doc_bubble(), safe_update())),
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                ),
                generated_doc_hint,
                generated_doc_dropdown,
                ft.Row([load_generated_docs_button, preview_generated_doc_button], wrap=True),
                generated_doc_preview,
            ],
            spacing=8,
        ),
    )
    generated_docs_rows_stack.controls = [
        generated_docs_rows_panel,
        generated_doc_preview_dismiss,
        generated_doc_preview_bubble,
    ]
    compliance_detail_container = ft.Column(
        [
            compliance_detail_header,
            compliance_detail,
            compliance_actions_allowed,
            _section(
                ft,
                "Row Workflow Actions",
                ft.Column(
                    [
                        ft.Text("Select individual rows, the current page, all rows, or the next batch before retriggering workflow components."),
                        workflow_action_hint,
                        ft.Row(
                            [
                                workflow_batch_size,
                                select_current_page_button,
                                select_all_button,
                                select_batch_button,
                                clear_selection_button,
                            ],
                            wrap=True,
                        ),
                        ft.Row(
                            [
                                regenerate_selected_button,
                                queue_selected_button,
                            ],
                            wrap=True,
                        ),
                    ]
                ),
            ),
            _section(ft, "Excel Row Workflow Status", generated_docs_rows_stack),
        ],
        visible=False,
    )
    compliance_page = _section(
        ft,
        "Compliance",
        ft.Column(
            [
                compliance_intro,
                compliance_clients,
                compliance_empty_detail,
                compliance_detail_container,
            ],
            scroll=ft.ScrollMode.AUTO,
        ),
    )
    workflow_page = _section(
        ft,
        "Workflow",
        ft.Column(
            [
                ft.Text("Run document generation and preview email queueing for the selected client's quarter."),
                workflow_context,
                _section(ft, "Workflow Step Counts", workflow_counts),
                ft.Row(
                    [
                        ft.FilledButton("Generate Documents", on_click=lambda _event: run_action(lambda: controller.run_generation(selected_quarter_id()))),
                        ft.FilledButton("Queue + Preview Send", on_click=lambda _event: run_action(lambda: controller.queue_and_preview_send(selected_quarter_id()))),
                    ]
                ),
                ft.Text("Preview send marks queued emails as sent without contacting the mail provider. Live sends use configured Gmail or Webtel SMTP settings."),
            ]
        ),
    )

    page_cards.extend([
        home_page,
        clients_page,
        configure_page,
        mail_input_page,
        mail_preview_page,
        document_setup_page,
        document_preview_page,
        mail_setup_page,
        compliance_page,
        workflow_page,
    ])
    navigable_pages = {"Home", "Clients", "Configure", "Mail Setup", "Compliance", "Workflow"}
    for card, name in zip(
        page_cards,
        (
            "Home",
            "Clients",
            "Configure",
            "Mail Input Setup",
            "Mail Preview",
            "Document Setup",
            "Document Preview",
            "Mail Setup",
            "Compliance",
            "Workflow",
        ),
        strict=True,
    ):
        card.data = name
        card.visible = name == active_page["name"]
        if name in navigable_pages:
            nav_buttons.append(
                ft.OutlinedButton(
                    name,
                    data=name,
                    disabled=name == active_page["name"],
                    on_click=lambda event: show_page(event.control.data),
                )
            )

    controls = [
        ft.Text("Mail Compliance Automation", size=28, weight=ft.FontWeight.BOLD),
        ft.Text(f"Local database: {controller.db_path}", selectable=True),
        ft.Row([selected_quarter, ft.OutlinedButton("Refresh", on_click=lambda _event: refresh())]),
        ft.Row(
            [
                ft.Column(controls=nav_buttons, width=180),
                ft.Container(content=ft.Stack(controls=page_cards), expand=True),
            ]
        ),
        status,
    ]
    return {
        "controls": controls,
        "refresh": refresh,
        "show_page": show_page,
        "selected_quarter": selected_quarter,
        "workbook_path": workbook_path,
        "document_template_paths": document_template_paths,
        "mail_subject": mail_subject,
        "mail_body": mail_body,
    }


def normalized_mail_provider_value(raw_provider: object | None = None) -> str:
    raw = str(raw_provider or "gmail_smtp").strip().lower()
    if raw in {"gmail_smtp", "webtel_smtp"}:
        return raw
    if "webtel" in raw:
        return "webtel_smtp"
    return "gmail_smtp"


def _smtp_password_key(provider: str, sender_email: str) -> str:
    sender = (sender_email or "default").strip().lower()
    safe_sender = "".join(char if char.isalnum() or char in {"@", ".", "_", "-"} else "_" for char in sender)
    return f"mail.smtp_password.{normalized_mail_provider_value(provider)}.{safe_sender}"


def _save_smtp_password(settings: SettingsService, provider: str, sender_email: str, password: str) -> None:
    settings.set_value(_smtp_password_key(provider, sender_email), password, secret=True)
    # Keep the legacy key populated so older app versions and existing installs continue to work.
    settings.set_value("mail.smtp_password", password, secret=True)


def _load_smtp_password(settings: SettingsService, provider: str, sender_email: str, default: str = "") -> str:
    scoped = settings.get_value(_smtp_password_key(provider, sender_email), "")
    if scoped:
        return scoped
    legacy = settings.get_value("mail.smtp_password", "")
    if legacy and sender_email:
        settings.set_value(_smtp_password_key(provider, sender_email), legacy, secret=True)
        return legacy
    return default


def _latest_quarter(quarters: list[ClientQuarter]) -> ClientQuarter | None:
    current = [quarter for quarter in quarters if quarter.current_quarter]
    return (current or quarters)[-1] if quarters else None


def _attach_file_picker(ft, page):
    if page is None:
        return None
    picker = ft.FilePicker()
    services = getattr(page, "services", None)
    if services is not None:
        services.append(picker)
    else:
        page.overlay.append(picker)
    return picker


def _client_cards(clients: list[Client], quarters: list[ClientQuarter], dashboard: DashboardService) -> list[dict]:
    cards: list[dict] = []
    for client in clients:
        if client.id is None:
            continue
        client_quarters = [quarter for quarter in quarters if quarter.client_id == client.id]
        current_quarter = _latest_quarter(client_quarters)
        summary = dashboard.compliance_summary(current_quarter.id) if current_quarter and current_quarter.id else None
        cards.append(
            {
                "client_id": client.id,
                "client_name": client.name,
                "client_type": client.client_type,
                "quarter_id": current_quarter.id if current_quarter else None,
                "quarter_label": f"{current_quarter.financial_year} {current_quarter.quarter}" if current_quarter else "No quarter",
                "summary": summary,
            }
        )
    return cards


def _metric_card(ft, label: str, value: int):
    return ft.Container(
        width=240,
        padding=18,
        border_radius=16,
        bgcolor=ft.Colors.BLUE_GREY_50,
        content=ft.Column([ft.Text(label, size=14, color=ft.Colors.BLUE_GREY_700), ft.Text(str(value), size=34, weight=ft.FontWeight.BOLD)]),
    )


def _section(ft, title: str, content):
    return ft.Container(
        padding=16,
        border_radius=16,
        bgcolor=ft.Colors.WHITE,
        border=_border_all(ft, 1, ft.Colors.BLUE_GREY_100),
        content=ft.Column([ft.Text(title, size=20, weight=ft.FontWeight.BOLD), content]),
    )


def _quarter_label(quarter: ClientQuarter, client_lookup: dict[int | None, str]) -> str:
    client_name = client_lookup.get(quarter.client_id, "Client")
    current = "current" if quarter.current_quarter else quarter.status
    return f"{client_name} | {quarter.financial_year} {quarter.quarter} | {current}"


def _client_card_controls(ft, cards: list[dict], open_compliance, open_settings, run_action) -> list:
    if not cards:
        return [ft.Text("No clients created yet.", selectable=True)]
    controls = []
    for card in cards:
        controls.append(
            ft.Container(
                padding=12,
                border_radius=12,
                bgcolor=ft.Colors.BLUE_GREY_50,
                content=ft.Column(
                    [
                        ft.Text(f"{card['client_name']} ({card['client_type']})", weight=ft.FontWeight.BOLD),
                        ft.Text(_client_summary_line(card), selectable=True),
                        ft.Row(
                            [
                                ft.FilledButton("Compliance", on_click=lambda _event, client_id=card["client_id"]: run_action(lambda: open_compliance(client_id))),
                                ft.OutlinedButton("Configure", on_click=lambda _event, client_id=card["client_id"]: run_action(lambda: open_settings(client_id))),
                            ]
                        ),
                    ]
                ),
            )
        )
    return controls


def _compliance_client_controls(ft, cards: list[dict], open_compliance, run_action) -> list:
    if not cards:
        return [ft.Text("No clients to summarize yet.", selectable=True)]
    controls = []
    for card in cards:
        controls.append(
            ft.Container(
                padding=12,
                border_radius=12,
                bgcolor=ft.Colors.WHITE,
                border=_border_all(ft, 1, ft.Colors.BLUE_GREY_100),
                content=ft.Row(
                    [
                        ft.Column([ft.Text(card["client_name"], weight=ft.FontWeight.BOLD), ft.Text(_client_summary_line(card), selectable=True)], expand=True),
                        ft.FilledButton("Current Quarter", on_click=lambda _event, client_id=card["client_id"]: run_action(lambda: open_compliance(client_id))),
                    ]
                ),
            )
        )
    return controls


def _client_summary_line(card: dict) -> str:
    summary = card["summary"]
    if summary is None:
        return f"{card['quarter_label']}: no imported counterparties yet."
    return (
        f"{card['quarter_label']}: {_client_compliance_status(summary)} client; "
        f"{summary.fully_compliant}/{summary.total} counterparties compliant, {summary.non_compliant} non-compliant."
    )


def _configured_document_rows(ft, rows: list[dict], open_preview, delete_template) -> list:
    if not rows:
        return [ft.Text("No document templates configured yet.", selectable=True)]
    controls = []
    for row in rows:
        variables = ", ".join(row.get("variables", [])) or "none"
        file_name = Path(row["file_path"]).name if row.get("file_path") else "inline template"
        label = f"{row['name']} ({file_name})"
        controls.append(
            ft.Container(
                padding=10,
                border_radius=12,
                border=_border_all(ft, 1, ft.Colors.BLUE_GREY_100),
                content=ft.Row(
                    [
                        ft.Column(
                            [
                                ft.Text(row["name"], weight=ft.FontWeight.BOLD, selectable=True),
                                ft.Text(f"{file_name} | Variables: {variables}", size=12, selectable=True),
                            ],
                            spacing=2,
                            expand=True,
                        ),
                        ft.OutlinedButton(
                            "Open",
                            on_click=lambda _event, template_id=row["id"], title=label: open_preview(int(template_id), title),
                        ),
                        ft.OutlinedButton(
                            "Delete",
                            on_click=lambda _event, template_id=row["id"]: delete_template(int(template_id)),
                        ),
                    ],
                ),
            )
        )
    return controls


def _client_compliance_status(summary) -> str:
    if summary.total == 0:
        return "non-compliant"
    if summary.fully_compliant == summary.total:
        return "compliant"
    if summary.fully_compliant > 0:
        return "partially compliant"
    return "non-compliant"


def _counterparty_detail_controls(ft, rows: list[dict], selected_ids: set[int], on_toggle, on_doc_preview) -> list:
    if not rows:
        return [ft.Text("No Excel rows have been imported for this quarter yet.", selectable=True)]
    table = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text("Select")),
            ft.DataColumn(ft.Text("Creditor/Debtor")),
            ft.DataColumn(ft.Text("Email")),
            ft.DataColumn(ft.Text("Balance")),
            ft.DataColumn(ft.Text("Compliance")),
            ft.DataColumn(ft.Text("Generated Docs")),
            ft.DataColumn(ft.Text("Mail Sent")),
        ],
        rows=[
            ft.DataRow(
                cells=[
                    ft.DataCell(
                        ft.Checkbox(
                            value=row["counterparty_id"] in selected_ids,
                            data=row["counterparty_id"],
                            on_change=lambda event: on_toggle(int(event.control.data), bool(event.control.value)),
                        )
                    ),
                    ft.DataCell(
                        ft.Column(
                            [
                                ft.Text(row["party_name"] or "(blank)", weight=ft.FontWeight.BOLD, selectable=True),
                                ft.Text(f"{row['party_type']} | {row['sheet']}", size=12, selectable=True),
                            ],
                            spacing=2,
                        )
                    ),
                    ft.DataCell(ft.Text(row["email"], selectable=True)),
                    ft.DataCell(ft.Text(row["balance"], selectable=True)),
                    ft.DataCell(ft.Text(row["compliance_status"], selectable=True)),
                    ft.DataCell(_generated_document_cell(ft, row, on_doc_preview)),
                    ft.DataCell(ft.Text(row["mail_sent"], selectable=True)),
                ]
            )
            for row in rows
        ],
    )
    return [
        ft.Row([table], scroll=ft.ScrollMode.AUTO),
    ]


def _generated_document_cell(ft, row: dict, on_doc_preview):
    documents = row.get("documents", [])
    if not documents:
        return ft.Text(row["document_status"], selectable=True)

    first_document = documents[0]
    return ft.Container(
        content=ft.Text(row["document_status"], selectable=True),
        padding=8,
        border_radius=12,
        bgcolor=ft.Colors.BLUE_GREY_50,
        tooltip="Hover or click to preview generated documents",
        data=(row["counterparty_id"], first_document["id"]),
        on_hover=lambda event: _preview_document_on_hover(event, on_doc_preview),
        on_click=lambda event: on_doc_preview(int(event.control.data[0]), int(event.control.data[1])),
    )


def _preview_document_on_hover(event, on_doc_preview) -> None:
    if str(getattr(event, "data", "")).lower() != "true":
        return
    on_doc_preview(int(event.control.data[0]), int(event.control.data[1]))


def _short_document_label(document: dict) -> str:
    name = str(document.get("file_name") or Path(str(document.get("file_path", ""))).name or "document")
    return name if len(name) <= 28 else f"{name[:25]}..."


def _preview_bubble_width(page) -> int:
    raw_width = getattr(page, "window_width", None) or getattr(page, "width", None) or 1280
    try:
        available_width = int(raw_width) - 260
    except (TypeError, ValueError):
        available_width = 1020
    return max(420, min(720, available_width - 80))


def _list_texts(ft, values: list[str], empty_text: str) -> list:
    rows = values or [empty_text]
    return [ft.Text(value, selectable=True) for value in rows]


def _count_summary(values: dict[str, int]) -> str:
    if not values:
        return "none"
    return ", ".join(f"{key}={value}" for key, value in sorted(values.items()))


def _document_status(jobs: list[DocumentJob], documents: list[GeneratedDocument]) -> str:
    if documents:
        return f"generated ({len(documents)})"
    if jobs:
        return _count_summary({status: sum(1 for job in jobs if job.status == status) for status in {job.status for job in jobs}})
    return "not generated"


def _generated_document_summary(document: GeneratedDocument) -> dict:
    path = Path(document.file_path)
    return {
        "id": document.id,
        "file_path": document.file_path,
        "file_name": path.name,
        "file_type": document.file_type,
        "created_at": document.created_at,
    }


def _mapping_errors_for_variables(variables: list[str], mappings: dict[str, object], available_columns: set[str]) -> list[str]:
    builtin_variables = {
        "party_name",
        "email",
        "party_type",
        "balance",
        "row.party_name",
        "row.email",
        "row.party_type",
        "row.balance",
        "quarter.name",
        "quarter.financial_year",
        "quarter.quarter",
    }
    errors: list[str] = []
    for variable in sorted(set(variables)):
        if variable in builtin_variables:
            continue
        mapping = mappings.get(variable)
        if mapping is None:
            errors.append(f"{variable}: missing mapping")
            continue
        source_type = getattr(mapping, "source_type", "")
        source_key = getattr(mapping, "source_key", "")
        constant_value = getattr(mapping, "constant_value", "")
        if source_type == "excel_column" and source_key not in available_columns:
            errors.append(f"{variable}: Excel column not found: {source_key}")
        elif source_type == "constant" and not constant_value:
            errors.append(f"{variable}: constant value is empty")
    return errors


def _counterparty_to_confirmation_row(counterparty: Counterparty, subject: str, body: str):
    from .models import ConfirmationRow, RowState

    return ConfirmationRow(
        row_id=str(counterparty.id or "test"),
        excel_row_number=counterparty.source_row_number,
        party=counterparty.party_type or "",
        party_name=counterparty.party_name or "Test",
        contact_name=counterparty.party_name or "Test",
        contact_first_name=(counterparty.party_name or "Test").split(" ")[0],
        contact_last_name=" ".join((counterparty.party_name or "Test").split(" ")[1:]),
        email=counterparty.email,
        cc=counterparty.cc or "",
        subject=subject,
        balance=counterparty.balance or "",
        balance_nature="",
        company_name="",
        address="",
        phone="",
        balance_as_on_date="",
        letter_date="",
        auditor_reply_email="",
        mail_body_override=body,
        extra_attachment_paths=[],
        state=RowState(ready_to_send="Y", verification_status="Passed"),
    )


def _normalize_user_path(raw: str) -> Path:
    text = (raw or "").strip().strip('"')
    return Path(text).expanduser().resolve(strict=False)


def _parse_template_paths(raw_paths: str) -> list[Path]:
    return [Path(value) for value in _parse_template_path_strings(raw_paths)]


def _parse_template_path_strings(raw_paths: str) -> list[str]:
    values: list[str] = []
    for line in raw_paths.replace(";", "\n").splitlines():
        value = line.strip().strip('"')
        if value:
            values.append(value)
    return values


def _border_all(ft, width: int, color):
    side = ft.BorderSide(width=width, color=color)
    return ft.Border(top=side, right=side, bottom=side, left=side)


def _best_column_match(variable: str, columns: set[str]) -> str:
    normalized_variable = variable.lower().replace("row.", "").replace("_", " ").strip()
    for column in columns:
        if column.lower() == normalized_variable:
            return column
    for column in columns:
        simplified = column.lower().replace("(address)", "").replace("_", " ").strip()
        if simplified == normalized_variable or normalized_variable in simplified:
            return column
    aliases = {
        "party name": "Party Name",
        "party_name": "Party Name",
        "email": "Email To(Address)",
        "balance": "Balance",
        "party type": "Party Type",
    }
    candidate = aliases.get(normalized_variable)
    return candidate if candidate in columns else ""
