from __future__ import annotations

from collections import Counter
from datetime import datetime
from pathlib import Path

from sqlalchemy import delete
from sqlmodel import Session, select

from .db_models import (
    AuditEvent,
    CAFirm,
    Client,
    ClientQuarter,
    Counterparty,
    CounterpartyField,
    DocumentJob,
    EmailBatch,
    EmailMessage,
    ExcelImport,
    GeneratedDocument,
    Template,
    TemplateVariable,
    VariableMapping,
    WorkflowRun,
    utc_now_text,
)


class ClientDAO:
    def __init__(self, session: Session) -> None:
        self.session = session

    def create_client(self, name: str, client_type: str = "listed_org", ca_firm_id: int | None = None, **metadata) -> Client:
        client = Client(name=name, client_type=client_type, ca_firm_id=ca_firm_id, metadata_json=metadata)
        self.session.add(client)
        self.session.commit()
        self.session.refresh(client)
        return client

    def list_clients(self) -> list[Client]:
        return list(self.session.exec(select(Client).order_by(Client.name)))

    def get_client(self, client_id: int) -> Client:
        client = self.session.get(Client, client_id)
        if client is None:
            raise ValueError(f"Client not found: {client_id}")
        return client

    def create_quarter(self, client_id: int, financial_year: str, quarter: str, current: bool = True) -> ClientQuarter:
        if current:
            for existing in self.session.exec(select(ClientQuarter).where(ClientQuarter.client_id == client_id)):
                existing.current_quarter = False
                existing.updated_at = utc_now_text()
        client_quarter = ClientQuarter(
            client_id=client_id,
            financial_year=financial_year,
            quarter=quarter,
            current_quarter=current,
        )
        self.session.add(client_quarter)
        self.session.commit()
        self.session.refresh(client_quarter)
        return client_quarter

    def list_quarters(self, client_id: int | None = None) -> list[ClientQuarter]:
        statement = select(ClientQuarter).order_by(ClientQuarter.financial_year, ClientQuarter.quarter)
        if client_id is not None:
            statement = statement.where(ClientQuarter.client_id == client_id)
        return list(self.session.exec(statement))


class ImportDAO:
    def __init__(self, session: Session) -> None:
        self.session = session

    def create_import(
        self,
        client_quarter_id: int,
        source_path: Path,
        workbook_hash: str,
        sheet_metadata: dict,
        detected_columns: list[str],
    ) -> ExcelImport:
        excel_import = ExcelImport(
            client_quarter_id=client_quarter_id,
            source_path=str(source_path),
            workbook_hash=workbook_hash,
            sheet_metadata=sheet_metadata,
            detected_columns=detected_columns,
        )
        self.session.add(excel_import)
        self.session.commit()
        self.session.refresh(excel_import)
        return excel_import

    def upsert_counterparty(
        self,
        client_id: int,
        client_quarter_id: int,
        excel_import_id: int,
        source_sheet: str,
        source_row_number: int,
        source_row_key: str,
        values: dict[str, str],
    ) -> Counterparty:
        statement = select(Counterparty).where(
            Counterparty.client_quarter_id == client_quarter_id,
            Counterparty.source_row_key == source_row_key,
        )
        counterparty = self.session.exec(statement).first()
        if counterparty is None:
            counterparty = Counterparty(
                client_id=client_id,
                client_quarter_id=client_quarter_id,
                excel_import_id=excel_import_id,
                source_sheet=source_sheet,
                source_row_number=source_row_number,
                source_row_key=source_row_key,
                party_type=values.get("Party Type", ""),
                party_name=values.get("Party Name", ""),
                email=values.get("Email To(Address)", ""),
                balance=values.get("Balance", ""),
                status=values.get("Imported Compliance Status", "non_compliant"),
            )
            self.session.add(counterparty)
            self.session.flush()
        else:
            counterparty.excel_import_id = excel_import_id
            counterparty.party_type = values.get("Party Type", counterparty.party_type)
            counterparty.party_name = values.get("Party Name", counterparty.party_name)
            counterparty.email = values.get("Email To(Address)", counterparty.email)
            counterparty.balance = values.get("Balance", counterparty.balance)
            counterparty.status = values.get("Imported Compliance Status", "non_compliant")
            counterparty.updated_at = utc_now_text()

        assert counterparty.id is not None
        self.session.exec(delete(CounterpartyField).where(CounterpartyField.counterparty_id == counterparty.id))
        for key, value in values.items():
            self.session.add(CounterpartyField(counterparty_id=counterparty.id, field_name=key, field_value=value))
        self.session.commit()
        self.session.refresh(counterparty)
        return counterparty

    def list_counterparties(self, client_quarter_id: int) -> list[Counterparty]:
        statement = select(Counterparty).where(Counterparty.client_quarter_id == client_quarter_id).order_by(Counterparty.party_name)
        return [
            row
            for row in self.session.exec(statement)
            if row.party_name or row.email or row.balance
        ]

    def fields_for_counterparty(self, counterparty_id: int) -> dict[str, str]:
        rows = self.session.exec(select(CounterpartyField).where(CounterpartyField.counterparty_id == counterparty_id))
        return {row.field_name: row.field_value for row in rows}


class TemplateDAO:
    def __init__(self, session: Session) -> None:
        self.session = session

    def create_template(
        self,
        template_type: str,
        name: str,
        version: str,
        checksum: str,
        client_id: int | None = None,
        client_quarter_id: int | None = None,
        file_path: str = "",
        content_text: str = "",
        variables: set[str] | None = None,
    ) -> Template:
        template = Template(
            template_type=template_type,
            name=name,
            version=version,
            checksum=checksum,
            client_id=client_id,
            client_quarter_id=client_quarter_id,
            file_path=file_path,
            content_text=content_text,
        )
        self.session.add(template)
        self.session.flush()
        assert template.id is not None
        for variable in sorted(variables or set()):
            self.session.add(TemplateVariable(template_id=template.id, variable_name=variable))
        self.session.commit()
        self.session.refresh(template)
        return template

    def list_templates(self, client_quarter_id: int | None = None) -> list[Template]:
        statement = select(Template).where(Template.is_active == True).order_by(Template.created_at)  # noqa: E712
        if client_quarter_id is not None:
            statement = statement.where(Template.client_quarter_id == client_quarter_id)
        return list(self.session.exec(statement))

    def deactivate_templates(self, client_quarter_id: int, template_types: set[str]) -> None:
        for template in self.session.exec(
            select(Template).where(
                Template.client_quarter_id == client_quarter_id,
                Template.template_type.in_(template_types),
                Template.is_active == True,  # noqa: E712
            )
        ):
            template.is_active = False
            self.session.add(template)
        self.session.commit()

    def variables_for_quarter(self, client_quarter_id: int) -> set[str]:
        templates = self.list_templates(client_quarter_id)
        variables: set[str] = set()
        for template in templates:
            assert template.id is not None
            rows = self.session.exec(select(TemplateVariable).where(TemplateVariable.template_id == template.id))
            variables.update(row.variable_name for row in rows)
        return variables

    def save_mappings(self, client_quarter_id: int, mappings: dict[str, tuple[str, str, str]]) -> None:
        self.session.exec(delete(VariableMapping).where(VariableMapping.client_quarter_id == client_quarter_id))
        for variable, (source_type, source_key, constant_value) in mappings.items():
            self.session.add(
                VariableMapping(
                    client_quarter_id=client_quarter_id,
                    variable_name=variable,
                    source_type=source_type,
                    source_key=source_key,
                    constant_value=constant_value,
                )
            )
        self.session.commit()

    def mappings_for_quarter(self, client_quarter_id: int) -> dict[str, VariableMapping]:
        rows = self.session.exec(select(VariableMapping).where(VariableMapping.client_quarter_id == client_quarter_id))
        return {row.variable_name: row for row in rows}


class WorkflowDAO:
    def __init__(self, session: Session) -> None:
        self.session = session

    def create_run(self, client_quarter_id: int, run_type: str, total_count: int = 0, initiated_by: str = "") -> WorkflowRun:
        run = WorkflowRun(client_quarter_id=client_quarter_id, run_type=run_type, total_count=total_count, initiated_by=initiated_by)
        self.session.add(run)
        self.session.commit()
        self.session.refresh(run)
        return run

    def create_document_jobs(self, client_quarter_id: int, workflow_run_id: int | None = None) -> list[DocumentJob]:
        counterparties = list(self.session.exec(select(Counterparty).where(Counterparty.client_quarter_id == client_quarter_id)))
        templates = list(
            self.session.exec(
                select(Template).where(Template.client_quarter_id == client_quarter_id, Template.template_type == "document", Template.is_active == True)  # noqa: E712
            )
        )
        jobs: list[DocumentJob] = []
        for counterparty in counterparties:
            for template in templates:
                assert counterparty.id is not None
                assert template.id is not None
                existing = self.session.exec(
                    select(DocumentJob).where(
                        DocumentJob.client_quarter_id == client_quarter_id,
                        DocumentJob.counterparty_id == counterparty.id,
                        DocumentJob.template_id == template.id,
                    )
                ).first()
                if existing:
                    jobs.append(existing)
                    continue
                job = DocumentJob(
                    workflow_run_id=workflow_run_id,
                    client_quarter_id=client_quarter_id,
                    counterparty_id=counterparty.id,
                    template_id=template.id,
                )
                self.session.add(job)
                jobs.append(job)
        self.session.commit()
        return jobs

    def pending_document_jobs(self, client_quarter_id: int) -> list[DocumentJob]:
        statement = select(DocumentJob).where(
            DocumentJob.client_quarter_id == client_quarter_id,
            DocumentJob.status.in_(["pending", "failed", "retry_scheduled"]),
            DocumentJob.retry_locked == False,  # noqa: E712
        )
        return [job for job in self.session.exec(statement) if not job.next_retry_at or job.next_retry_at <= utc_now_text()]

    def mark_document_job(self, job_id: int, status: str, error: str = "") -> DocumentJob:
        job = self.session.get(DocumentJob, job_id)
        if job is None:
            raise ValueError(f"Document job not found: {job_id}")
        job.status = status
        job.error = error
        job.updated_at = utc_now_text()
        if status == "generated":
            job.retry_locked = False
        self.session.add(job)
        self.session.commit()
        self.session.refresh(job)
        return job

    def create_email_batch(self, client_quarter_id: int, batch_key: str, batch_size: int, workflow_run_id: int | None = None) -> EmailBatch:
        batch = EmailBatch(
            client_quarter_id=client_quarter_id,
            workflow_run_id=workflow_run_id,
            batch_key=batch_key,
            batch_size=batch_size,
        )
        self.session.add(batch)
        self.session.commit()
        self.session.refresh(batch)
        return batch

    def upsert_email_message(
        self,
        client_quarter_id: int,
        counterparty_id: int,
        to_email: str,
        subject: str,
        body: str,
        email_batch_id: int | None = None,
        status: str = "queued",
    ) -> EmailMessage:
        existing = self.session.exec(
            select(EmailMessage).where(
                EmailMessage.client_quarter_id == client_quarter_id,
                EmailMessage.counterparty_id == counterparty_id,
                EmailMessage.status != "sent",
            )
        ).first()
        message = existing or EmailMessage(
            client_quarter_id=client_quarter_id,
            counterparty_id=counterparty_id,
            to_email=to_email,
            subject=subject,
        )
        message.email_batch_id = email_batch_id
        message.body = body
        message.status = status
        if status == "queued":
            message.attempts = 0
            message.next_retry_at = ""
            message.retry_locked = False
            message.error = ""
        message.updated_at = utc_now_text()
        self.session.add(message)
        self.session.commit()
        self.session.refresh(message)
        return message

    def mark_email_sent(self, message_id: int, smtp_message_id: str, gmail_thread_id: str = "") -> EmailMessage:
        message = self.session.get(EmailMessage, message_id)
        if message is None:
            raise ValueError(f"Email message not found: {message_id}")
        message.status = "sent"
        message.smtp_message_id = smtp_message_id
        message.gmail_thread_id = gmail_thread_id
        message.sent_at = utc_now_text()
        message.error = ""
        self.session.add(message)
        self.session.commit()
        self.session.refresh(message)
        return message

    def generated_document_counts(self, client_quarter_id: int) -> Counter:
        rows = self.session.exec(select(DocumentJob.status).where(DocumentJob.client_quarter_id == client_quarter_id))
        return Counter(rows)

    def email_status_counts(self, client_quarter_id: int) -> Counter:
        rows = self.session.exec(select(EmailMessage.status).where(EmailMessage.client_quarter_id == client_quarter_id))
        return Counter(rows)


class AuditDAO:
    def __init__(self, session: Session) -> None:
        self.session = session

    def log(
        self,
        event_type: str,
        message: str,
        client_quarter_id: int | None = None,
        counterparty_id: int | None = None,
        entity_type: str = "",
        entity_id: str = "",
        details: dict | None = None,
    ) -> AuditEvent:
        event = AuditEvent(
            event_type=event_type,
            message=message,
            client_quarter_id=client_quarter_id,
            counterparty_id=counterparty_id,
            entity_type=entity_type,
            entity_id=entity_id,
            details=details or {},
        )
        self.session.add(event)
        self.session.commit()
        self.session.refresh(event)
        return event


def parse_datetime(value: str) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None
