from __future__ import annotations

from datetime import UTC, datetime

from sqlalchemy import Column, UniqueConstraint
from sqlalchemy.types import JSON
from sqlmodel import Field, SQLModel


def utc_now_text() -> str:
    return datetime.now(UTC).strftime("%Y-%m-%d %H:%M:%S")


class CAFirm(SQLModel, table=True):
    __tablename__ = "ca_firms"

    id: int | None = Field(default=None, primary_key=True)
    name: str = Field(index=True)
    contact_email: str = ""
    phone: str = ""
    created_at: str = Field(default_factory=utc_now_text)
    updated_at: str = Field(default_factory=utc_now_text)


class Client(SQLModel, table=True):
    __tablename__ = "clients"

    id: int | None = Field(default=None, primary_key=True)
    ca_firm_id: int | None = Field(default=None, foreign_key="ca_firms.id", index=True)
    name: str = Field(index=True)
    client_type: str = Field(default="listed_org", index=True)
    cin: str = ""
    pan: str = ""
    gstin: str = ""
    contact_email: str = ""
    is_active: bool = True
    metadata_json: dict = Field(default_factory=dict, sa_column=Column(JSON))
    created_at: str = Field(default_factory=utc_now_text)
    updated_at: str = Field(default_factory=utc_now_text)


class ClientQuarter(SQLModel, table=True):
    __tablename__ = "client_quarters"
    __table_args__ = (UniqueConstraint("client_id", "financial_year", "quarter", name="uq_client_quarter"),)

    id: int | None = Field(default=None, primary_key=True)
    client_id: int = Field(foreign_key="clients.id", index=True)
    financial_year: str = Field(index=True)
    quarter: str = Field(index=True)
    due_date: str = ""
    status: str = Field(default="draft", index=True)
    current_quarter: bool = Field(default=False, index=True)
    created_at: str = Field(default_factory=utc_now_text)
    updated_at: str = Field(default_factory=utc_now_text)


class ExcelImport(SQLModel, table=True):
    __tablename__ = "excel_imports"

    id: int | None = Field(default=None, primary_key=True)
    client_quarter_id: int = Field(foreign_key="client_quarters.id", index=True)
    source_path: str
    workbook_hash: str = Field(index=True)
    sheet_metadata: dict = Field(default_factory=dict, sa_column=Column(JSON))
    detected_columns: list[str] = Field(default_factory=list, sa_column=Column(JSON))
    imported_at: str = Field(default_factory=utc_now_text)


class Counterparty(SQLModel, table=True):
    __tablename__ = "counterparties"
    __table_args__ = (UniqueConstraint("client_quarter_id", "source_row_key", name="uq_counterparty_source_row"),)

    id: int | None = Field(default=None, primary_key=True)
    client_id: int = Field(foreign_key="clients.id", index=True)
    client_quarter_id: int = Field(foreign_key="client_quarters.id", index=True)
    excel_import_id: int | None = Field(default=None, foreign_key="excel_imports.id", index=True)
    source_sheet: str = Field(index=True)
    source_row_number: int
    source_row_key: str = Field(index=True)
    party_type: str = Field(default="", index=True)
    party_name: str = Field(index=True)
    email: str = Field(index=True)
    cc: str = ""
    balance: str = ""
    status: str = Field(default="pending", index=True)
    created_at: str = Field(default_factory=utc_now_text)
    updated_at: str = Field(default_factory=utc_now_text)


class CounterpartyField(SQLModel, table=True):
    __tablename__ = "counterparty_fields"
    __table_args__ = (UniqueConstraint("counterparty_id", "field_name", name="uq_counterparty_field"),)

    id: int | None = Field(default=None, primary_key=True)
    counterparty_id: int = Field(foreign_key="counterparties.id", index=True)
    field_name: str = Field(index=True)
    field_value: str = ""


class Template(SQLModel, table=True):
    __tablename__ = "templates"

    id: int | None = Field(default=None, primary_key=True)
    client_id: int | None = Field(default=None, foreign_key="clients.id", index=True)
    client_quarter_id: int | None = Field(default=None, foreign_key="client_quarters.id", index=True)
    template_type: str = Field(index=True)
    name: str
    version: str = Field(index=True)
    file_path: str = ""
    content_text: str = ""
    checksum: str = Field(index=True)
    is_active: bool = True
    created_at: str = Field(default_factory=utc_now_text)


class TemplateVariable(SQLModel, table=True):
    __tablename__ = "template_variables"
    __table_args__ = (UniqueConstraint("template_id", "variable_name", name="uq_template_variable"),)

    id: int | None = Field(default=None, primary_key=True)
    template_id: int = Field(foreign_key="templates.id", index=True)
    variable_name: str = Field(index=True)
    required: bool = True
    source_hint: str = ""


class VariableMapping(SQLModel, table=True):
    __tablename__ = "variable_mappings"
    __table_args__ = (UniqueConstraint("client_quarter_id", "variable_name", name="uq_quarter_variable_mapping"),)

    id: int | None = Field(default=None, primary_key=True)
    client_quarter_id: int = Field(foreign_key="client_quarters.id", index=True)
    variable_name: str = Field(index=True)
    source_type: str = Field(index=True)
    source_key: str
    constant_value: str = ""
    created_at: str = Field(default_factory=utc_now_text)


class WorkflowRun(SQLModel, table=True):
    __tablename__ = "workflow_runs"

    id: int | None = Field(default=None, primary_key=True)
    client_quarter_id: int = Field(foreign_key="client_quarters.id", index=True)
    run_type: str = Field(index=True)
    status: str = Field(default="queued", index=True)
    total_count: int = 0
    success_count: int = 0
    failed_count: int = 0
    started_at: str = Field(default_factory=utc_now_text)
    finished_at: str = ""
    initiated_by: str = ""


class DocumentJob(SQLModel, table=True):
    __tablename__ = "document_jobs"
    __table_args__ = (UniqueConstraint("client_quarter_id", "counterparty_id", "template_id", name="uq_document_job"),)

    id: int | None = Field(default=None, primary_key=True)
    workflow_run_id: int | None = Field(default=None, foreign_key="workflow_runs.id", index=True)
    client_quarter_id: int = Field(foreign_key="client_quarters.id", index=True)
    counterparty_id: int = Field(foreign_key="counterparties.id", index=True)
    template_id: int | None = Field(default=None, foreign_key="templates.id", index=True)
    status: str = Field(default="pending", index=True)
    attempts: int = 0
    next_retry_at: str = ""
    retry_locked: bool = False
    error: str = ""
    created_at: str = Field(default_factory=utc_now_text)
    updated_at: str = Field(default_factory=utc_now_text)


class GeneratedDocument(SQLModel, table=True):
    __tablename__ = "generated_documents"

    id: int | None = Field(default=None, primary_key=True)
    document_job_id: int = Field(foreign_key="document_jobs.id", index=True)
    counterparty_id: int = Field(foreign_key="counterparties.id", index=True)
    template_id: int | None = Field(default=None, foreign_key="templates.id", index=True)
    file_type: str = Field(index=True)
    file_path: str
    checksum: str = Field(index=True)
    template_version: str = ""
    created_at: str = Field(default_factory=utc_now_text)


class EmailBatch(SQLModel, table=True):
    __tablename__ = "email_batches"

    id: int | None = Field(default=None, primary_key=True)
    client_quarter_id: int = Field(foreign_key="client_quarters.id", index=True)
    workflow_run_id: int | None = Field(default=None, foreign_key="workflow_runs.id", index=True)
    batch_key: str = Field(index=True)
    sequence_number: int = 0
    batch_size: int = 0
    status: str = Field(default="queued", index=True)
    daily_limit: int = 0
    per_email_delay_seconds: int = 0
    created_at: str = Field(default_factory=utc_now_text)
    completed_at: str = ""


class EmailMessage(SQLModel, table=True):
    __tablename__ = "email_messages"

    id: int | None = Field(default=None, primary_key=True)
    email_batch_id: int | None = Field(default=None, foreign_key="email_batches.id", index=True)
    client_quarter_id: int = Field(foreign_key="client_quarters.id", index=True)
    counterparty_id: int = Field(foreign_key="counterparties.id", index=True)
    to_email: str = Field(index=True)
    cc_email: str = ""
    subject: str
    body: str = ""
    status: str = Field(default="queued", index=True)
    attempts: int = 0
    next_retry_at: str = ""
    retry_locked: bool = False
    smtp_message_id: str = ""
    gmail_thread_id: str = ""
    sent_at: str = ""
    bounce_received: bool = False
    bounce_date: str = ""
    reply_received: bool = False
    reply_received_date: str = ""
    error: str = ""
    created_at: str = Field(default_factory=utc_now_text)
    updated_at: str = Field(default_factory=utc_now_text)


class AuditEvent(SQLModel, table=True):
    __tablename__ = "audit_events"

    id: int | None = Field(default=None, primary_key=True)
    client_quarter_id: int | None = Field(default=None, foreign_key="client_quarters.id", index=True)
    counterparty_id: int | None = Field(default=None, foreign_key="counterparties.id", index=True)
    event_type: str = Field(index=True)
    entity_type: str = ""
    entity_id: str = ""
    message: str
    details: dict = Field(default_factory=dict, sa_column=Column(JSON))
    created_at: str = Field(default_factory=utc_now_text)


class AppSetting(SQLModel, table=True):
    __tablename__ = "app_settings"

    key: str = Field(primary_key=True)
    value: str = ""
    is_secret: bool = False
    updated_at: str = Field(default_factory=utc_now_text)
