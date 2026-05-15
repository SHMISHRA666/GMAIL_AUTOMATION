from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook
from sqlmodel import Session, select

from gmail_automation.dao import ImportDAO, TemplateDAO, WorkflowDAO
from gmail_automation.db import init_db
from gmail_automation.db_models import Counterparty, CounterpartyField, DocumentJob, EmailMessage, GeneratedDocument
from gmail_automation.docx_utils import build_docx_from_paragraphs, extract_docx_text
from gmail_automation.liquid_utils import extract_liquid_variables, render_liquid_template
from gmail_automation.modern_ui import ModernComplianceController, _attach_file_picker, _load_smtp_password, build_modern_ui_controls
from gmail_automation.services import ClientService, DashboardService, ImportService, TemplateService, WorkflowService
from gmail_automation.settings_store import SettingsService


def test_liquid_variable_extraction_and_strict_rendering() -> None:
    text = "Dear {{ party_name }}, balance {{ row.balance }} for {{ quarter.name }}"

    assert extract_liquid_variables(text) == {"party_name", "row.balance", "quarter.name"}
    rendered = render_liquid_template(
        text,
        {
            "party_name": "ABC Lender",
            "row": {"balance": "1000"},
            "quarter": {"name": "Q1"},
        },
    )

    assert rendered.text == "Dear ABC Lender, balance 1000 for Q1"


def test_excel_column_style_placeholder_rendering() -> None:
    text = "Dear {{Party Name}}, balance {{ Balance }} for {{ row.party_name }}"

    assert extract_liquid_variables(text) == {"Party Name", "Balance", "row.party_name"}
    rendered = render_liquid_template(
        text,
        {
            "Party Name": "ABC Lender",
            "Balance": "1000",
            "row": {"party_name": "ABC Lender"},
        },
    )

    assert rendered.text == "Dear ABC Lender, balance 1000 for ABC Lender"


def test_excel_import_preserves_standard_and_dynamic_fields(tmp_path: Path) -> None:
    engine = init_db(tmp_path / "compliance.db")
    workbook_path = _sample_workbook(tmp_path)

    with Session(engine) as session:
        client, quarter = _client_and_quarter(session)
        summary = ImportService(session).import_excel(client.id, quarter.id, workbook_path)
        rows = ImportDAO(session).list_counterparties(quarter.id)
        fields = session.exec(select(CounterpartyField).where(CounterpartyField.counterparty_id == rows[0].id)).all()

    assert summary.rows_imported == 2
    assert summary.sheet_counts == {"Creditors": 2}
    assert rows[0].party_name == "Alpha Finance"
    assert {field.field_name for field in fields} >= {"Party Name", "Email To(Address)", "Balance", "Custom Ref"}


def test_excel_import_status_columns_seed_compliance_summary(tmp_path: Path) -> None:
    engine = init_db(tmp_path / "compliance.db")
    workbook_path = _status_workbook(tmp_path)

    with Session(engine) as session:
        client, quarter = _client_and_quarter(session)
        ImportService(session).import_excel(client.id, quarter.id, workbook_path)
        rows = session.exec(select(Counterparty).where(Counterparty.client_quarter_id == quarter.id)).all()
        fields = {
            row.party_name: ImportDAO(session).fields_for_counterparty(row.id)
            for row in rows
            if row.id is not None
        }
        summary = DashboardService(session).compliance_summary(quarter.id)

    assert {row.party_name: row.status for row in rows} == {
        "Alpha Finance": "non_compliant",
        "Beta Bank": "non_compliant",
        "Gamma Capital": "non_compliant",
    }
    assert fields["Alpha Finance"]["Status"] == "Sent"
    assert fields["Beta Bank"]["ReadyToSend"] == "Y"
    assert fields["Gamma Capital"]["BounceReceived"] == "Yes"
    assert summary.fully_compliant == 0
    assert summary.partially_compliant == 0
    assert summary.non_compliant == 3
    assert summary.party_type_counts == {"Creditor": 3}


def test_excel_import_skips_readme_or_non_counterparty_sheets(tmp_path: Path) -> None:
    engine = init_db(tmp_path / "compliance.db")
    workbook_path = _workbook_with_readme_sheet(tmp_path)

    with Session(engine) as session:
        client, quarter = _client_and_quarter(session)
        summary = ImportService(session).import_excel(client.id, quarter.id, workbook_path)
        rows = ImportDAO(session).list_counterparties(quarter.id)

    assert summary.rows_imported == 2
    assert summary.sheet_counts == {"Tracking": 2}
    assert [row.party_name for row in rows] == ["Alpha Finance", "Beta Bank"]
    assert all(row.source_sheet == "Tracking" for row in rows)


def test_templates_mappings_generation_and_dashboard_queries(tmp_path: Path) -> None:
    engine = init_db(tmp_path / "compliance.db")
    workbook_path = _sample_workbook(tmp_path)

    with Session(engine) as session:
        client, quarter = _client_and_quarter(session)
        ImportService(session).import_excel(client.id, quarter.id, workbook_path)

        template_service = TemplateService(session)
        template_service.save_text_template(quarter.id, "document", "Confirmation Text", "Hello {{ party_name }}: {{ balance }}")
        template_service.save_text_template(quarter.id, "mail_subject", "Subject", "Confirm {{ party_name }}")
        template_service.save_text_template(quarter.id, "mail_body", "Body", "Balance is {{ balance }}")
        template_service.save_mappings(
            quarter.id,
            {
                "party_name": ("excel_column", "Party Name", ""),
                "balance": ("excel_column", "Balance", ""),
            },
        )

        assert template_service.validate_mappings(quarter.id, {"Party Name", "Balance"}) == []
        assert TemplateDAO(session).variables_for_quarter(quarter.id) == {"party_name", "balance"}

        workflow = WorkflowService(session, output_root=tmp_path / "generated")
        jobs = workflow.enqueue_document_generation(quarter.id)
        generated = workflow.generate_pending_documents(quarter.id)
        messages = workflow.queue_email_messages(quarter.id, "Confirm {{ party_name }}", "Balance is {{ balance }}")
        sent = workflow.mark_preview_batch_sent(quarter.id)
        summary = DashboardService(session).compliance_summary(quarter.id)
        job_counts = WorkflowDAO(session).generated_document_counts(quarter.id)
        email_counts = WorkflowDAO(session).email_status_counts(quarter.id)

        persisted_jobs = session.exec(select(DocumentJob).where(DocumentJob.client_quarter_id == quarter.id)).all()
        persisted_messages = session.exec(select(EmailMessage).where(EmailMessage.client_quarter_id == quarter.id)).all()
        persisted_counterparties = session.exec(select(Counterparty).where(Counterparty.client_quarter_id == quarter.id)).all()

    assert len(jobs) == 2
    assert len(generated) == 2
    assert all(result.status == "generated" for result in generated)
    assert len(messages) == 2
    assert sent == 2
    assert summary.fully_compliant == 2
    assert summary.non_compliant == 0
    assert job_counts["generated"] == 2
    assert email_counts["sent"] == 2
    assert {job.status for job in persisted_jobs} == {"generated"}
    assert {message.status for message in persisted_messages} == {"sent"}
    assert {counterparty.status for counterparty in persisted_counterparties} == {"compliant"}


def test_document_template_files_generate_for_all_counterparties(tmp_path: Path) -> None:
    engine = init_db(tmp_path / "compliance.db")
    workbook_path = _sample_workbook(tmp_path)
    letter_path = tmp_path / "letter.txt"
    annexure_path = tmp_path / "annexure.txt"
    letter_path.write_text("Letter for {{ party_name }}: {{ balance }}", encoding="utf-8")
    annexure_path.write_text("Annexure for {{ row.party_name }}", encoding="utf-8")

    with Session(engine) as session:
        client, quarter = _client_and_quarter(session)
        ImportService(session).import_excel(client.id, quarter.id, workbook_path)

        template_service = TemplateService(session)
        templates = template_service.save_document_templates(
            quarter.id,
            [letter_path, annexure_path],
            storage_dir=tmp_path / "template_store",
            client_id=client.id,
        )
        template_service.save_mappings(
            quarter.id,
            {
                "party_name": ("excel_column", "Party Name", ""),
                "balance": ("excel_column", "Balance", ""),
                "row.party_name": ("counterparty_field", "party_name", ""),
            },
        )

        workflow = WorkflowService(session, output_root=tmp_path / "generated")
        jobs = workflow.enqueue_document_generation(quarter.id)
        generated = workflow.generate_pending_documents(quarter.id)
        documents = session.exec(select(GeneratedDocument)).all()
        variables = TemplateDAO(session).variables_for_quarter(quarter.id)

    assert len(templates) == 2
    assert variables == {"party_name", "balance", "row.party_name"}
    assert len(jobs) == 4
    assert len(generated) == 4
    assert all(result.status == "generated" for result in generated)
    assert len(documents) == 4
    assert any("Alpha Finance" in Path(result.file_path).read_text(encoding="utf-8") for result in generated)


def test_pdf_authorisation_attachment_variables_validate_against_excel_columns(tmp_path: Path) -> None:
    engine = init_db(tmp_path / "compliance.db")
    workbook_path = _sample_workbook(tmp_path)
    authorisation_pdf = tmp_path / "Authorisation for Direct Balance Confirmation.pdf"
    authorisation_pdf.write_bytes(b"%PDF-1.4\nAuthorization for {{Party Name}}\n%%EOF")

    with Session(engine) as session:
        client, quarter = _client_and_quarter(session)
        ImportService(session).import_excel(client.id, quarter.id, workbook_path)
        template_service = TemplateService(session)
        templates = template_service.save_document_templates(
            quarter.id,
            [authorisation_pdf],
            storage_dir=tmp_path / "template_store",
            client_id=client.id,
        )
        template_names = [template.name for template in templates]
        assert TemplateDAO(session).variables_for_quarter(quarter.id) == {"Party Name"}
        template_service.save_mappings(quarter.id, {"Party Name": ("excel_column", "Party Name", "")})
        mapping_errors = template_service.validate_mappings(quarter.id, {"Party Name"})
        workflow = WorkflowService(session, output_root=tmp_path / "generated")
        jobs = workflow.enqueue_document_generation(quarter.id)
        generated = workflow.generate_pending_documents(quarter.id)

    assert template_names == ["Authorisation for Direct Balance Confirmation"]
    assert mapping_errors == []
    assert len(jobs) == 2
    assert len(generated) == 2
    assert all(Path(result.file_path).read_bytes() == authorisation_pdf.read_bytes() for result in generated)


def test_modern_ui_controller_and_controls_smoke(tmp_path: Path) -> None:
    """Flet smoke test without a Page. Manual: run launch_modern_ui(), Configure → Browse… or Import Excel (opens native picker on desktop); status shows success or Error: …"""
    ft = pytest.importorskip("flet")
    controller = ModernComplianceController(tmp_path / "modern.db")
    try:
        created = controller.create_client_record("Purple United")
        ui = build_modern_ui_controls(ft, controller)

        ui["refresh"]()
        ui["show_page"]("Clients")
        cards = controller.client_cards()

        assert len(ui["controls"]) >= 4
        assert ui["selected_quarter"].value is None
        assert cards[0]["client_name"] == "Purple United"
        assert cards[0]["quarter_label"] == "No quarter"
        with pytest.raises(ValueError, match="Create a quarter"):
            controller.current_quarter_id_for_client(created["id"])

        controller.create_quarter(created["id"], "2026-27", "Q1")
        ui["refresh"]()
        assert controller.current_quarter_id_for_client(created["id"]) == int(ui["selected_quarter"].value)
    finally:
        controller.close()


def test_modern_controller_import_workbook_validates_path_and_client(tmp_path: Path) -> None:
    controller = ModernComplianceController(tmp_path / "modern.db")
    workbook_path = _sample_workbook(tmp_path)
    try:
        a = controller.create_client_record("Client A")
        controller.create_quarter(a["id"], "2026-27", "Q1")
        quarter_a = controller.current_quarter_id_for_client(a["id"])
        b = controller.create_client_record("Client B")
        controller.create_quarter(b["id"], "2026-27", "Q1")
        quarter_b = controller.current_quarter_id_for_client(b["id"])

        with pytest.raises(ValueError, match="Workbook not found"):
            controller.import_workbook(quarter_a, str(tmp_path / "missing.xlsx"), client_id=a["id"])
        with pytest.raises(ValueError, match="does not belong"):
            controller.import_workbook(quarter_b, str(workbook_path), client_id=a["id"])
        assert controller.import_workbook(quarter_a, str(workbook_path), client_id=a["id"]).startswith("Imported")
    finally:
        controller.close()


def test_modern_controller_import_reports_legacy_status_wiring(tmp_path: Path) -> None:
    controller = ModernComplianceController(tmp_path / "modern.db")
    workbook_path = _status_workbook(tmp_path)
    try:
        created = controller.create_client_record("Client A")
        controller.create_quarter(created["id"], "2026-27", "Q1")
        quarter_id = controller.current_quarter_id_for_client(created["id"])

        message = controller.import_workbook(quarter_id, str(workbook_path), client_id=created["id"])
        summary = controller.get_client_compliance_summary(created["id"])["summary"]

        assert "Legacy status columns wired: BounceReceived, ReadyToSend, Status." in message
        assert summary.fully_compliant == 0
        assert summary.partially_compliant == 0
        assert summary.non_compliant == 3
        assert summary.party_type_counts == {"Creditor": 3}
    finally:
        controller.close()


def test_modern_controller_excel_persistence_overwrites_stale_quarter_rows_and_workflow(tmp_path: Path) -> None:
    controller = ModernComplianceController(tmp_path / "modern.db")
    first_workbook = _sample_workbook(tmp_path)
    replacement_workbook = _single_row_workbook(tmp_path, "replacement.xlsx", "Gamma Capital", "gamma@example.com", "3000")
    letter_path = tmp_path / "letter.txt"
    letter_path.write_text("Letter for {{ party_name }}: {{ balance }}", encoding="utf-8")

    try:
        created = controller.create_client_record("Purple United")
        controller.create_quarter(created["id"], "2026-27", "Q1")
        quarter_id = controller.current_quarter_id_for_client(created["id"])
        controller.import_workbook(quarter_id, str(first_workbook), client_id=created["id"])
        controller.save_mail_templates(quarter_id, "Confirm {{ party_name }}", "Balance is {{ balance }}")
        controller.save_document_templates(quarter_id, str(letter_path))
        assert controller.auto_map_variables(quarter_id) == "Mappings valid."
        controller.run_generation(quarter_id)
        controller.queue_and_preview_send(quarter_id, "Confirm {{ party_name }}", "Balance is {{ balance }}")

        assert len(controller.quarter_counterparty_statuses(quarter_id)) == 2

        message = controller.import_workbook(quarter_id, str(replacement_workbook), client_id=created["id"])
        rows = controller.quarter_counterparty_statuses(quarter_id)

        assert message == "Imported 1 row(s) from 1 sheet(s)."
        assert [row["party_name"] for row in rows] == ["Gamma Capital"]
        assert rows[0]["document_status"] == "not generated"
        assert rows[0]["mail_status"] == "not generated"
    finally:
        controller.close()


def test_modern_ui_file_picker_attaches_to_flet_services() -> None:
    ft = pytest.importorskip("flet")

    class DummyPage:
        def __init__(self) -> None:
            self.services = []
            self.overlay = []

    page = DummyPage()
    picker = _attach_file_picker(ft, page)

    assert picker in page.services
    assert page.overlay == []


def test_modern_controller_configure_flow_import_templates_mapping_and_summary(tmp_path: Path) -> None:
    controller = ModernComplianceController(tmp_path / "modern.db")
    workbook_path = _sample_workbook(tmp_path)
    letter_path = tmp_path / "letter.txt"
    letter_path.write_text("Letter for {{ party_name }}: {{ balance }}", encoding="utf-8")

    try:
        created = controller.create_client_record("Purple United")
        assert controller.client_cards()[0]["quarter_id"] is None

        controller.create_quarter(created["id"], "2026-27", "Q1")
        quarter_id = controller.current_quarter_id_for_client(created["id"])
        import_message = controller.import_workbook(quarter_id, str(workbook_path))
        template_message = controller.save_mail_templates(quarter_id, "Confirm {{ party_name }}", "Balance is {{ balance }}")
        document_message = controller.save_document_templates(quarter_id, str(letter_path))
        mapping_message = controller.auto_map_variables(quarter_id)
        imported_summary = controller.get_client_compliance_summary(created["id"])

        generation_message = controller.run_generation(quarter_id)
        queue_message = controller.queue_and_preview_send(quarter_id, "Confirm {{ party_name }}", "Balance is {{ balance }}")
        completed_summary = controller.get_client_compliance_summary(created["id"])

        assert import_message == "Imported 2 row(s) from 1 sheet(s)."
        assert "party_name" in template_message
        assert "Saved 1 document template" in document_message
        assert mapping_message == "Mappings valid."
        assert imported_summary["summary"].total == 2
        assert imported_summary["summary"].non_compliant == 2
        assert generation_message == "Generated 2 document job(s); 0 need attention."
        assert queue_message == "Queued 2 email(s); marked 2 sent in preview mode."
        assert completed_summary["summary"].fully_compliant == 2
        assert completed_summary["summary"].email_counts["sent"] == 2
        assert completed_summary["summary"].generation_counts["generated"] == 2
    finally:
        controller.close()


def test_modern_controller_row_level_compliance_detail_tracks_workflow_steps(tmp_path: Path) -> None:
    controller = ModernComplianceController(tmp_path / "modern.db")
    workbook_path = _sample_workbook(tmp_path)
    letter_path = tmp_path / "letter.txt"
    letter_path.write_text("Letter for {{ party_name }}: {{ balance }}", encoding="utf-8")

    try:
        created = controller.create_client_record("Purple United")
        controller.create_quarter(created["id"], "2026-27", "Q1")
        quarter_id = controller.current_quarter_id_for_client(created["id"])

        controller.import_workbook(quarter_id, str(workbook_path), client_id=created["id"])
        imported_rows = controller.quarter_counterparty_statuses(quarter_id)

        assert [row["party_name"] for row in imported_rows] == ["Alpha Finance", "Beta Bank"]
        assert {row["document_status"] for row in imported_rows} == {"not generated"}
        assert {row["compliance_status"] for row in imported_rows} == {"non_compliant"}
        assert {row["mail_status"] for row in imported_rows} == {"not generated"}
        assert {row["mail_sent"] for row in imported_rows} == {"No"}

        controller.save_mail_templates(quarter_id, "Confirm {{ party_name }}", "Balance is {{ balance }}")
        controller.save_document_templates(quarter_id, str(letter_path))
        assert controller.auto_map_variables(quarter_id) == "Mappings valid."
        controller.run_generation(quarter_id)
        controller.queue_and_preview_send(quarter_id, "Confirm {{ party_name }}", "Balance is {{ balance }}")

        completed_rows = controller.quarter_counterparty_statuses(quarter_id)

        assert {row["document_status"] for row in completed_rows} == {"generated (1)"}
        assert {row["document_count"] for row in completed_rows} == {1}
        assert {row["compliance_status"] for row in completed_rows} == {"compliant"}
        assert {row["mail_status"] for row in completed_rows} == {"sent"}
        assert {row["mail_sent"] for row in completed_rows} == {"Yes"}
    finally:
        controller.close()


def test_modern_controller_selected_workflow_retrigger_affects_only_selected_rows(tmp_path: Path) -> None:
    controller = ModernComplianceController(tmp_path / "modern.db")
    workbook_path = _sample_workbook(tmp_path)
    letter_path = tmp_path / "letter.txt"
    letter_path.write_text("Letter for {{ party_name }}: {{ balance }}", encoding="utf-8")

    try:
        created = controller.create_client_record("Purple United")
        controller.create_quarter(created["id"], "2026-27", "Q1")
        quarter_id = controller.current_quarter_id_for_client(created["id"])
        controller.import_workbook(quarter_id, str(workbook_path), client_id=created["id"])
        controller.save_mail_templates(quarter_id, "Confirm {{ party_name }}", "Balance is {{ balance }}")
        controller.save_document_templates(quarter_id, str(letter_path))
        assert controller.auto_map_variables(quarter_id) == "Mappings valid."

        rows = controller.quarter_counterparty_statuses(quarter_id)
        alpha_id = next(row["counterparty_id"] for row in rows if row["party_name"] == "Alpha Finance")

        with pytest.raises(ValueError, match="Generate documents"):
            controller.queue_and_preview_send_selected(quarter_id, {alpha_id}, "Confirm {{ party_name }}", "Balance is {{ balance }}")
        assert controller.regenerate_documents(quarter_id, {alpha_id}) == "Regenerated 1 selected document job(s); 0 need attention."
        with Session(controller.engine) as session:
            session.add(
                EmailMessage(
                    client_quarter_id=quarter_id,
                    counterparty_id=alpha_id,
                    to_email="alpha@example.com",
                    subject="Old failed message",
                    status="failed",
                    attempts=4,
                    retry_locked=True,
                    next_retry_at="2099-01-01 00:00:00",
                    error="old authentication failure",
                )
            )
            session.commit()
        assert controller.queue_and_preview_send_selected(quarter_id, {alpha_id}, "Confirm {{ party_name }}", "Balance is {{ balance }}") == (
            "Queued 1 selected email(s); marked 1 sent in preview mode."
        )

        completed_rows = {row["party_name"]: row for row in controller.quarter_counterparty_statuses(quarter_id)}
        with Session(controller.engine) as session:
            alpha_messages = list(session.exec(select(EmailMessage).where(EmailMessage.counterparty_id == alpha_id)))

        assert completed_rows["Alpha Finance"]["document_status"] == "generated (1)"
        assert completed_rows["Alpha Finance"]["compliance_status"] == "compliant"
        assert completed_rows["Alpha Finance"]["mail_status"] == "sent"
        assert completed_rows["Alpha Finance"]["mail_sent"] == "Yes"
        assert {message.retry_locked for message in alpha_messages} == {False}
        assert {message.error for message in alpha_messages} == {""}
        assert completed_rows["Beta Bank"]["document_status"] == "not generated"
        assert completed_rows["Beta Bank"]["compliance_status"] == "non_compliant"
        assert completed_rows["Beta Bank"]["mail_status"] == "not generated"
        assert completed_rows["Beta Bank"]["mail_sent"] == "No"
    finally:
        controller.close()


def test_modern_controller_doc_regeneration_readiness_not_blocked_by_mail_only_mapping_errors(tmp_path: Path) -> None:
    controller = ModernComplianceController(tmp_path / "modern.db")
    workbook_path = _sample_workbook(tmp_path)
    letter_path = tmp_path / "letter.txt"
    letter_path.write_text("Letter for {{ party_name }}", encoding="utf-8")

    try:
        created = controller.create_client_record("Purple United")
        controller.create_quarter(created["id"], "2026-27", "Q1")
        quarter_id = controller.current_quarter_id_for_client(created["id"])
        controller.import_workbook(quarter_id, str(workbook_path), client_id=created["id"])
        controller.save_document_templates(quarter_id, str(letter_path))
        controller.save_mail_templates(quarter_id, "Confirm {{ party_name }}", "Balance for {{ quarter.name }}")
        controller.auto_map_variables(quarter_id)
        readiness = controller.quarter_workflow_readiness(quarter_id)
        row_ids = {row["counterparty_id"] for row in controller.quarter_counterparty_statuses(quarter_id)}

        assert readiness["can_generate_documents"] is True
        assert readiness["mail_mapping_errors"] == []
        assert readiness["can_send_mail"] is readiness["send_enabled"]
        assert controller.regenerate_documents(quarter_id, row_ids).startswith("Regenerated")
    finally:
        controller.close()


def test_modern_controller_send_mail_selected_requires_send_mode_and_credentials(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.chdir(tmp_path)
    controller = ModernComplianceController(tmp_path / "modern.db")
    workbook_path = _sample_workbook(tmp_path)
    letter_path = tmp_path / "letter.txt"
    letter_path.write_text("Letter for {{ party_name }}", encoding="utf-8")

    try:
        created = controller.create_client_record("Purple United")
        controller.create_quarter(created["id"], "2026-27", "Q1")
        quarter_id = controller.current_quarter_id_for_client(created["id"])
        controller.import_workbook(quarter_id, str(workbook_path), client_id=created["id"])
        controller.save_document_templates(quarter_id, str(letter_path))
        controller.save_mail_templates(quarter_id, "Confirm {{ party_name }}", "Balance is {{ balance }}")
        controller.auto_map_variables(quarter_id)
        row_ids = {row["counterparty_id"] for row in controller.quarter_counterparty_statuses(quarter_id)}
        controller.regenerate_documents(quarter_id, row_ids)

        with pytest.raises(ValueError, match="Enable send mode"):
            controller.send_mail_selected(quarter_id, row_ids, "Confirm {{ party_name }}", "Balance is {{ balance }}")
    finally:
        controller.close()


def test_modern_controller_mail_builtin_variables_do_not_require_mapping(tmp_path: Path) -> None:
    controller = ModernComplianceController(tmp_path / "modern.db")
    workbook_path = _sample_workbook(tmp_path)
    letter_path = tmp_path / "letter.txt"
    letter_path.write_text("Letter for {{ party_name }}", encoding="utf-8")

    try:
        created = controller.create_client_record("Purple United")
        controller.create_quarter(created["id"], "2026-27", "Q1")
        quarter_id = controller.current_quarter_id_for_client(created["id"])
        controller.import_workbook(quarter_id, str(workbook_path), client_id=created["id"])
        controller.save_document_templates(quarter_id, str(letter_path))
        controller.save_mail_templates(quarter_id, "Confirm {{ party_name }}", "Balance {{ balance }} for {{ quarter.name }}")

        readiness = controller.quarter_workflow_readiness(quarter_id)

        assert readiness["mail_mapping_errors"] == []
    finally:
        controller.close()


def test_modern_controller_mail_settings_roundtrip(tmp_path: Path) -> None:
    controller = ModernComplianceController(tmp_path / "modern.db")
    try:
        message = controller.save_mail_settings(
            mail_provider="gmail_smtp",
            send_mode="send",
            sender_email="sender@example.com",
            fallback_providers="webtel_smtp",
            daily_send_limit=500,
            per_email_delay_seconds=2,
            smtp_host="smtp.gmail.com",
            smtp_port=587,
            smtp_username="sender@example.com",
            smtp_password="secret",
        )
        settings = controller.get_mail_settings()

        assert message == "Mail settings saved."
        assert settings["mail_provider"] == "gmail_smtp"
        assert settings["send_mode"] == "send"
        assert settings["sender_email"] == "sender@example.com"
        assert settings["fallback_providers"] == "webtel_smtp"
        assert settings["daily_send_limit"] == "500"
        assert settings["smtp_password_saved"] is True
        with Session(controller.engine) as session:
            assert _load_smtp_password(SettingsService(session), "gmail_smtp", "sender@example.com") == "secret"
    finally:
        controller.close()


def test_modern_controller_generated_document_preview_for_selected_counterparty(tmp_path: Path) -> None:
    controller = ModernComplianceController(tmp_path / "modern.db")
    workbook_path = _sample_workbook(tmp_path)
    letter_path = tmp_path / "letter.txt"
    letter_path.write_text("Letter for {{ party_name }}: {{ balance }}", encoding="utf-8")
    note_path = tmp_path / "note.txt"
    note_path.write_text("Note for {{ party_name }}: {{ balance }}", encoding="utf-8")

    try:
        created = controller.create_client_record("Purple United")
        controller.create_quarter(created["id"], "2026-27", "Q1")
        quarter_id = controller.current_quarter_id_for_client(created["id"])
        controller.import_workbook(quarter_id, str(workbook_path), client_id=created["id"])
        controller.save_document_templates(quarter_id, f"{letter_path}\n{note_path}")
        assert controller.auto_map_variables(quarter_id) == "Mappings valid."
        assert controller.regenerate_documents(quarter_id, {row["counterparty_id"] for row in controller.quarter_counterparty_statuses(quarter_id)}).startswith(
            "Regenerated"
        )

        rows = controller.quarter_counterparty_statuses(quarter_id)
        alpha_id = next(row["counterparty_id"] for row in rows if row["party_name"] == "Alpha Finance")
        generated_docs = controller.generated_documents_for_counterparty(quarter_id, alpha_id)
        preview = controller.preview_generated_document(generated_docs[0]["id"])

        assert len(generated_docs) == 2
        assert next(row for row in rows if row["party_name"] == "Alpha Finance")["document_count"] == 2
        assert len(next(row for row in rows if row["party_name"] == "Alpha Finance")["documents"]) == 2
        assert generated_docs[0]["file_path"].endswith(".txt")
        assert "Alpha Finance" in preview
        assert "1000" in preview
        assert "Alpha Finance" in controller.preview_generated_document(generated_docs[1]["id"])
    finally:
        controller.close()


def test_modern_controller_uploaded_docx_and_static_attachments_generate_row_values(tmp_path: Path) -> None:
    controller = ModernComplianceController(tmp_path / "modern.db")
    workbook_path = _sample_workbook(tmp_path)
    letter_path = tmp_path / "templated_letter.docx"
    letter_path.write_bytes(
        build_docx_from_paragraphs(
            [
                "Dear {{ party_name }}",
                "Balance {{ balance }} for {{ row.party_name }}",
            ]
        )
    )
    note_path = tmp_path / "templated_note.txt"
    note_path.write_text("Note for {{ party_name }}: {{ balance }}", encoding="utf-8")
    static_pdf_path = tmp_path / "authorisation.pdf"
    static_pdf_path.write_bytes(b"%PDF-1.4 static attachment")

    try:
        created = controller.create_client_record("Purple United")
        controller.create_quarter(created["id"], "2026-27", "Q1")
        quarter_id = controller.current_quarter_id_for_client(created["id"])

        controller.import_workbook(quarter_id, str(workbook_path), client_id=created["id"])
        controller.save_document_templates(quarter_id, f"{letter_path}\n{note_path}\n{static_pdf_path}")
        assert controller.auto_map_variables(quarter_id) == "Mappings valid."
        assert controller.run_generation(quarter_id) == "Generated 6 document job(s); 0 need attention."

        with Session(controller.engine) as session:
            documents = session.exec(select(GeneratedDocument)).all()

        generated_suffixes = {Path(document.file_path).name.split("_", 1)[-1] for document in documents}
        letter_texts = [extract_docx_text(Path(document.file_path)) for document in documents if Path(document.file_path).name.endswith("templated_letter.docx")]
        note_texts = [Path(document.file_path).read_text(encoding="utf-8") for document in documents if Path(document.file_path).name.endswith("templated_note.txt")]
        pdf_paths = [Path(document.file_path) for document in documents if Path(document.file_path).name.endswith("authorisation.pdf")]

        assert generated_suffixes == {"templated_letter.docx", "templated_note.txt", "authorisation.pdf"}
        assert any("Dear Alpha Finance" in text and "Balance 1000" in text for text in letter_texts)
        assert any("Dear Beta Bank" in text and "Balance 2000" in text for text in letter_texts)
        assert "Note for Alpha Finance: 1000" in note_texts
        assert "Note for Beta Bank: 2000" in note_texts
        assert len(pdf_paths) == 2
        assert all(path.read_bytes() == static_pdf_path.read_bytes() for path in pdf_paths)
    finally:
        controller.close()


def test_modern_controller_multi_document_save_validates_pdf_and_text_variables(tmp_path: Path) -> None:
    controller = ModernComplianceController(tmp_path / "modern.db")
    workbook_path = _sample_workbook(tmp_path)
    letter_path = tmp_path / "letter.txt"
    letter_path.write_text("Letter for {{ party_name }}", encoding="utf-8")
    authorisation_pdf = tmp_path / "Authorisation for Direct Balance Confirmation.pdf"
    authorisation_pdf.write_bytes(b"%PDF-1.4\nAuthorization for {{Party Name}}\n%%EOF")

    try:
        created = controller.create_client_record("Purple United")
        controller.create_quarter(created["id"], "2026-27", "Q1")
        quarter_id = controller.current_quarter_id_for_client(created["id"])
        controller.import_workbook(quarter_id, str(workbook_path), client_id=created["id"])

        message = controller.save_document_templates(quarter_id, f"{letter_path}\n{authorisation_pdf}")
        mapping = controller.auto_map_variables(quarter_id)
        readiness = controller.quarter_workflow_readiness(quarter_id)

        assert "Saved 2 document template" in message
        assert "Party Name" in message
        assert mapping == "Mappings valid."
        assert readiness["document_template_count"] == 2
        assert readiness["document_mapping_errors"] == []
        assert readiness["can_generate_documents"] is True
    finally:
        controller.close()


def _client_and_quarter(session: Session):
    service = ClientService(session)
    client = service.create_client("Purple United", "listed_org")
    assert client.id is not None
    quarter = service.create_quarter(client.id, "2026-27", "Q1", current=True)
    assert quarter.id is not None
    return client, quarter


def _sample_workbook(tmp_path: Path) -> Path:
    workbook_path = tmp_path / "master.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Creditors"
    ws.append(["S.No.", "Party Type", "Party Name", "Email To(Address)", "Balance", "Custom Ref"])
    ws.append([1, "Creditor", "Alpha Finance", "alpha@example.com", "1000", "A-1"])
    ws.append([2, "Creditor", "Beta Bank", "beta@example.com", "2000", "B-1"])
    wb.save(workbook_path)
    return workbook_path


def _single_row_workbook(tmp_path: Path, filename: str, party_name: str, email: str, balance: str) -> Path:
    workbook_path = tmp_path / filename
    wb = Workbook()
    ws = wb.active
    ws.title = "Tracking"
    ws.append(["S.No.", "Party Type", "Party Name", "Email To(Address)", "Balance"])
    ws.append([1, "Creditor", party_name, email, balance])
    wb.save(workbook_path)
    return workbook_path


def _status_workbook(tmp_path: Path) -> Path:
    workbook_path = tmp_path / "status_master.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Creditors"
    ws.append(
        [
            "S.No.",
            "Party Type",
            "Party Name",
            "Email To(Address)",
            "Balance",
            "Status",
            "ReadyToSend",
            "BounceReceived",
        ]
    )
    ws.append([1, "Creditor", "Alpha Finance", "alpha@example.com", "1000", "Sent", "", ""])
    ws.append([2, "Creditor", "Beta Bank", "beta@example.com", "2000", "", "Y", ""])
    ws.append([3, "Creditor", "Gamma Capital", "gamma@example.com", "3000", "", "", "Yes"])
    wb.save(workbook_path)
    return workbook_path


def _workbook_with_readme_sheet(tmp_path: Path) -> Path:
    workbook_path = tmp_path / "readme_and_tracking.xlsx"
    wb = Workbook()
    readme = wb.active
    readme.title = "README"
    readme.append(["Instructions"])
    readme.append(["Use the Tracking sheet for creditor/debtor data."])
    readme.append(["Do not import this sheet as workflow rows."])
    tracking = wb.create_sheet("Tracking")
    tracking.append(["S.No.", "Party Type", "Party Name", "Email To(Address)", "Balance"])
    tracking.append([1, "Creditor", "Alpha Finance", "alpha@example.com", "1000"])
    tracking.append([2, "Debtor", "Beta Bank", "beta@example.com", "2000"])
    wb.save(workbook_path)
    return workbook_path
