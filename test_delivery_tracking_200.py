from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import src.order_workflow as workflow
from src.order_workflow import ManualTransitOrder


def _use_temp_db(monkeypatch, tmp_path: Path) -> Path:
    database = tmp_path / "orders.sqlite3"
    monkeypatch.setattr(workflow, "DRAFT_DB", database)
    monkeypatch.setattr(workflow, "get_cloud_storage", lambda: None)
    return database


def test_local_completed_order_receipt_can_be_checked_and_reverted(monkeypatch, tmp_path):
    _use_temp_db(monkeypatch, tmp_path)

    first = workflow.set_order_received("report-hash", workflow.ORDER_MODE_STONES, True)
    assert first["received"] is True
    assert first["received_at"]
    stored = workflow._local_receipt_status("report-hash", workflow.ORDER_MODE_STONES)
    assert stored["received"] is True
    assert stored["received_at"]

    second = workflow.set_order_received("report-hash", workflow.ORDER_MODE_STONES, False)
    assert second["received"] is False
    assert second["received_at"] == ""
    stored = workflow._local_receipt_status("report-hash", workflow.ORDER_MODE_STONES)
    assert stored["received"] is False
    assert stored["received_at"] == ""


def test_manual_transit_order_lifecycle_is_persistent(monkeypatch, tmp_path):
    _use_temp_db(monkeypatch, tmp_path)

    created = workflow.save_manual_transit_order(
        ManualTransitOrder(
            order_id="manual-1",
            title="Жемчуг",
            order_date="2026-07-19",
            note="Ожидаем поставку",
        )
    )
    assert created.received is False
    assert created.title == "Жемчуг"

    rows = workflow.list_manual_transit_orders()
    assert len(rows) == 1
    assert rows[0].order_id == "manual-1"
    assert rows[0].received is False

    received = workflow.set_manual_transit_order_received(rows[0], True)
    assert received.received is True
    assert received.received_at
    rows = workflow.list_manual_transit_orders()
    assert rows[0].received is True

    reverted = workflow.set_manual_transit_order_received(rows[0], False)
    assert reverted.received is False
    assert reverted.received_at == ""

    workflow.delete_manual_transit_order("manual-1")
    assert workflow.list_manual_transit_orders() == ()


def test_delivery_ui_has_consistent_primary_and_secondary_actions():
    text = Path(workflow.__file__).read_text(encoding="utf-8")
    assert 'type="secondary"' in text
    assert 'type="primary" if mode_exists else "secondary"' in text
    assert '"Заказ отправлен"' in text
    assert '"Заказ в работе"' in text
    assert '"Shipping"' in text
    assert '"Получен"' in text
    assert '"Сохранить статус"' in text
    assert '"Изменить даты этапов"' in text


def test_pdf_guide_is_packaged_and_used_by_the_site():
    root = Path(__file__).resolve().parent
    pdf = root / "Analitika_USER_GUIDE.pdf"
    assert pdf.exists()
    assert pdf.read_bytes().startswith(b"%PDF")
    app = (root / "streamlit_app.py").read_text(encoding="utf-8")
    assert 'file_name="Analitika_USER_GUIDE.pdf"' in app
    assert 'mime="application/pdf"' in app
    assert 'Скачать красиво оформленное руководство в PDF' in app

from src.order_persistence import S3OrderStorage, S3StorageConfig
from src.order_workflow import OrderDraft


class _MemoryDeliveryStorage(S3OrderStorage):
    def __init__(self) -> None:
        self.config = S3StorageConfig(
            endpoint_url="https://storage.example.test",
            access_key_id="key",
            secret_access_key="secret",
            bucket="orders",
            prefix="analitika/supplier-orders",
        )
        self.objects: dict[str, bytes] = {}

    def exists(self, key: str) -> bool:
        return key in self.objects

    def put_bytes(self, key: str, payload: bytes, content_type: str) -> None:
        self.objects[key] = bytes(payload)

    def get_bytes(self, key: str) -> bytes:
        return self.objects[key]


def test_cloud_receipt_status_updates_mode_manifest_and_index():
    storage = _MemoryDeliveryStorage()
    digest = "7" * 64
    storage.save_workbook(digest, "completed.xlsx", b"xlsx")
    draft = OrderDraft(
        source_hash=digest,
        source_name="completed.xlsx",
        mode=workflow.ORDER_MODE_PEARLS,
        orders={"pearl": 5},
        status="completed",
    )
    storage.save_draft(draft.as_payload())

    saved = storage.set_mode_received(digest, workflow.ORDER_MODE_PEARLS, True)
    assert saved["received"] is True
    assert saved["received_at"]
    row = storage.list_order_index()[0]
    assert row["drafts"][workflow.ORDER_MODE_PEARLS]["received"] is True

    reverted = storage.set_mode_received(digest, workflow.ORDER_MODE_PEARLS, False)
    assert reverted["received"] is False
    assert reverted["received_at"] == ""


def test_cloud_manual_order_is_stored_as_independent_json():
    storage = _MemoryDeliveryStorage()
    payload = storage.save_manual_order(
        {
            "order_id": "casts-25",
            "title": "Касты",
            "order_date": "2026-07-25",
            "note": "В пути",
            "received": False,
        }
    )
    assert payload["title"] == "Касты"
    assert payload["received"] is False
    key = storage.manual_order_key("casts-25")
    assert key in storage.objects


def test_about_page_opens_on_capabilities_and_keeps_history_separate():
    root = Path(__file__).resolve().parent
    app = (root / "streamlit_app.py").read_text(encoding="utf-8")
    about_start = app.index("def render_about()")
    guide_start = app.index("def render_user_guide()")
    about = app[about_start:guide_start]
    assert 'st.markdown("## Возможности")' in about
    assert "Текущая версия" not in about
    assert "Рабочих модулей" not in about
    assert "product-flow" not in about
    assert 'options = ("О программе", "Руководство", "История обновлений")' in app


def test_dated_delivery_statuses_preserve_chronology(monkeypatch, tmp_path):
    _use_temp_db(monkeypatch, tmp_path)
    dates = {
        "sent_at": "2026-06-19",
        "approved_at": "2026-06-22",
        "in_progress_at": "2026-06-30",
        "received_at": "",
    }
    saved = workflow.set_order_delivery_status(
        "dated-report",
        workflow.ORDER_MODE_PEARLS,
        workflow.DELIVERY_STATUS_IN_PROGRESS,
        status_date="2026-06-30",
        delivery_dates=dates,
    )
    assert saved["delivery_status"] == workflow.DELIVERY_STATUS_IN_PROGRESS
    assert saved["delivery_dates"]["sent_at"] == "2026-06-19"
    assert saved["delivery_dates"]["approved_at"] == "2026-06-22"
    assert saved["delivery_dates"]["in_progress_at"] == "2026-06-30"
    assert "Отправлен: 19.06.2026" in workflow.delivery_history_text(
        saved["delivery_dates"], saved["delivery_status"]
    )


def test_manual_quantity_and_stage_dates_are_persisted(monkeypatch, tmp_path):
    _use_temp_db(monkeypatch, tmp_path)
    order = workflow.save_manual_transit_order(
        ManualTransitOrder(
            order_id="manual-dated",
            title="Касты",
            order_date="2026-06-19",
            quantity=420,
            delivery_status=workflow.DELIVERY_STATUS_APPROVED,
            delivery_dates={
                "sent_at": "2026-06-19",
                "approved_at": "2026-06-22",
                "in_progress_at": "",
                "received_at": "",
            },
        )
    )
    assert order.quantity == 420
    assert order.delivery_status == workflow.DELIVERY_STATUS_APPROVED
    assert order.delivery_dates["approved_at"] == "2026-06-22"


def test_manual_order_can_be_edited_without_recreating(monkeypatch, tmp_path):
    _use_temp_db(monkeypatch, tmp_path)
    original = workflow.save_manual_transit_order(
        ManualTransitOrder(
            order_id="editable-manual",
            title="Жемчуг",
            order_date="2026-06-19",
            quantity=120,
            note="Первый комментарий",
            delivery_status=workflow.DELIVERY_STATUS_IN_PROGRESS,
            delivery_dates={
                "sent_at": "2026-06-19",
                "approved_at": "2026-06-22",
                "in_progress_at": "2026-06-30",
                "received_at": "",
            },
        )
    )

    edited = workflow.update_manual_transit_order(
        original,
        title="Жемчуг — июнь",
        order_date="2026-06-20",
        quantity=135,
        note="Исправленный комментарий",
    )

    assert edited.order_id == original.order_id
    assert edited.title == "Жемчуг — июнь"
    assert edited.quantity == 135
    assert edited.note == "Исправленный комментарий"
    assert edited.delivery_status == workflow.DELIVERY_STATUS_IN_PROGRESS
    assert edited.delivery_dates["sent_at"] == "2026-06-20"
    assert edited.delivery_dates["approved_at"] == "2026-06-22"
    assert edited.delivery_dates["in_progress_at"] == "2026-06-30"
    rows = workflow.list_manual_transit_orders()
    assert len(rows) == 1
    assert rows[0].title == "Жемчуг — июнь"


def test_manual_order_form_reset_clears_only_add_form_keys():
    workflow.st.session_state.clear()
    workflow.st.session_state.update(
        {
            "manual_transit_title": "Касты",
            "manual_transit_quantity": 250,
            "manual_transit_note": "Ожидаем",
            "unrelated_state": "keep",
        }
    )
    workflow._reset_manual_transit_form_state()
    assert "manual_transit_title" not in workflow.st.session_state
    assert "manual_transit_quantity" not in workflow.st.session_state
    assert "manual_transit_note" not in workflow.st.session_state
    assert workflow.st.session_state["unrelated_state"] == "keep"


def test_manual_order_form_reset_is_deferred_until_next_rerun():
    text = Path(workflow.__file__).read_text(encoding="utf-8")
    assert 'st.session_state["manual_transit_form_reset_pending"] = True' in text
    assert 'st.session_state.pop("manual_transit_form_reset_pending", False)' in text
    assert 'def _reset_manual_transit_form_state()' in text
    assert 'with st.expander("Редактировать заказ"' in text
    assert '"Сохранить изменения"' in text


def test_password_gate_is_wired_before_the_main_workspace():
    root = Path(__file__).resolve().parent
    app = (root / "streamlit_app.py").read_text(encoding="utf-8")
    auth = (root / "src" / "auth.py").read_text(encoding="utf-8")
    main = app[app.index("def main() -> None:"):]
    assert "if not require_password():" in main
    assert "hmac.compare_digest" in auth
    assert "ANALITIKA_APP_PASSWORD" in auth
    assert "2242" not in auth
