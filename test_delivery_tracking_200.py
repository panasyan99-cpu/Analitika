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
    assert '"Получено"' in text
    assert '● В пути' in text
    assert '✓ Получено' in text


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
