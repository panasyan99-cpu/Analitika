from __future__ import annotations

from pathlib import Path

import src.order_workflow as workflow


class _RepairStorage:
    def __init__(self, drafts):
        self.drafts = {key: dict(value) for key, value in drafts.items()}
        self.status_calls = []

    def load_draft(self, source_hash: str, mode: str):
        value = self.drafts.get((source_hash, mode))
        return dict(value) if value else None

    def save_draft(self, payload):
        key = (str(payload["source_hash"]), str(payload["mode"]))
        self.drafts[key] = dict(payload)
        return {"ok": True}

    def set_mode_delivery_status(self, source_hash, mode, status, *, status_date="", delivery_dates=None):
        self.status_calls.append((source_hash, mode, status, status_date, dict(delivery_dates or {})))
        return {"status": "completed"}


def test_known_transmitted_order_is_repaired_as_completed_with_sent_dates(monkeypatch, tmp_path: Path):
    monkeypatch.setattr(workflow, "DRAFT_DB", tmp_path / "orders.sqlite3")
    source_hash = "a" * 64
    drafts = {
        (source_hash, workflow.ORDER_MODE_STONES): workflow.OrderDraft(
            source_hash=source_hash,
            source_name="заказnew.xlsx",
            mode=workflow.ORDER_MODE_STONES,
            orders={"stone": 2591},
            status="draft",
        ).as_payload(),
        (source_hash, workflow.ORDER_MODE_PEARLS): workflow.OrderDraft(
            source_hash=source_hash,
            source_name="заказnew.xlsx",
            mode=workflow.ORDER_MODE_PEARLS,
            orders={"pearl": 1846},
            status="draft",
        ).as_payload(),
    }
    storage = _RepairStorage(drafts)
    monkeypatch.setattr(workflow, "get_cloud_storage", lambda: storage)
    workspace = workflow.SavedOrderWorkspace(
        source_hash=source_hash,
        source_name="заказnew.xlsx",
        upload_path="",
        updated_at="2026-07-27T17:19:19",
        modes=workflow.ORDER_MODES,
        preferred_mode=workflow.ORDER_MODE_STONES,
        selected_positions=1168,
        total_quantity=4437,
        storage="cloud",
        created_at="2026-07-23T06:44:20",
        status="draft",
        mode_details={
            workflow.ORDER_MODE_STONES: {
                "status": "draft",
                "total_quantity": 2591,
                "delivery_status": workflow.DELIVERY_STATUS_SENT,
                "delivery_dates": {},
            },
            workflow.ORDER_MODE_PEARLS: {
                "status": "draft",
                "total_quantity": 1846,
                "delivery_status": workflow.DELIVERY_STATUS_SENT,
                "delivery_dates": {},
            },
        },
    )

    assert workflow._repair_historical_completed_order(workspace) is True
    assert storage.drafts[(source_hash, workflow.ORDER_MODE_STONES)]["status"] == "completed"
    assert storage.drafts[(source_hash, workflow.ORDER_MODE_PEARLS)]["status"] == "completed"
    dates = {mode: sent_at for _, mode, _, sent_at, _ in storage.status_calls}
    assert dates[workflow.ORDER_MODE_PEARLS] == "2026-07-24"
    assert dates[workflow.ORDER_MODE_STONES] == "2026-07-25"


def test_completed_library_open_sets_preservation_guard():
    source = Path(workflow.__file__).read_text(encoding="utf-8")
    assert "supplier_order_expected_completed::" in source
    assert "Reconcile old cloud payloads" in source
    assert "_flush_session_draft(draft)" in source
    assert "st.session_state.pop(" in source


def test_repair_is_narrowly_guarded_by_filename_and_quantities():
    repair = workflow.HISTORICAL_COMPLETED_ORDER_REPAIRS["заказnew.xlsx"]
    assert repair[workflow.ORDER_MODE_PEARLS]["sent_at"] == "2026-07-24"
    assert repair[workflow.ORDER_MODE_PEARLS]["total_quantity"] == 1846
    assert repair[workflow.ORDER_MODE_STONES]["sent_at"] == "2026-07-25"
    assert repair[workflow.ORDER_MODE_STONES]["total_quantity"] == 2591
