from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace

import pytest

import src.order_persistence as persistence
import src.order_workflow as workflow
import src.sonu as sonu
from src.order_workflow import DraftPersistenceResult, ORDER_MODE_STONES, OrderDraft


def test_total_draft_save_failure_never_claims_success(monkeypatch: pytest.MonkeyPatch) -> None:
    state: dict[str, object] = {}
    monkeypatch.setattr(workflow, "st", SimpleNamespace(session_state=state))
    monkeypatch.setattr(workflow, "diagnostic_event", lambda *args, **kwargs: None)
    monkeypatch.setattr(
        workflow,
        "persist_draft",
        lambda draft, sync_cloud=True: DraftPersistenceResult(
            saved_at=draft.updated_at,
            local_saved=False,
            cloud_configured=True,
            cloud_saved=False,
            local_error="disk full",
            cloud_error="R2 offline",
        ),
    )
    draft = OrderDraft(source_hash="hash", source_name="report.xlsx", mode=ORDER_MODE_STONES)

    assert workflow._save_session_draft(draft, sync_cloud=True) is False
    assert state[workflow._draft_dirty_key(draft)] is True
    assert state["supplier_order_save_blocked"] is True
    assert str(state["supplier_order_save_status"]).startswith("НЕ СОХРАНЕНО")
    assert "Локально сохранено" not in str(state["supplier_order_save_status"])


def test_local_draft_fallback_is_reported_truthfully(monkeypatch: pytest.MonkeyPatch) -> None:
    state: dict[str, object] = {}
    monkeypatch.setattr(workflow, "st", SimpleNamespace(session_state=state))
    monkeypatch.setattr(workflow, "diagnostic_event", lambda *args, **kwargs: None)
    monkeypatch.setattr(
        workflow,
        "persist_draft",
        lambda draft, sync_cloud=True: DraftPersistenceResult(
            saved_at=draft.updated_at,
            local_saved=True,
            cloud_configured=True,
            cloud_saved=False,
            cloud_error="timeout",
        ),
    )
    draft = OrderDraft(source_hash="hash", source_name="report.xlsx", mode=ORDER_MODE_STONES)

    assert workflow._save_session_draft(draft, sync_cloud=True) is True
    assert state[workflow._draft_dirty_key(draft)] is True
    assert state["supplier_order_save_blocked"] is False
    assert "Сохранено локально" in str(state["supplier_order_save_status"])
    assert "облако временно недоступно" in str(state["supplier_order_save_status"])


def test_cloud_status_retries_and_force_refresh(monkeypatch: pytest.MonkeyPatch) -> None:
    config = persistence.S3StorageConfig(
        endpoint_url="https://example.invalid",
        access_key_id="key",
        secret_access_key="secret",
        bucket="bucket",
    )
    attempts = {"count": 0}

    class FakeStorage:
        def check(self) -> None:
            attempts["count"] += 1
            if attempts["count"] < 3:
                raise persistence.CloudStorageError("temporary")

    monkeypatch.setattr(persistence, "load_storage_config", lambda: config)
    def fake_get_cloud_storage():
        return FakeStorage()
    fake_get_cloud_storage.cache_clear = lambda: None
    monkeypatch.setattr(persistence, "get_cloud_storage", fake_get_cloud_storage)
    monkeypatch.setattr(persistence.time, "sleep", lambda _seconds: None)
    persistence._cloud_storage_status_for_bucket.cache_clear()

    status = persistence.get_cloud_storage_status(force=True)
    assert status.available is True
    assert attempts["count"] == 3

    # Cached result does not perform another health request inside the TTL bucket.
    status_again = persistence.get_cloud_storage_status()
    assert status_again.available is True
    assert attempts["count"] == 3

    persistence.get_cloud_storage_status(force=True)
    assert attempts["count"] == 4


def test_shared_state_key_rejects_parent_segments() -> None:
    storage = object.__new__(persistence.S3OrderStorage)
    storage.config = persistence.S3StorageConfig("", "a", "b", "bucket", prefix="root")
    assert storage.shared_key("sonu/bracelets.json") == "root/system/sonu/bracelets.json"
    assert storage.shared_key("../sonu/./bracelets.json") == "root/system/sonu/bracelets.json"


class FakeSonuStorage:
    def __init__(self, rows: list[dict[str, str]] | None = None) -> None:
        self.rows = list(rows or [])
        self.saved: dict[str, dict[str, object]] = {}

    def list_json_prefix(self, _prefix: str, *, limit: int = 1000):
        return tuple(self.rows[:limit])

    def load_shared_json(self, name: str):
        return self.saved.get(name)

    def save_shared_json(self, name: str, payload):
        self.saved[name] = dict(payload)


def _fake_streamlit_state() -> SimpleNamespace:
    return SimpleNamespace(runtime=SimpleNamespace(exists=lambda: True), session_state={})


def test_sonu_cloud_overrides_take_precedence_over_local_backup(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    local_path = tmp_path / "bracelet_overrides.json"
    local_path.write_text(
        json.dumps({"sku_overrides": {"SKU-1": sonu.CENTERED_BRACELET_LABEL}}, ensure_ascii=False),
        encoding="utf-8",
    )
    storage = FakeSonuStorage(
        [{"key": "SKU-1", "value": sonu.FULL_CIRCLE_BRACELET_LABEL}]
    )
    monkeypatch.setattr(sonu, "BRACELET_OVERRIDE_FILE", local_path)
    monkeypatch.setattr(sonu, "get_cloud_storage", lambda: storage)
    monkeypatch.setattr(sonu, "st", _fake_streamlit_state())
    sonu._clear_bracelet_cloud_cache()

    loaded = sonu.load_bracelet_overrides()
    assert loaded["SKU-1"] == sonu.FULL_CIRCLE_BRACELET_LABEL


def test_sonu_next_successful_save_flushes_pending_local_decisions(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    local_path = tmp_path / "bracelet_overrides.json"
    local_path.write_text(
        json.dumps({"sku_overrides": {"OLD-SKU": sonu.CENTERED_BRACELET_LABEL}}, ensure_ascii=False),
        encoding="utf-8",
    )
    storage = FakeSonuStorage()
    fake_st = _fake_streamlit_state()
    monkeypatch.setattr(sonu, "BRACELET_OVERRIDE_FILE", local_path)
    monkeypatch.setattr(sonu, "get_cloud_storage", lambda: storage)
    monkeypatch.setattr(sonu, "st", fake_st)
    sonu._clear_bracelet_cloud_cache()

    merged, persisted, message = sonu.save_bracelet_overrides(
        {"NEW-SKU": sonu.FULL_CIRCLE_BRACELET_LABEL}
    )

    assert persisted is True
    assert "постоянно в облаке" in message
    assert merged["OLD-SKU"] == sonu.CENTERED_BRACELET_LABEL
    assert merged["NEW-SKU"] == sonu.FULL_CIRCLE_BRACELET_LABEL
    snapshot = storage.saved[sonu.BRACELET_CLOUD_SNAPSHOT]
    assert snapshot["sku_overrides"]["OLD-SKU"] == sonu.CENTERED_BRACELET_LABEL
    assert snapshot["sku_overrides"]["NEW-SKU"] == sonu.FULL_CIRCLE_BRACELET_LABEL
    assert fake_st.session_state[sonu.BRACELET_PENDING_SYNC_KEY] is False


def test_sonu_cloud_failure_keeps_local_backup_pending_sync(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    local_path = tmp_path / "bracelet_overrides.json"
    fake_st = _fake_streamlit_state()
    monkeypatch.setattr(sonu, "BRACELET_OVERRIDE_FILE", local_path)
    monkeypatch.setattr(sonu, "get_cloud_storage", lambda: None)
    monkeypatch.setattr(sonu, "st", fake_st)
    sonu._clear_bracelet_cloud_cache()

    merged, persisted, message = sonu.save_bracelet_overrides(
        {"SKU-2": sonu.CENTERED_BRACELET_LABEL}
    )

    assert persisted is True
    assert merged["SKU-2"] == sonu.CENTERED_BRACELET_LABEL
    assert "ожидают синхронизации" in message
    assert local_path.exists()
    assert fake_st.session_state[sonu.BRACELET_PENDING_SYNC_KEY] is True


def test_sonu_pending_file_wins_over_older_cloud_until_sync(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    local_path = tmp_path / "bracelet_overrides.json"
    pending_path = tmp_path / "bracelet_classification_overrides.pending.json"
    local_path.write_text(
        json.dumps({"sku_overrides": {"SKU-P": sonu.CENTERED_BRACELET_LABEL}}, ensure_ascii=False),
        encoding="utf-8",
    )
    pending_path.write_text(
        json.dumps({"sku_overrides": {"SKU-P": sonu.FULL_CIRCLE_BRACELET_LABEL}}, ensure_ascii=False),
        encoding="utf-8",
    )
    storage = FakeSonuStorage(
        [{"key": "SKU-P", "value": sonu.CENTERED_BRACELET_LABEL}]
    )
    fake_st = _fake_streamlit_state()
    monkeypatch.setattr(sonu, "BRACELET_OVERRIDE_FILE", local_path)
    monkeypatch.setattr(sonu, "get_cloud_storage", lambda: storage)
    monkeypatch.setattr(sonu, "st", fake_st)
    sonu._clear_bracelet_cloud_cache()

    loaded = sonu.load_bracelet_overrides()

    assert loaded["SKU-P"] == sonu.FULL_CIRCLE_BRACELET_LABEL
    assert fake_st.session_state[sonu.BRACELET_PENDING_SYNC_KEY] is True
