from __future__ import annotations

import json
from pathlib import Path

import pytest

import src.order_workflow as workflow
from src.order_persistence import CloudStorageError, S3OrderStorage, S3StorageConfig
from src.order_workflow import ORDER_MODE_PEARLS, ORDER_MODE_STONES, OrderDraft


class MemoryIndexedStorage(S3OrderStorage):
    def __init__(self) -> None:
        self.config = S3StorageConfig(
            endpoint_url="https://storage.example.test",
            access_key_id="key",
            secret_access_key="secret",
            bucket="orders",
            prefix="analitika/supplier-orders",
        )
        self.objects: dict[str, bytes] = {}
        self.protected: set[str] = set()

    def exists(self, key: str) -> bool:
        return key in self.objects

    def put_bytes(self, key: str, payload: bytes, content_type: str) -> None:
        self.objects[key] = bytes(payload)

    def get_bytes(self, key: str) -> bytes:
        return self.objects[key]

    def download_file(self, key: str, destination: Path) -> Path:
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(self.objects[key])
        return destination

    def list_manifests(self):
        values = []
        for key, payload in self.objects.items():
            if key.endswith("/manifest.json"):
                values.append(json.loads(payload.decode("utf-8")))
        return tuple(sorted(values, key=lambda row: row.get("updated_at", ""), reverse=True))

    def list_workspace_keys(self, source_hash: str):
        prefix = self._workspace_prefix(source_hash).rstrip("/") + "/"
        return tuple(sorted(key for key in self.objects if key.startswith(prefix)))

    def delete_keys(self, keys):
        failures = []
        for key in keys:
            if key in self.protected:
                failures.append(key)
            else:
                self.objects.pop(key, None)
        return tuple(failures)


def test_cloud_index_contains_resume_metadata_without_loading_workbook() -> None:
    storage = MemoryIndexedStorage()
    digest = "c" * 64
    storage.save_workbook(digest, "July order.xlsx", b"xlsx")
    stones = OrderDraft(
        source_hash=digest,
        source_name="July order.xlsx",
        mode=ORDER_MODE_STONES,
        orders={"sku-1": 4, "sku-2": 3},
        limited_orders={"sku-limited": True},
        stage="rings",
    )
    storage.save_draft(stones.as_payload())
    pearls = OrderDraft(
        source_hash=digest,
        source_name="July order.xlsx",
        mode=ORDER_MODE_PEARLS,
        orders={"pearl-1": 8},
        status="completed",
    )
    storage.save_draft(pearls.as_payload())

    rows = storage.list_order_index()
    assert len(rows) == 1
    row = rows[0]
    assert row["source_name"] == "July order.xlsx"
    assert row["created_at"]
    assert row["drafts"][ORDER_MODE_STONES]["selected_positions"] == 2
    assert row["drafts"][ORDER_MODE_STONES]["total_quantity"] == 7
    assert row["drafts"][ORDER_MODE_STONES]["limited_positions"] == 1
    assert row["drafts"][ORDER_MODE_STONES]["stage"] == "rings"
    assert row["drafts"][ORDER_MODE_PEARLS]["status"] == "completed"
    assert storage.index_key() in storage.objects


def test_verified_cloud_delete_removes_workspace_and_index_entry() -> None:
    storage = MemoryIndexedStorage()
    digest = "d" * 64
    storage.save_workbook(digest, "Delete me.xlsx", b"xlsx")
    draft = OrderDraft(
        source_hash=digest,
        source_name="Delete me.xlsx",
        mode=ORDER_MODE_STONES,
        orders={"sku": 2},
    )
    storage.save_draft(draft.as_payload())
    backup_key = f"{storage._workspace_prefix(digest)}/history/backup-1.json"
    storage.put_bytes(backup_key, b"{}", "application/json")

    removed = storage.delete_workspace(digest)
    assert backup_key in removed
    assert storage.list_workspace_keys(digest) == ()
    assert storage.list_order_index() == ()


def test_partial_cloud_delete_keeps_order_in_index_and_reports_remaining_key() -> None:
    storage = MemoryIndexedStorage()
    digest = "e" * 64
    storage.save_workbook(digest, "Protected.xlsx", b"xlsx")
    draft = OrderDraft(
        source_hash=digest,
        source_name="Protected.xlsx",
        mode=ORDER_MODE_STONES,
        orders={"sku": 1},
    )
    storage.save_draft(draft.as_payload())
    protected = storage.draft_key(digest, ORDER_MODE_STONES)
    storage.protected.add(protected)

    with pytest.raises(CloudStorageError, match="Остались объекты") as error:
        storage.delete_workspace(digest)
    assert protected in str(error.value)
    assert len(storage.list_order_index()) == 1


def test_local_delete_removes_all_drafts_and_cached_workbooks(tmp_path: Path, monkeypatch) -> None:
    runtime = tmp_path / "runtime"
    monkeypatch.setattr(workflow, "UPLOAD_DIR", runtime / "uploads")
    monkeypatch.setattr(workflow, "DRAFT_DB", runtime / "drafts.sqlite3")
    workflow.UPLOAD_DIR.mkdir(parents=True)
    digest = "f" * 64
    (workflow.UPLOAD_DIR / f"{digest}.xlsx").write_bytes(b"xlsx")
    (workflow.UPLOAD_DIR / f"{digest}.xlsm").write_bytes(b"xlsm")
    for mode in (ORDER_MODE_STONES, ORDER_MODE_PEARLS):
        draft = OrderDraft(source_hash=digest, source_name="Local.xlsx", mode=mode, orders={mode: 3})
        workflow._save_draft_locally(draft.as_payload())

    rows, files = workflow._delete_local_order_workspace(digest)
    assert rows == 2
    assert files == 2
    assert workflow._find_uploaded_workbook(digest) is None
    with workflow._connect_drafts() as connection:
        count = connection.execute("SELECT COUNT(*) FROM order_drafts WHERE source_hash = ?", (digest,)).fetchone()[0]
    assert count == 0


def test_cloud_library_ui_has_direct_resume_refresh_delete_and_no_json_uploader() -> None:
    source = Path(workflow.__file__).read_text(encoding="utf-8")
    assert "Незавершённые заказы" in source
    assert "Обновить список" in source
    assert "Продолжить заказ" in source
    assert "Да, удалить отовсюду" in source
    assert "Показать завершённые" in source
    assert "Восстановить JSON" not in source
    assert 'st.file_uploader("Восстановить JSON"' not in source
    assert "list_order_index" in Path(workflow.__file__).with_name("order_persistence.py").read_text(encoding="utf-8")
