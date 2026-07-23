from __future__ import annotations

import json
from pathlib import Path
from zipfile import ZipFile

import src.order_workflow as workflow
from src.order_persistence import S3OrderStorage, S3StorageConfig
from src.order_workflow import ORDER_MODE_STONES, OrderDraft, ParsedOrderWorkbook, full_order_backup_bytes


class MemoryOrderStorage(S3OrderStorage):
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


def test_cloud_storage_persists_workbook_and_draft_and_restores_to_new_runtime(tmp_path: Path) -> None:
    storage = MemoryOrderStorage()
    report = b"PK\x03\x04fake-large-xlsx"
    digest = "a" * 64
    storage.save_workbook(digest, "July supplier order.xlsx", report)
    draft = OrderDraft(
        source_hash=digest,
        source_name="July supplier order.xlsx",
        mode=ORDER_MODE_STONES,
        orders={"set|sku|1": 5, "zero": 0},
        sizes={"set|sku|1": {"18": 2, "19": 3, "20": 0}},
        stock_checked={"set|sku|1": True, "zero": False},
        stage="rings",
    )
    payload = draft.as_payload()
    storage.save_draft(payload)

    restored_path, manifest = storage.restore_workbook(digest, tmp_path / "fresh-runtime")
    assert restored_path.read_bytes() == report
    assert manifest["source_name"] == "July supplier order.xlsx"
    assert manifest["drafts"][ORDER_MODE_STONES]["selected_positions"] == 1
    assert manifest["drafts"][ORDER_MODE_STONES]["total_quantity"] == 5
    assert storage.load_draft(digest, ORDER_MODE_STONES)["stage"] == "rings"


def test_draft_payload_is_sparse_for_fast_cloud_autosave() -> None:
    draft = OrderDraft(
        source_hash="hash",
        source_name="report.xlsx",
        mode=ORDER_MODE_STONES,
        orders={"selected": 7, "not-selected": 0},
        sizes={"selected": {"18": 4, "19": 3, "20": 0}, "empty": {"18": 0}},
        stock_checked={"selected": True, "unchecked": False},
    )
    payload = draft.as_payload()
    assert payload["orders"] == {"selected": 7}
    assert payload["sizes"] == {"selected": {"18": 4, "19": 3}}
    assert payload["stock_checked"] == {"selected": True}


def test_cloud_manifest_becomes_saved_workspace() -> None:
    manifest = {
        "source_hash": "b" * 64,
        "source_name": "not-called-order.xlsx",
        "workbook_key": "analitika/supplier-orders/workspaces/hash/source.xlsx",
        "updated_at": "2026-07-23T10:00:00+00:00",
        "drafts": {
            ORDER_MODE_STONES: {
                "updated_at": "2026-07-23T10:01:00+00:00",
                "selected_positions": 12,
                "total_quantity": 57,
            }
        },
    }
    workspace = workflow._cloud_workspace_from_manifest(manifest)
    assert workspace is not None
    assert workspace.storage == "cloud"
    assert workspace.source_name == "not-called-order.xlsx"
    assert workspace.selected_positions == 12
    assert workspace.total_quantity == 57


def test_full_emergency_backup_contains_source_workbook_and_draft(tmp_path: Path) -> None:
    report = tmp_path / "custom report.xlsx"
    report.write_bytes(b"xlsx-bytes")
    parsed = ParsedOrderWorkbook(
        source_name=report.name,
        source_hash="hash",
        upload_path=str(report),
        period="",
        supplier="Y&J",
        store_columns=(),
        has_actual_ntr2=True,
        items=(),
    )
    draft = OrderDraft(
        source_hash="hash",
        source_name=report.name,
        mode=ORDER_MODE_STONES,
        orders={"item": 3},
    )
    backup = tmp_path / "backup.zip"
    backup.write_bytes(full_order_backup_bytes(parsed, draft))
    with ZipFile(backup) as archive:
        assert archive.read(report.name) == b"xlsx-bytes"
        saved_draft = json.loads(archive.read("order_draft.json"))
        assert saved_draft["orders"] == {"item": 3}
        assert json.loads(archive.read("backup_manifest.json"))["source_hash"] == "hash"


def test_new_report_no_longer_deletes_previous_saved_workspace() -> None:
    assert workflow.purge_order_workspaces_except("new-hash") == (0, 0)
