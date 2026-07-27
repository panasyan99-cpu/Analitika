from __future__ import annotations

import json

from src.order_persistence import S3OrderStorage, S3StorageConfig
from src.order_workflow import ORDER_MODE_PEARLS, ORDER_MODE_STONES, OrderDraft


class MemoryConcurrentStorage(S3OrderStorage):
    def __init__(self) -> None:
        self.config = S3StorageConfig(
            endpoint_url="https://storage.example.test",
            access_key_id="key",
            secret_access_key="secret",
            bucket="orders",
            prefix="analitika/supplier-orders",
        )
        self.objects: dict[str, bytes] = {}
        self.entry_scans = 0

    def exists(self, key: str) -> bool:
        return key in self.objects

    def put_bytes(self, key: str, payload: bytes, content_type: str) -> None:
        self.objects[key] = bytes(payload)

    def get_bytes(self, key: str) -> bytes:
        return self.objects[key]

    def _list_index_entries(self):
        self.entry_scans += 1
        return super()._list_index_entries()


def test_mode_manifests_preserve_stones_and_pearls_for_same_workbook() -> None:
    storage = MemoryConcurrentStorage()
    digest = "1" * 64
    storage.save_workbook(digest, "report.xlsx", b"xlsx")

    storage.save_draft(
        OrderDraft(
            source_hash=digest,
            source_name="report.xlsx",
            mode=ORDER_MODE_STONES,
            orders={"stone": 4},
        ).as_payload()
    )
    storage.save_draft(
        OrderDraft(
            source_hash=digest,
            source_name="report.xlsx",
            mode=ORDER_MODE_PEARLS,
            orders={"pearl": 7},
        ).as_payload()
    )

    manifest = json.loads(storage.objects[storage.manifest_key(digest)].decode("utf-8"))
    assert manifest["drafts"][ORDER_MODE_STONES]["total_quantity"] == 4
    assert manifest["drafts"][ORDER_MODE_PEARLS]["total_quantity"] == 7
    assert storage.mode_manifest_key(digest, ORDER_MODE_STONES) in storage.objects
    assert storage.mode_manifest_key(digest, ORDER_MODE_PEARLS) in storage.objects


def test_normal_library_open_reads_compact_index_without_per_order_scan() -> None:
    storage = MemoryConcurrentStorage()
    digest = "2" * 64
    storage.save_workbook(digest, "report.xlsx", b"xlsx")
    storage.save_draft(
        OrderDraft(
            source_hash=digest,
            source_name="report.xlsx",
            mode=ORDER_MODE_STONES,
            orders={"stone": 3},
        ).as_payload()
    )

    storage.entry_scans = 0
    rows = storage.list_order_index()

    assert len(rows) == 1
    assert storage.entry_scans == 0
