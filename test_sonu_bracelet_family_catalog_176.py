import json
from pathlib import Path

from src.sonu import (
    BRACELET_CATALOG_FILE,
    CENTERED_BRACELET_LABEL,
    FULL_CIRCLE_BRACELET_LABEL,
    load_bracelet_catalog,
)


def test_sonunew_catalog_contains_all_reviewed_skus_and_small_pending_queue():
    payload = json.loads(Path(BRACELET_CATALOG_FILE).read_text(encoding="utf-8"))
    metadata = payload["metadata"]
    assert metadata["bracelet_sku_total"] == 350
    assert metadata["classified_sku"] == 340
    assert metadata["pending_sku"] == 10
    assert metadata["pending_families"] == 9
    assert len(payload["sku_overrides"]) == 340
    assert set(payload["sku_overrides"].values()) <= {
        CENTERED_BRACELET_LABEL,
        FULL_CIRCLE_BRACELET_LABEL,
    }


def test_catalog_loader_exposes_family_and_sku_rules():
    catalog = load_bracelet_catalog()
    assert len(catalog["sku_overrides"]) == 340
    assert len(catalog["family_overrides"]) == 209
    assert len(catalog["pending_families"]) == 9
