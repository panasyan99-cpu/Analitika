from __future__ import annotations

from pathlib import Path

from src.order_workflow import is_store_63
from src.report import detect_store, normalize_store_from_report
from src.store_normalization import analytics_store_name, supplier_order_store_name


def test_63_is_split_for_analytics_labels() -> None:
    assert analytics_store_name("63NDC-Retail") == "63 Retail"
    assert analytics_store_name("63NDC-Timings") == "63 Timing"
    assert normalize_store_from_report("63NDC-Retail") == "63 Retail"
    assert normalize_store_from_report("63NDC-Timings") == "63 Timing"
    assert detect_store(Path("63NDC-Retail.xlsx")) == "63 Retail"
    assert detect_store(Path("63NDC-Timings.xlsx")) == "63 Timing"


def test_63_remains_combined_for_supplier_order() -> None:
    for label in ("63NDC-Retail", "63NDC-Timings", "63.1", "63.2"):
        assert supplier_order_store_name(label) == "63"
        assert is_store_63(label)
    assert supplier_order_store_name("63 Retail") == "63"
    assert supplier_order_store_name("63 Timing") == "63"


def test_legacy_ambiguous_63_is_not_guessed_in_analytics() -> None:
    assert analytics_store_name("63") == "63"
    assert analytics_store_name("63.1") == "63"
    assert analytics_store_name("63.2") == "63"
