from __future__ import annotations

import json
from pathlib import Path

from src.management_report_analytics import (
    build_management_snapshot,
    canonical_store,
    is_technical_manager,
)
from src.management_report_parser import (
    Metrics,
    ParsedReport,
    ProductFact,
    ReportMeta,
    _period_from_title,
)
from src.management_report_suppliers import SupplierCatalog, UNKNOWN_SUPPLIER, normalize_sku


ROOT = Path(__file__).resolve().parent


def _report(label: str, start: str, end: str, facts: list[ProductFact], totals: Metrics) -> ParsedReport:
    return ParsedReport(
        meta=ReportMeta(
            source_file=f"{label}.xlsx",
            title=label,
            period_label=label,
            period_start=start,
            period_end=end,
            period_days=30,
            generated_at=None,
            generated_by="",
            grouping_label="Магазин; Менеджер; Товар; Камень/вставка; Проба; Номенклатурная группа",
        ),
        totals=totals,
        facts=facts,
        stores={},
        validation={},
    )


def _fact(*, sku: str, revenue: float, quantity: float, manager: str = "Seller", store: str = "TT") -> ProductFact:
    return ProductFact(
        row_number=1,
        store=store,
        manager=manager,
        top_group="Jewelry",
        product_section="Silver",
        category="Earrings",
        sku=sku,
        stone="BLUE SAPPHIRE",
        assay="B 925",
        note="",
        metrics=Metrics(quantity=quantity, average_price=revenue / quantity, revenue=revenue),
    )


def test_management_mode_is_between_supplier_order_and_about() -> None:
    product = (ROOT / "src" / "product_info.py").read_text(encoding="utf-8")
    modes = product[product.index("REPORT_MODES"):]
    assert modes.index('"Заказ поставщику"') < modes.index('"Управленческий отчет"') < modes.index('"О программе"')

    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    assert 'elif mode == "Управленческий отчет"' in app
    assert "render_management_report_dashboard()" in app
    assert '"Управленческий отчет": "7. Управленческий отчет"' in app


def test_help_assets_and_two_workspace_tabs_are_present() -> None:
    assert (ROOT / "assets" / "management_report_setup_1.png").exists()
    assert (ROOT / "assets" / "management_report_setup_2.png").exists()
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    assert 'options = ("Работа", "Как с этим работать")' in app
    assert 'mode == "Управленческий отчет"' in app


def test_supported_period_titles_are_detected() -> None:
    start, end, label = _period_from_title("Отчет за период 01.07.2026 - 30.07.2026")
    assert (start.isoformat(), end.isoformat(), label) == (
        "2026-07-01",
        "2026-07-30",
        "01.07.2026–30.07.2026",
    )
    start, end, label = _period_from_title("Отчет о продажах товаров за период Июнь 2026 г.")
    assert start.isoformat() == "2026-06-01"
    assert end.isoformat() == "2026-06-30"
    assert label == "Июнь 2026"


def test_store_and_manager_normalization() -> None:
    assert canonical_store("63NDC-Retail") == "63"
    assert canonical_store("Gifts-ТТ") == "Gifts-TT"
    assert canonical_store("Cafe TT") == "Cafe"
    assert is_technical_manager("AB-cashier")
    assert is_technical_manager("Vietnamese Staff")
    assert not is_technical_manager("Aimani Bisultanova")


def test_supplier_catalog_uses_exact_then_override_and_keeps_unknown_visible() -> None:
    catalog = SupplierCatalog(
        exact={"SKU1": "Taiwan"},
        family_rules={},
        overrides={"SKU1": "Sonu"},
        suppliers=("Sonu", "Taiwan"),
    )
    assert normalize_sku(" sku 1 ") == "SKU1"
    resolution = catalog.resolve("SKU1")
    assert resolution.supplier == "Sonu"
    assert resolution.source == "manual"
    assert catalog.resolve("UNKNOWN-123").supplier == UNKNOWN_SUPPLIER


def test_snapshot_uses_official_totals_daily_metrics_and_excludes_technical_sellers() -> None:
    june_facts = [
        _fact(sku="SKU1", revenue=900.0, quantity=9.0, manager="Seller", store="TT"),
        _fact(sku="SKU2", revenue=100.0, quantity=1.0, manager="Admin", store="Gifts-ТТ"),
    ]
    july_facts = [
        _fact(sku="SKU1", revenue=1_200.0, quantity=12.0, manager="Seller", store="TT"),
        _fact(sku="SKU2", revenue=300.0, quantity=3.0, manager="Admin", store="Cafe"),
    ]
    june = _report("June", "2026-06-01", "2026-06-30", june_facts, Metrics(quantity=10, average_price=100, revenue=1_000))
    july = _report("July", "2026-07-01", "2026-07-30", july_facts, Metrics(quantity=15, average_price=100, revenue=1_500))
    catalog = SupplierCatalog(
        exact={"SKU1": "Taiwan", "SKU2": "Other"},
        family_rules={},
        overrides={},
        suppliers=("Other", "Taiwan"),
    )

    snapshot = build_management_snapshot(june, july, catalog)
    assert snapshot["overall"]["revenue_pct"] == 50.0
    assert snapshot["overall"]["daily"]["new_revenue"] == 50.0
    assert list(snapshot["dimensions"]["managers"]["Позиция"]) == ["Seller"]
    assert snapshot["outlet"].iloc[0]["Выручка · Период 2"] == 1_500.0
    assert snapshot["supplier_quality"]["new_revenue_coverage_pct"] == 100.0


def test_trained_supplier_mapping_is_versioned_and_nonempty() -> None:
    payload = json.loads((ROOT / "data" / "management_supplier_mapping.json").read_text(encoding="utf-8"))
    assert payload["schema_version"] == 1
    assert len(payload["exact"]) >= 3_000
    assert len(payload["family_rules"]) >= 1_000
    assert "2026-06-01/2026-06-30" in payload["trained_periods"]


def test_management_copy_contains_no_operational_advice() -> None:
    source = (ROOT / "src" / "management_report.py").read_text(encoding="utf-8").casefold()
    for phrase in (
        "рекомендуем сделать",
        "следует увеличить",
        "необходимо усилить",
        "советуем",
        "приоритетными задачами",
    ):
        assert phrase not in source


def test_release_is_on_stable_2510_line_not_260() -> None:
    version = json.loads((ROOT / "version.json").read_text(encoding="utf-8"))
    assert version["version"] == "2.5.10"
    assert version["channel"] == "stable"
    build = (ROOT / "BUILD_INFO.txt").read_text(encoding="utf-8")
    assert "Base: Analitika Web 2.5.9" in build
    assert "Version 2.6.0: not used" in build
