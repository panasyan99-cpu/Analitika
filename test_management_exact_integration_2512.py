from pathlib import Path

ROOT = Path(__file__).resolve().parent


def test_management_mode_is_present_between_supplier_order_and_about():
    source = (ROOT / "src" / "product_info.py").read_text(encoding="utf-8")
    expected = """REPORT_MODES: tuple[str, ...] = (
    "Обычный отчет",
    "Сравнение периодов",
    "Сувениры и касты на складе",
    "Заказ Sonu",
    "Заказ поставщику",
    "Управленческий отчет",
    "О программе",
)"""
    assert expected in source


def test_management_and_exact_full_report_paths_coexist():
    source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    assert "from src.full_sales_report import" in source
    assert "parse_full_sales_report_with_period" in source
    assert 'elif mode == "Управленческий отчет":' in source
    assert "render_management_report_mode()" in source
    assert "render_seller_workspace(detail_df)" in source
    assert "render_supplier_fragment(supplier_df)" in source


def test_full_report_does_not_treat_empty_supplier_filter_as_supplier_hierarchy():
    source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    assert "empty filter label ``Поставщик(и):``" in source
    assert 'hierarchy = " ".join(str(ws.cell(4, column).value or "")' in source


def test_store_63_is_split_in_analytics_but_supplier_order_remains_unchanged():
    normalization = (ROOT / "src" / "store_normalization.py").read_text(encoding="utf-8")
    assert 'return "63 Retail"' in normalization
    assert 'return "63 Timing"' in normalization
    # Supplier-order modules are intentionally outside the 2.5.12 merge scope.
    build = (ROOT / "BUILD_INFO.txt").read_text(encoding="utf-8")
    assert "Supplier order/Sonu: unchanged" in build
