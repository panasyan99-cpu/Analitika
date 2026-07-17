from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from streamlit_app import (
    METAL_GROUPS,
    classify_metal_group,
    filter_metal_groups,
    parse_supplier_report_with_period,
)


def test_metal_group_business_rules():
    assert classify_metal_group("B 925") == "Серебро"
    assert classify_metal_group("Ag 925") == "Серебро"
    assert classify_metal_group("AU 585") == "Золото и платина"
    assert classify_metal_group("Pt 900") == "Золото и платина"
    assert classify_metal_group("OTHER 0") == "Другое"
    assert classify_metal_group("") == "Другое"
    assert METAL_GROUPS == ("Серебро", "Золото и платина", "Другое")


def test_current_report_parser_reads_purity_and_skips_chain(tmp_path: Path):
    path = tmp_path / "sales_with_purity.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Отчет о продажах товаров за период 01.06.2026 - 30.06.2026"
    sheet["A4"] = "Магазин; Камень/вставка; Проба; Номенклатурная группа; Поставщик"

    sheet["A7"] = "AB-Retail"
    sheet["A7"].font = Font(bold=True)

    sheet["A8"] = "BLUE SAPPHIRE"
    sheet["A8"].alignment = Alignment(indent=2)
    sheet["A9"] = "B 925"
    sheet["A9"].alignment = Alignment(indent=4)
    sheet["A10"] = "Earrings"
    sheet["A10"].alignment = Alignment(indent=6)
    sheet["H10"] = 2
    sheet["I10"] = 10_000_000

    sheet["A11"] = "Chain"
    sheet["A11"].alignment = Alignment(indent=6)
    sheet["H11"] = 5
    sheet["I11"] = 5_000_000

    # Blank purity must be preserved as Other / «Не указано».
    sheet["A12"] = "RUBY"
    sheet["A12"].alignment = Alignment(indent=2)
    sheet["A13"].alignment = Alignment(indent=4)
    sheet["A14"] = "Ring"
    sheet["A14"].alignment = Alignment(indent=6)
    sheet["H14"] = 1
    sheet["I14"] = 7_000_000

    workbook.save(path)

    detail, period = parse_supplier_report_with_period(path)
    assert period is not None
    assert len(detail) == 2
    assert int(detail["Количество"].sum()) == 3
    assert set(detail["Проба"]) == {"B 925", "Не указано"}
    assert set(detail["Группа металла"]) == {"Серебро", "Другое"}
    assert "Chain" not in set(detail["Код группы"])

    silver = filter_metal_groups(detail, ["Серебро"])
    assert int(silver["Количество"].sum()) == 2
    assert set(silver["Группа металла"]) == {"Серебро"}


def test_comparison_ui_has_prominent_global_metal_controls_and_purity_section():
    source = Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")
    assert 'id="comparison-filter"' in source
    assert '"Фильтры по пробам", "#comparison-filter"' in source
    assert 'key="comparison_metal_groups"' in source
    assert 'selection_mode="multi"' in source
    assert 'list(METAL_GROUPS)' in source
    assert 'id="comparison-metals"' in source
    assert '"Металлы и пробы", "#comparison-metals"' in source
    assert "render_comparison_metal_section" in source
    assert "rebuild_filtered_stores" in source
