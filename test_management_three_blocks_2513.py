from __future__ import annotations

from io import BytesIO
import zipfile
from xml.sax.saxutils import escape

from src.management_block_reports import (
    CONSULTANTS,
    SALES,
    SUPPLIERS,
    UNKNOWN,
    cross_block_validation,
    parse_block_report,
    validate_period_bundle,
)


def _xlsx(grouping: str, data_rows: list[tuple[int, str, float, float]], *, total_quantity: float = 10.5) -> bytes:
    rows: list[str] = []

    def cell(column: str, value: object) -> str:
        if isinstance(value, str):
            return f'<c r="{column}{row_number}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
        return f'<c r="{column}{row_number}"><v>{value}</v></c>'

    source = [
        (1, 0, "Отчет о продажах товаров за период Июль 2026 г.", 0, 0),
        (2, 0, "Поставщик(и):", 0, 0),
        (3, 0, "Товар(ы):", 0, 0),
        (4, 0, grouping, 0, 0),
        (5, 0, "", 0, 0),
        (6, 0, "", 0, 0),
    ]
    source.extend((index + 7, level, label, quantity, revenue) for index, (level, label, quantity, revenue) in enumerate(data_rows))
    total_row = len(source) + 1
    source.append((total_row, 0, "Итого:", total_quantity, 1100))
    source.append((total_row + 1, 0, "01.08.2026 13:00:00 Vladimir Panasian", 0, 0))

    for row_number, level, label, quantity, revenue in source:
        attributes = f' r="{row_number}"' + (f' outlineLevel="{level}"' if level else "")
        cells = [cell("A", label)]
        if row_number >= 7 and row_number <= total_row:
            cells.extend([cell("D", quantity), cell("G", revenue), cell("H", quantity), cell("I", revenue)])
        rows.append(f"<row{attributes}>{''.join(cells)}</row>")

    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData>' + "".join(rows) + '</sheetData></worksheet>'
    )
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("xl/worksheets/sheet1.xml", xml)
    return buffer.getvalue()


def _reports():
    sales = parse_block_report(
        _xlsx(
            "Магазин; Камень/вставка; Проба; Номенклатурная группа",
            [(0, "63NDC-Retail", 5, 500), (0, "63NDC-Timing", 6, 600)],
        ),
        kind=SALES,
    )
    consultants = parse_block_report(
        _xlsx(
            "Менеджер; Проба; Номенклатурная группа",
            [(0, "Admin", 1, 100), (0, "", 2, 200), (0, "Alice", 8, 800)],
        ),
        kind=CONSULTANTS,
    )
    suppliers = parse_block_report(
        _xlsx(
            "Номенклатурная группа; Поставщик",
            [
                (0, "Bracelet", 11, 1100),
                (1, "", 2, 200),
                (1, "Поставщики", 9, 900),
                (2, "Own production service", 4, 400),
                (2, "Taiwan", 5, 500),
            ],
        ),
        kind=SUPPLIERS,
    )
    return {SALES: sales, CONSULTANTS: consultants, SUPPLIERS: suppliers}


def test_sales_block_uses_exact_total_and_splits_63_locations():
    sales = _reports()[SALES]
    assert sales.totals.quantity == 10.5
    assert sales.dimensions["stores"]["63 Retail"].revenue == 500
    assert sales.dimensions["stores"]["63 Timing"].revenue == 600
    assert sales.validation["quantity_difference"] == 0.5
    assert sales.validation["valid"] is True


def test_consultant_block_keeps_technical_and_blank_manager_rows():
    consultants = _reports()[CONSULTANTS]
    assert consultants.dimensions["consultants"]["Admin"].revenue == 100
    assert consultants.dimensions["consultants"]["Менеджер не указан"].revenue == 200
    assert sum(item.revenue for item in consultants.dimensions["consultants"].values()) == 1100


def test_supplier_block_reads_only_dedicated_supplier_hierarchy():
    suppliers = _reports()[SUPPLIERS]
    assert suppliers.dimensions["suppliers"]["Own production"].quantity == 4
    assert suppliers.dimensions["suppliers"][UNKNOWN].revenue == 200
    assert sum(item.revenue for item in suppliers.dimensions["suppliers"].values()) == 1100


def test_three_blocks_cross_validate_against_sales_total():
    reports = _reports()
    assert validate_period_bundle(reports) == []
    control = cross_block_validation(reports)
    assert control["Δ выручки к ПРОД"].abs().sum() == 0
    assert control["Δ количества к ПРОД"].abs().sum() == 0
    assert control["Структура блока сходится"].all()
