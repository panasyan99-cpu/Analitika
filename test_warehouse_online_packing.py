from pathlib import Path

from openpyxl import Workbook

from src.warehouse_management.packing import load_products


def test_online_master_parser_supports_blank_boxes(tmp_path: Path):
    path = tmp_path / "master.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["№", "Артикул", "Коробки", "По документу", "Материал", "Камень", "Получено", "Факт"])
    sheet.append([1, "SKU-NEW", "", 12, "Steel", "Lapis Lazuli", "Да", 10])
    workbook.save(path)

    products = load_products(path, tmp_path / "images")
    assert len(products) == 1
    product = products[0]
    assert product.sku == "SKU-NEW"
    assert product.boxes == ""
    assert product.qty_document == 12
    assert product.actual_qty == 10
    assert product.stone == "Lapis Lazurite"
