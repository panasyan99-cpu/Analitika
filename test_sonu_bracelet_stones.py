from pathlib import Path

ROOT = Path(__file__).parent


def test_sonu_has_stone_breakdown_inside_each_bracelet_type():
    sonu = (ROOT / "src" / "sonu.py").read_text(encoding="utf-8")
    assert "def bracelet_stone_summary" in sonu
    assert '["Тип браслета", "Камень"]' in sonu
    assert "Доля продаж внутри типа" in sonu
    assert "Камни внутри типов браслетов" in sonu
    assert "С затяжкой" in sonu
    assert "В круг с камнями" in sonu
    assert "sonu_bracelet_stone_qty_" in sonu
    assert "sonu_bracelet_stone_sales_" in sonu
