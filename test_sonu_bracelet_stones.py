from pathlib import Path

ROOT = Path(__file__).parent


def test_sonu_has_stone_breakdown_inside_each_bracelet_type():
    sonu = (ROOT / "src" / "sonu.py").read_text(encoding="utf-8")
    assert "def bracelet_stone_summary" in sonu
    assert '["Тип браслета", "Группа камня", "Камень группы"]' in sonu
    assert "Доля продаж внутри типа" in sonu
    assert "Виды камней" in sonu
    assert "С центральной композицией" in sonu
    assert "Полный круг" in sonu
    assert "модели на кольцах" in sonu
    assert "Правило 50/50" in sonu
