from pathlib import Path

ROOT = Path(__file__).parent


def test_sonu_has_stone_breakdown_inside_each_bracelet_type():
    sonu = (ROOT / "src" / "sonu.py").read_text(encoding="utf-8")
    assert "def bracelet_stone_summary" in sonu
    assert '["Тип браслета", "Группа камня", "Камень группы"]' in sonu
    assert "Виды камней" in sonu
    assert 'CENTERED_BRACELET_LABEL = "С затяжкой"' in sonu
    assert 'FULL_CIRCLE_BRACELET_LABEL = "Без затяжки (в круг)"' in sonu
    assert "модели на кольцах" in sonu
    assert "50/50" in sonu
