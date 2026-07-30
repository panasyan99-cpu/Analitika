from pathlib import Path


def test_silver_registration_and_receipt_are_separate_steps() -> None:
    silver = Path("src/warehouse_management/silver.py").read_text(encoding="utf-8")
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")

    assert "auto_receive = False" in silver
    assert "received=False" in silver
    assert "actual_manual=None" in silver
    assert 'create_label = "Создать поставку в Baserow"' in ui
    assert "Создать Master и провести приход" not in ui
    assert "Приёмка → По поставке" in ui
