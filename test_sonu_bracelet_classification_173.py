import pandas as pd

from src.sonu import (
    CENTERED_BRACELET_LABEL,
    FULL_CIRCLE_BRACELET_LABEL,
    bracelet_classification_audit,
)


def _frame():
    return pd.DataFrame([
        {"SKU": "SSNB-AFG018-BS", "Категория": "Bracelet", "Категория RU": "Браслеты", "Камень": "Blue Sapphire", "Магазин": "TT", "Скорость продаж": 8, "Продажи VND": 8000000, "Остаток сети": 2, "Проба": "B 925"},
        {"SKU": "SSNB-AFG018-BS", "Категория": "Bracelet", "Категория RU": "Браслеты", "Камень": "Blue Sapphire", "Магазин": "SCR", "Скорость продаж": 2, "Продажи VND": 2000000, "Остаток сети": 2, "Проба": "B 925"},
        {"SKU": "UNKNOWN-A", "Категория": "Bracelet", "Категория RU": "Браслеты", "Камень": "Ruby", "Магазин": "TT", "Скорость продаж": 6, "Продажи VND": 6000000, "Остаток сети": 1, "Проба": "B 925"},
        {"SKU": "UNKNOWN-B", "Категория": "Bracelet", "Категория RU": "Браслеты", "Камень": "Moissanite", "Магазин": "TT", "Скорость продаж": 2, "Продажи VND": 2000000, "Остаток сети": 5, "Проба": "B 925"},
    ])


def test_audit_exposes_ambiguous_models_and_assignment():
    summary, detail = bracelet_classification_audit(_frame(), 25000, 30)
    ambiguous = detail[detail["Статус классификации"] == "Спорная модель"]
    assert len(ambiguous) == 2
    assert set(ambiguous["Тип браслета"]) == {CENTERED_BRACELET_LABEL, FULL_CIRCLE_BRACELET_LABEL}
    assert "Правило 50/50" in set(ambiguous["Источник классификации"])
    assert summary["Моделей"].sum() == detail["SKU"].nunique()
