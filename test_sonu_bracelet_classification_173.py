import pandas as pd

from src.sonu import bracelet_classification_audit


def _frame():
    return pd.DataFrame([
        {"SKU": "SSNB-AFG018-BS", "Категория": "Bracelet", "Категория RU": "Браслеты", "Камень": "Blue Sapphire", "Магазин": "TT", "Скорость продаж": 8, "Продажи VND": 8000000, "Остаток сети": 2, "Проба": "B 925"},
        {"SKU": "SSNB-AFG018-BS", "Категория": "Bracelet", "Категория RU": "Браслеты", "Камень": "Blue Sapphire", "Магазин": "SCR", "Скорость продаж": 2, "Продажи VND": 2000000, "Остаток сети": 2, "Проба": "B 925"},
        {"SKU": "UNKNOWN-A", "Категория": "Bracelet", "Категория RU": "Браслеты", "Камень": "Ruby", "Магазин": "TT", "Скорость продаж": 6, "Продажи VND": 6000000, "Остаток сети": 1, "Проба": "B 925"},
        {"SKU": "UNKNOWN-B", "Категория": "Bracelet", "Категория RU": "Браслеты", "Камень": "Moissanite", "Магазин": "TT", "Скорость продаж": 2, "Продажи VND": 2000000, "Остаток сети": 5, "Проба": "B 925"},
    ])


def test_audit_keeps_unseen_models_pending_without_hidden_5050():
    summary, detail = bracelet_classification_audit(_frame(), 25000, 30)
    pending = detail[detail["Статус классификации"] == "Требует классификации"]
    assert len(pending) == 2
    assert pending["Тип браслета"].isna().all()
    assert set(pending["Источник классификации"]) == {"Требует классификации"}
    assert summary["Моделей"].sum() == detail["SKU"].nunique()
