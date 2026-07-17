from pathlib import Path

import pandas as pd

from src.sonu import (
    CENTERED_BRACELET_LABEL,
    FULL_CIRCLE_BRACELET_LABEL,
    classify_bracelets,
    network_assortment_summary,
    network_sku_snapshot,
    stock_conflict_details,
)


def _row(store, sku, category, stone, sold, sales, stock):
    return {
        "Магазин": store,
        "Раздел": "Ювелирные",
        "Металл": "Silver",
        "Категория": category,
        "Категория RU": {"Earrings": "Серьги", "Bracelet": "Браслеты"}[category],
        "SKU": sku,
        "Камень": stone,
        "Проба": "B 925",
        "Отгружено": sold,
        "Скорость продаж": sold,
        "Продажи VND": sales,
        "Средняя цена VND": sales / sold if sold else 0,
        "Остаток сети": stock,
        "Фото": None,
    }


def test_network_stock_is_taken_once_while_sales_are_summed():
    frame = pd.DataFrame([
        _row("63", "EES000013-BS", "Earrings", "BLUE SAPPHIRE", 2, 2_000_000, 8),
        _row("SCR", "EES000013-BS", "Earrings", "BLUE SAPPHIRE", 1, 1_000_000, 8),
    ])
    result = network_sku_snapshot(frame, 25_000, 30)
    assert len(result) == 1
    row = result.iloc[0]
    assert row["Остаток сети"] == 8
    assert row["Продано за период"] == 3
    assert row["Магазинов с продажами"] == 2


def test_network_summary_matches_stone_category_business_shape():
    frame = pd.DataFrame([
        _row("63", "EES000013-BS", "Earrings", "BLUE SAPPHIRE", 2, 2_000_000, 8),
        _row("SCR", "EES000013-BS", "Earrings", "BLUE SAPPHIRE", 1, 1_000_000, 8),
        _row("AB", "EES000099-BS", "Earrings", "BLUE SAPPHIRE", 1, 1_200_000, 5),
    ])
    result = network_assortment_summary(frame, 25_000, 30)
    assert len(result) == 1
    row = result.iloc[0]
    assert row["Группа камня"] == "Top Stones"
    assert row["Камень группы"] == "Blue Sapphire"
    assert row["Категория RU"] == "Серьги"
    assert row["SKU моделей"] == 2
    assert row["Остаток сети"] == 13
    assert row["Продано за период"] == 4


def test_conflicting_repeated_stock_is_reported():
    frame = pd.DataFrame([
        _row("63", "EES000013-BS", "Earrings", "BLUE SAPPHIRE", 2, 2_000_000, 8),
        _row("SCR", "EES000013-BS", "Earrings", "BLUE SAPPHIRE", 1, 1_000_000, 9),
    ])
    conflicts = stock_conflict_details(frame)
    assert conflicts["SKU"].tolist() == ["EES000013-BS"]
    assert conflicts.iloc[0]["Значения остатка"] == "8, 9"


def test_bracelets_use_centered_and_full_circle_groups():
    frame = pd.DataFrame([
        _row("63", "SSNB-000042-BS", "Bracelet", "BLUE SAPPHIRE", 2, 2_000_000, 8),
        _row("SCR", "XB-SN-KSB075-MOIS", "Bracelet", "MOISSANITE", 1, 2_000_000, 4),
    ])
    result = classify_bracelets(frame, 25_000, 30).set_index("SKU")
    assert result.loc["SSNB-000042-BS", "Тип браслета"] == CENTERED_BRACELET_LABEL
    assert result.loc["XB-SN-KSB075-MOIS", "Тип браслета"] == FULL_CIRCLE_BRACELET_LABEL


def test_sonu_has_global_table_card_switch():
    source = Path("src/sonu.py").read_text(encoding="utf-8")
    assert '"Вид представления"' in source
    assert '["Карточки", "Таблица"]' in source
    assert "def _render_assortment_cards" in source
    assert "def _render_model_cards" in source
