from pathlib import Path

import pandas as pd

from src.sonu import add_stone_classification, stone_group_summary, stone_member_summary


def sample_frame() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"SKU": "SSNB-000042-BS", "Камень": "BS", "Магазин": "TT", "Скорость продаж": 5, "Продажи VND": 5_000_000},
            {"SKU": "XB-SN-KSB075-MOIS", "Камень": "MOIS", "Магазин": "NTR1", "Скорость продаж": 2, "Продажи VND": 4_000_000},
            {"SKU": "XB-SN-000024-FPW", "Камень": "FPW", "Магазин": "AB", "Скорость продаж": 3, "Продажи VND": 1_500_000},
            {"SKU": "SSNB-AFG019-CIT", "Камень": "CIT", "Магазин": "TT", "Скорость продаж": 4, "Продажи VND": 2_000_000},
        ]
    )


def test_sonu_abbreviations_map_to_business_groups():
    result = add_stone_classification(sample_frame())
    lookup = result.set_index("Сокращение")[["Группа камня", "Камень группы"]].to_dict("index")
    assert lookup["BS"] == {"Группа камня": "Top Stones", "Камень группы": "Blue Sapphire"}
    assert lookup["MOIS"] == {"Группа камня": "Top Stones", "Камень группы": "Moissanite"}
    assert lookup["FPW"] == {"Группа камня": "Pearls", "Камень группы": "White Freshwater Pearl"}
    assert lookup["CIT"] == {"Группа камня": "Other Stones", "Камень группы": "Quartz Group"}


def test_sonu_group_and_member_summaries_are_complete():
    groups = stone_group_summary(sample_frame(), 25_000)
    members = stone_member_summary(sample_frame(), 25_000)
    assert groups["Группа камня"].tolist() == ["Top Stones", "Pearls", "Other Stones"]
    assert int(groups["Скорость продаж"].sum()) == 14
    assert {"Blue Sapphire", "Moissanite", "White Freshwater Pearl", "Quartz Group"}.issubset(set(members["Камень группы"]))


def test_source_contains_all_three_group_labels():
    source = Path("src/sonu.py").read_text(encoding="utf-8")
    for label in ["Top Stones", "Pearls", "Other Stones"]:
        assert label in source
