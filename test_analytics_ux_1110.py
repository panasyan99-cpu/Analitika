from pathlib import Path

import pandas as pd

from streamlit_app import (
    annotate_change_status,
    network_conclusions,
    retail_leader_summary,
    supplier_has_meaningful_detail,
)


ROOT = Path(__file__).parent


def test_retail_leader_summary_excludes_outlet_and_63():
    frame = pd.DataFrame([
        {"Магазин": "OUTLET", "Выручка": 1000, "Количество": 10},
        {"Магазин": "63", "Выручка": 900, "Количество": 9},
        {"Магазин": "AB", "Выручка": 500, "Количество": 5},
        {"Магазин": "NTR1", "Выручка": 400, "Количество": 4},
    ])
    result = retail_leader_summary(frame)
    assert result["Магазин"].tolist() == ["AB", "NTR1"]


def test_network_conclusions_use_weighted_segment_share():
    frame = pd.DataFrame([
        {
            "Магазин": "AB", "Выручка": 900, "Количество": 9,
            "Top Stones — продажи %": 1.0,
            "Pearls — продажи %": 0.0,
            "Other Stones — продажи %": 0.0,
        },
        {
            "Магазин": "NTR1", "Выручка": 100, "Количество": 1,
            "Top Stones — продажи %": 0.0,
            "Pearls — продажи %": 1.0,
            "Other Stones — продажи %": 0.0,
        },
    ])
    lines = network_conclusions(frame)
    assert any("Top Stones" in line and "90,00%" in line for line in lines)


def test_supplier_other_only_is_not_meaningful_detail():
    only_other = pd.DataFrame([
        {"Поставщик": "Other", "Количество": 2, "Выручка": 100},
    ])
    real = pd.DataFrame([
        {"Поставщик": "Sonu", "Количество": 2, "Выручка": 100},
    ])
    assert supplier_has_meaningful_detail(only_other) is False
    assert supplier_has_meaningful_detail(real) is True


def test_change_threshold_keeps_new_and_disappeared_visible():
    frame = pd.DataFrame([
        {"name": "new", "Выручка · Период 1": 0, "Выручка · Период 2": 1},
        {"name": "lost", "Выручка · Период 1": 1, "Выручка · Период 2": 0},
        {"name": "small", "Выручка · Период 1": 100, "Выручка · Период 2": 104},
    ])
    result = annotate_change_status(frame, 5).set_index("name")
    assert result.loc["new", "Статус"] == "Новая группа"
    assert result.loc["lost", "Статус"] == "Исчезла из продаж"
    assert result.loc["small", "Статус"] == "Без существенного изменения"


def test_about_is_not_appended_to_working_modules():
    text = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    for name in (
        "render_standard_report_mode",
        "render_comparison_mode",
        "render_warehouse_mode",
        "render_sonu_mode",
        "render_supplier_order_mode",
    ):
        start = text.index(f"def {name}(")
        next_def = text.find("\ndef ", start + 5)
        body = text[start: next_def if next_def != -1 else len(text)]
        assert "render_about()" not in body


def test_user_guide_and_release_documents_exist():
    for name in (
        "USER_GUIDE.md",
        "RELEASE_NOTES_1.11.0.md",
        "DEPLOY_1.11.0.md",
        "VALIDATION_1.11.0.md",
    ):
        assert (ROOT / name).is_file()
