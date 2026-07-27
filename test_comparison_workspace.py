from pathlib import Path

import pandas as pd

from streamlit_app import annotate_change_status, compare_metric_frames, period_days


def source() -> str:
    return Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")


def test_separate_standard_and_comparison_uploads_are_present():
    text = source()
    assert '"Обычный отчет"' in text
    assert '"Сравнение периодов"' in text
    assert 'key="upload_widget"' in text
    assert 'key="comparison_upload_1"' in text
    assert 'key="comparison_upload_2"' in text
    assert '"Запустить сравнительный анализ"' in text
    assert 'with st.form("comparison_upload_form"' in text


def test_comparison_uses_one_switchable_workspace():
    text = source()
    assert 'id="comparison-workspace"' in text
    assert 'key="comparison_workspace"' in text
    for label in (
        "Итог изменений", "Драйверы", "Магазины", "Камни и группы",
        "Металлы и пробы", "Поставщики", "Исследование данных",
    ):
        assert f'"{label}"' in text
    assert "render_comparison_drivers_fragment" in text
    assert "render_comparison_stones_groups_fragment" in text


def test_comparison_table_uses_outer_join_and_numeric_deltas():
    first = pd.DataFrame([
        {"Магазин": "AB", "Количество": 10, "Выручка": 1000, "Средняя стоимость": 100},
        {"Магазин": "SCR", "Количество": 4, "Выручка": 400, "Средняя стоимость": 100},
    ])
    second = pd.DataFrame([
        {"Магазин": "AB", "Количество": 12, "Выручка": 1440, "Средняя стоимость": 120},
        {"Магазин": "NTR1", "Количество": 3, "Выручка": 450, "Средняя стоимость": 150},
    ])
    result = compare_metric_frames(first, second, ["Магазин"])
    assert set(result["Магазин"]) == {"AB", "SCR", "NTR1"}
    ab = result[result["Магазин"] == "AB"].iloc[0]
    assert ab["Δ количества"] == 2
    assert ab["Δ выручки"] == 440
    assert round(float(ab["Δ выручки %"]), 2) == 0.44
    assert pd.api.types.is_numeric_dtype(result["Выручка · Период 1"])
    assert pd.api.types.is_numeric_dtype(result["Выручка · Период 2"])


def test_new_and_disappeared_statuses_are_explicit():
    frame = pd.DataFrame([
        {"Группа": "new", "Выручка · Период 1": 0, "Выручка · Период 2": 10},
        {"Группа": "lost", "Выручка · Период 1": 10, "Выручка · Период 2": 0},
        {"Группа": "flat", "Выручка · Период 1": 100, "Выручка · Период 2": 102},
    ])
    result = annotate_change_status(frame, 3).set_index("Группа")
    assert result.loc["new", "Статус"] == "Новая группа"
    assert result.loc["lost", "Статус"] == "Исчезла из продаж"
    assert result.loc["flat", "Статус"] == "Без существенного изменения"


def test_period_days_are_inclusive():
    from datetime import date
    assert period_days(date(2026, 4, 1), date(2026, 4, 30)) == 30
