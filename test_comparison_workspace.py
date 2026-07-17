from pathlib import Path

import pandas as pd

from streamlit_app import compare_metric_frames


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
    assert 'st.form_submit_button(' in text


def test_comparison_has_network_store_interactive_and_supplier_sections():
    text = source()
    assert 'id="comparison-summary"' in text
    assert 'id="comparison-metals"' in text
    assert 'id="comparison-stores"' in text
    assert 'id="comparison-interactive"' in text
    assert 'id="comparison-suppliers"' in text
    assert 'render_comparison_metal_section' in text
    assert 'render_comparison_store_fragment' in text
    assert 'render_comparison_interactive_fragment' in text
    assert 'render_comparison_supplier_fragment' in text


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
