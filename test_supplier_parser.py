from pathlib import Path

import pandas as pd

from streamlit_app import supplier_pie_data, supplier_summary


def test_service_supplier_is_merged_into_other():
    detail = pd.DataFrame(
        {
            "Поставщик": ["Сеть", "", "Sonu"],
            "Количество": [2, 3, 5],
            "Выручка": [200.0, 300.0, 1000.0],
        }
    )
    summary = supplier_summary(detail)
    names = set(summary["Поставщик"])
    assert "Сеть" not in names
    assert "Other" in names
    assert int(summary["Количество"].sum()) == 10


def test_supplier_pie_merges_below_4_5_percent_into_other():
    summary = pd.DataFrame(
        {
            "Поставщик": ["Taiwan", "Sonu", "Small A", "Small B", "Other"],
            "% выручки": [0.63, 0.08, 0.044, 0.01, 0.236],
            "% количества": [0.42, 0.0473, 0.0449, 0.01, 0.4778],
        }
    )

    revenue_labels, revenue_values = supplier_pie_data(summary, "% выручки")
    quantity_labels, quantity_values = supplier_pie_data(summary, "% количества")

    assert revenue_labels == ["Taiwan", "Sonu", "Other"]
    assert abs(revenue_values[-1] - (0.044 + 0.01 + 0.236)) < 1e-9

    assert quantity_labels == ["Taiwan", "Sonu", "Other"]
    assert abs(quantity_values[-1] - (0.0449 + 0.01 + 0.4778)) < 1e-9


def test_supplier_pie_keeps_exactly_4_5_percent_separate():
    summary = pd.DataFrame(
        {
            "Поставщик": ["Borderline", "Small"],
            "% выручки": [0.045, 0.01],
        }
    )
    labels, values = supplier_pie_data(summary, "% выручки")
    assert labels == ["Borderline", "Other"]
    assert values == [0.045, 0.01]
