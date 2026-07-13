import pandas as pd

from streamlit_app import formatted_table, table_column_config


def test_formatted_table_preserves_numeric_dtypes_for_frontend_sorting():
    source = pd.DataFrame(
        {
            "Поставщик": ["A", "B", "C"],
            "Количество": [2, 100, 11],
            "Выручка": [9000.0, 100.0, 500.0],
            "% выручки": [0.9, 0.01, 0.05],
        }
    )
    display = formatted_table(source)

    assert pd.api.types.is_numeric_dtype(display["Количество"])
    assert pd.api.types.is_numeric_dtype(display["Выручка"])
    assert pd.api.types.is_numeric_dtype(display["% выручки"])
    assert display.sort_values("Количество")["Поставщик"].tolist() == ["A", "C", "B"]
    assert display.sort_values("Выручка")["Поставщик"].tolist() == ["B", "C", "A"]


def test_numeric_formatting_is_applied_through_column_config_not_strings():
    frame = pd.DataFrame(
        {
            "Количество": [10],
            "Выручка": [123456.0],
            "Средняя стоимость": [12345.6],
            "% выручки": [0.25],
        }
    )
    config = table_column_config(frame)
    assert set(config) == set(frame.columns)
