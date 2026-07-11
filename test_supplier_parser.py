from pathlib import Path

import pandas as pd

from streamlit_app import supplier_summary


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
