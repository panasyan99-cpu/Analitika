from pathlib import Path


def test_sonu_summary_does_not_show_calculated_stock_card():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    summary_start = text.index('    _anchor("sonu-summary")')
    summary_end = text.index('    st.caption(', summary_start)
    summary = text[summary_start:summary_end]

    assert 'st.columns(5)' in summary
    assert '_kpi("Расчетный остаток"' not in summary
    assert '_kpi("Моделей"' in summary


def test_stock_based_planning_is_removed_until_actual_inventory_arrives():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    assert "def _forecast_data" not in text
    assert "def _order_matrix_data" not in text
    assert "def _recommendation_data" not in text
    assert "def _render_forecast_section" not in text
    assert "def _render_order_matrix_section" not in text
    assert "def _render_recommendations_section" not in text
    assert "Рекомендованный заказ" not in text
