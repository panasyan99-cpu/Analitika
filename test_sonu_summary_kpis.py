from pathlib import Path


def test_sonu_summary_uses_actual_network_stock():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    summary_start = text.index('    _anchor("sonu-summary")')
    summary_end = text.index('    st.caption(', summary_start)
    summary = text[summary_start:summary_end]

    assert "st.columns(5)" in summary
    assert '_kpi("Остаток сети"' in summary
    assert '_kpi("SKU"' in summary
    assert '_kpi("Расчетный остаток"' not in summary
    assert '_kpi("Нужно на 30 дней"' in summary
    assert '_kpi("Нужно на 45 дней"' in summary
    assert '_kpi("Нужно на 90 дней"' in summary


def test_stock_planning_is_compact_without_forecast_matrix_or_priority():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    assert "def add_order_horizons" in text
    assert 'result[f"Нужно на {horizon} дней"]' in text
    assert "def _forecast_data" not in text
    assert "def _order_matrix_data" not in text
    assert "def _recommendation_data" not in text
    assert "def _render_forecast_section" not in text
    assert "def _render_order_matrix_section" not in text
    assert "def _render_recommendations_section" not in text
    assert "Приоритет" not in text
