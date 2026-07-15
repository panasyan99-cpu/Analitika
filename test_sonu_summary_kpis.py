from pathlib import Path


def test_sonu_summary_does_not_show_calculated_stock_card():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    summary_start = text.index('    _anchor("sonu-summary")')
    summary_end = text.index('    st.caption(', summary_start)
    summary = text[summary_start:summary_end]

    assert 'st.columns(5)' in summary
    assert '_kpi("Расчетный остаток"' not in summary
    assert '_kpi("Моделей"' in summary


def test_calculated_stock_remains_available_for_order_planning():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    assert 'frame["Расчетный остаток"]' in text
    assert 'Матрица использует расчетный остаток' in text
