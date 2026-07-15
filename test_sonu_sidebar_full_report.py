from pathlib import Path


def test_sonu_navigation_is_sidebar_anchor_navigation():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    assert "def _sonu_sidebar_navigation" in text
    assert '<nav class="side-nav sonu-side-nav">' in text
    assert '("sonu-stores", "Продажи по магазинам")' in text
    assert '("sonu-recommendations", "Рекомендации")' in text
    assert 'st.segmented_control(\n        "Навигация по блокам Sonu"' not in text
    assert 'st.sidebar.markdown("[Сводка](#sonu-summary)")' not in text


def test_sonu_renders_all_main_sections_in_one_report():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    expected = [
        '_anchor("sonu-stores")',
        '_render_store_section(frame, rate)',
        '_anchor("sonu-average-sales")',
        '_render_average_sales_section(frame, rate, period_days)',
        '_anchor("sonu-bracelets")',
        '_render_bracelet_section(frame, rate)',
        '_anchor("sonu-forecast")',
        '_render_forecast_section(frame, rate, period_days)',
        '_anchor("sonu-order-matrix")',
        '_render_order_matrix_section(frame, rate, period_days)',
        '_anchor("sonu-recommendations")',
        '_render_recommendations_section(frame, rate, period_days)',
    ]
    positions = [text.index(token) for token in expected]
    assert positions == sorted(positions)


def test_sonu_full_excel_export_is_available():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    assert "def build_full_sonu_export" in text
    assert '"Прогноз 30 дней"' in text
    assert '"Прогноз 60 дней"' in text
    assert '"Прогноз 90 дней"' in text
    assert '"Матрица заказа"' in text
    assert '"Исходные данные"' in text
    assert 'st.download_button(' in text
    assert 'file_name=f"Sonu_full_report_' in text
