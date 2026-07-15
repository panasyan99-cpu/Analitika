from pathlib import Path


def test_sonu_navigation_is_sidebar_anchor_navigation():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    assert "def _sonu_sidebar_navigation" in text
    assert "def _sonu_mobile_navigation" in text
    assert '<nav class="side-nav sonu-side-nav">' in text
    assert '("sonu-stores", "Продажи по магазинам")' in text
    assert '("sonu-stones", "Камни")' in text
    assert '("sonu-models", "Модели")' in text
    assert "sonu-forecast" not in text
    assert "sonu-order-matrix" not in text
    assert "sonu-recommendations" not in text
    assert 'st.segmented_control(\n        "Навигация по блокам Sonu"' not in text


def test_sonu_renders_fact_sections_in_one_report():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    expected = [
        '_anchor("sonu-stores")',
        '_render_store_section(frame, rate)',
        '_anchor("sonu-average-sales")',
        '_render_average_sales_section(frame, rate, period_days)',
        '_anchor("sonu-stones")',
        '_render_stones_section(frame, rate)',
        '_anchor("sonu-bracelets")',
        '_render_bracelet_section(frame, rate)',
        '_anchor("sonu-models")',
        '_render_models_section(frame, rate)',
    ]
    positions = [text.index(token) for token in expected]
    assert positions == sorted(positions)


def test_sonu_full_excel_export_contains_only_fact_blocks():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    assert "def build_full_sonu_export" in text
    assert '"Группы камней"' in text
    assert '"Камни по группам"' in text
    assert '"Камни по магазинам"' in text
    assert '"Все модели"' in text
    assert '"Исходные данные"' in text
    assert '"Прогноз 30 дней"' not in text
    assert '"Матрица заказа"' not in text
    assert 'st.download_button(' in text
    assert 'file_name=f"Sonu_full_report_' in text
