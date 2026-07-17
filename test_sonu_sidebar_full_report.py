from pathlib import Path


def test_sonu_navigation_is_sidebar_anchor_navigation():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    assert "def _sonu_sidebar_navigation" in text
    assert "def _sonu_mobile_navigation" in text
    assert "return render_sidebar(" in text
    assert "items=sonu_navigation_items(has_report)" in text
    assert '("sonu-network", "Аналитика сети")' in text
    assert '("sonu-bracelets", "Браслеты")' in text
    assert '("sonu-stores", "Продажи по магазинам")' in text
    assert '("sonu-models", "Модели")' in text
    assert "sonu-forecast" not in text
    assert "sonu-order-matrix" not in text
    assert "sonu-recommendations" not in text


def test_sonu_renders_network_sections_in_one_report():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    expected = [
        '_anchor("sonu-network")',
        '_render_network_section(frame, rate, period_days, view_mode)',
        '_anchor("sonu-bracelets")',
        '_render_bracelet_section(frame, rate, period_days, view_mode)',
        '_anchor("sonu-stores")',
        '_render_store_section(frame, rate, view_mode)',
        '_anchor("sonu-models")',
        '_render_models_section(frame, rate, period_days, view_mode)',
    ]
    positions = [text.index(token) for token in expected]
    assert positions == sorted(positions)


def test_sonu_full_excel_export_contains_network_stock_blocks():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    assert "def build_full_sonu_export" in text
    assert '"Сеть по группам"' in text
    assert '"Сеть камни-категории"' in text
    assert '"SKU сети"' in text
    assert '"Контроль остатков"' in text
    assert '"Типы браслетов"' in text
    assert '"Модели браслетов"' in text
    assert '"Прогноз 30 дней"' not in text
    assert '"Матрица заказа"' not in text
    assert "st.download_button(" in text
    assert 'file_name=f"Sonu_network_stock_' in text
