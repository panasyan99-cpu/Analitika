from pathlib import Path


def test_sonu_navigation_is_sidebar_anchor_navigation():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    assert "def _sonu_sidebar_navigation" in text
    assert "def _sonu_mobile_navigation" in text
    assert "return render_sidebar(" in text
    assert "items=sonu_navigation_items(has_report)" in text
    assert '("sonu-summary", "Общий отчет и AI", "#sonu-summary", has_report)' in text
    assert '("sonu-network", "Изделия без браслетов")' in text
    assert '("sonu-bracelets", "Браслеты")' in text
    assert '("sonu-export", "Полная выгрузка", "#sonu-export", has_report)' in text
    assert "sonu-forecast" not in text
    assert "sonu-order-matrix" not in text


def test_sonu_renders_final_sections_in_one_report():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    expected = [
        '_anchor("sonu-summary")',
        '_render_ai_overview(frame, rate, period_days, report.period)',
        '_anchor("sonu-network")',
        '_render_network_section(frame, rate, period_days)',
        '_anchor("sonu-bracelets")',
        '_render_bracelet_section(frame, rate, period_days)',
        '_anchor("sonu-export")',
    ]
    positions = [text.index(token) for token in expected]
    assert positions == sorted(positions)


def test_sonu_full_excel_export_contains_final_business_blocks():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    assert "def build_full_sonu_export" in text
    assert '"AI приоритет заказа"' in text
    assert '"Общий отчет"' in text
    assert '"Без браслетов"' in text
    assert '"Браслеты с затяжкой"' in text
    assert '"Браслеты без затяжки"' in text
    assert '"SKU сети"' in text
    assert '"Контроль остатков"' in text
    assert "st.download_button(" in text
    assert 'file_name=f"Sonu_AI_order_report_' in text
