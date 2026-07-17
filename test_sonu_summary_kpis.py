from pathlib import Path


def test_sonu_has_complete_management_summary_and_ai_priority():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    start = text.index("def _render_ai_overview")
    end = text.index("@st.fragment\ndef _render_network_section", start)
    summary = text[start:end]

    assert 'st.markdown("### Общий отчет Sonu")' in summary
    assert "st.columns(5)" in summary
    assert '_kpi("Продано SKU"' in summary
    assert '_kpi("Продажи"' in summary
    assert '_kpi("SKU на остатке"' in summary
    assert '_kpi("Средняя цена"' in summary
    assert '"Горизонт заказа", [30, 45, 90]' in summary
    assert "AI-аналитика продаж и заказа" in summary
    assert "Полный отчет по продажам, остаткам и приоритету заказа" in summary


def test_sonu_priority_is_transparent_and_uses_network_stock():
    text = Path("src/sonu.py").read_text(encoding="utf-8")
    assert "def order_priority_report" in text
    assert "Очень нужно заказать" in text
    assert "Нужно заказать" in text
    assert "Желательно заказать" in text
    assert "Плановое пополнение" in text
    assert "Не критично" in text
    assert '"Остаток сети": ("Остаток сети", "max")' in text
    assert '"Продано за период": ("Скорость продаж", "sum")' in text
