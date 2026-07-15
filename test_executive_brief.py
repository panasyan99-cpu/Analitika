from pathlib import Path


def source() -> str:
    return Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")


def test_executive_navigation_and_section_are_present():
    text = source()
    assert '("executive", "Оперативная сводка", "#executive", has_report)' in text
    assert '<div id="executive"></div>' in text
    assert "render_executive_brief(stores, summary_df, supplier_df)" in text


def test_executive_brief_has_management_metrics():
    text = source()
    assert "def executive_store_summary" in text
    assert "def network_segment_summary" in text
    assert "def executive_insights" in text
    assert '"Лидер розничной сети по выручке"' in text
    assert '"Доля топ-3 поставщиков"' in text
    assert 'Магазины одним взглядом' in text
    assert '"Главный сегмент по выручке"' in text
    assert 'def leader_kpi_card' in text


def test_sellers_are_not_part_of_the_product():
    text = source().casefold()
    assert "продавц" not in text
    assert "seller" not in text
