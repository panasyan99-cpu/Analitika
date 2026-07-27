from pathlib import Path


def source() -> str:
    return Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")


def test_executive_brief_and_compact_workspace_are_present_without_sidebar_navigation():
    text = source()
    assert '<div id="executive"></div>' in text
    assert '<div id="workspace"></div>' in text
    assert "render_executive_brief(stores, summary_df, supplier_df)" in text
    assert "render_standard_workspace(stores, summary_df, supplier_df)" in text
    assert 'def sidebar_navigation(' not in text


def test_executive_brief_has_retail_management_metrics_without_large_tables():
    text = source()
    assert "def executive_store_summary" in text
    assert "def network_segment_summary" in text
    assert "def executive_insights" in text
    assert '"Лидер розничной сети по выручке"' in text
    assert '"Лидер розничной сети по количеству"' in text
    assert '"Главный сегмент по выручке"' in text
    brief = text[text.index("def render_executive_brief"):text.index("def segment_bar")]
    assert "data_table(" not in brief
    assert "locked_plotly_chart(" not in brief


def test_sellers_are_not_part_of_the_product():
    text = source().casefold()
    assert "продавц" not in text
    assert "seller" not in text
