from pathlib import Path


def source() -> str:
    return Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")


def test_retail_leader_filter_excludes_outlet_and_63_only_for_leaders():
    text = source()
    assert "def is_tourist_flow_store" in text
    assert 'normalized == "OUTLET" or normalized.startswith("63")' in text
    assert "def retail_leader_summary" in text
    assert 'retail_summary = retail_leader_summary(store_summary)' in text


def test_retail_leader_labels_are_explicit_without_all_network_leader():
    text = source()
    assert '"Лидер розничной сети по выручке"' in text
    assert '"Лидер розничной сети по количеству"' in text
    standard = text[text.index("def render_executive_brief"):text.index("def segment_bar")]
    assert "Лидер всей сети" not in standard


def test_store_charts_use_full_network_but_table_is_collapsible():
    text = source()
    assert 'horizontal_bar(' in text and 'store_summary.head(10)' in text
    assert 'with st.expander("Полная таблица по магазинам"' in text
