from pathlib import Path


def source() -> str:
    return Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")


def test_retail_leader_filter_excludes_outlet_and_63_only_for_leaders():
    text = source()
    assert "def is_tourist_flow_store" in text
    assert 'normalized == "OUTLET" or normalized.startswith("63")' in text
    assert "def retail_leader_summary" in text
    assert 'retail_summary = retail_leader_summary(store_summary)' in text


def test_retail_leader_labels_are_explicit_without_internal_exclusion_note():
    text = source()
    assert '"Лидер розничной сети по выручке"' in text
    assert '"Лидер розничной сети по количеству"' in text
    assert "без OUTLET и 63" not in text


def test_full_store_summary_is_still_used_for_charts_and_table():
    text = source()
    assert 'horizontal_bar(' in text and 'store_summary' in text
    assert 'data_table(store_summary, key="executive_store_table")' in text
