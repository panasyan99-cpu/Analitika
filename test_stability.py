from pathlib import Path


def source() -> str:
    return Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")


def test_report_parsing_uses_bounded_session_cache():
    text = source()
    assert '@st.cache_resource(scope="session", ttl=900, max_entries=3' in text
    assert 'def parse_report_bundle(' in text
    assert 'report_cache_stores' not in text or 'pop("report_cache_stores"' in text


def test_heavy_interactive_sections_are_fragments():
    text = source()
    assert '@st.fragment\ndef render_store_fragment' in text
    assert '@st.fragment\ndef render_interactive_fragment' in text
    assert '@st.fragment\ndef render_supplier_fragment' in text
    assert 'render_store_fragment(stores)' in text
    assert 'render_interactive_fragment(stores)' in text
    assert 'render_supplier_fragment(supplier_df)' in text


def test_supplier_workbook_is_parsed_once_per_upload():
    text = source()
    assert 'def parse_supplier_report_with_period' in text
    assert 'supplier_units_from_detail(detail, period, path.name)' in text
    assert 'load_supplier_frames' not in text


def test_hidden_detail_panels_are_lazy():
    text = source()
    assert 'st.segmented_control(' in text
    assert 'Детализация магазина' in text
    assert 'Таблица детализации поставщика' in text
    assert 'st.tabs(["Сегменты", "Номенклатурные группы", "Камни", "Полная детализация"])' not in text
