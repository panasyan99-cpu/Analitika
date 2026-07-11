from pathlib import Path


def test_one_page_structure_and_unique_widget_namespaces():
    source = Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")
    assert 'id="summary"' in source
    assert 'id="stores"' in source
    assert 'id="interactive"' in source
    assert 'id="suppliers"' in source
    assert 'id="export"' in source
    assert 'namespace="main_interactive"' in source
    assert 'interactive_explorer(store, all_stores)' not in source
    assert 'key="store_page_select"' in source
    assert 'key="interactive_store_select"' in source
    assert 'key="supplier_selected"' in source
    assert 'key="prepare_export"' in source


def test_section_analytics_are_present():
    source = Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")
    assert "section_divider('Сводка по сети'" in source
    assert "insight_panel('Аналитика по сети'" in source
    assert '"Аналитика по магазину"' in source
    assert '"Аналитика по выбранным параметрам"' in source
    assert '"Аналитика по поставщикам"' in source
