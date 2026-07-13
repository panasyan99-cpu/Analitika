from pathlib import Path


def test_production_navigation_and_scope():
    source = Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")
    assert 'id="executive"' in source
    assert 'id="summary"' in source
    assert 'id="stores"' in source
    assert 'id="interactive"' in source
    assert 'id="suppliers"' in source
    assert 'id="about"' in source
    assert 'id="export"' not in source
    assert '#export' not in source
    assert 'prepare_export' not in source
    assert 'build_excel(' not in source
    assert 'namespace="main_interactive"' in source
    assert 'key="store_page_select"' in source
    assert 'key="interactive_store_select"' in source
    assert 'key="supplier_selected"' in source


def test_sidebar_is_conditional_and_styled():
    source = Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")
    assert 'def sidebar_navigation(has_report: bool)' in source
    assert 'if has_report:' in source
    assert 'side-nav a:visited' in source
    assert 'text-decoration:none !important' in source
    assert '#f1cc85' in source


def test_about_platform_content():
    source = Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")
    assert 'Продажи товаров' in source
    assert 'Номенклатурная группа' in source
    assert 'Камень / вставка' in source
    assert 'Обновления' in source
    assert 'Analitika Web 1.1.10 — Executive brief clarity' in source
    assert 'Analitika Web 1.1.7 — Stability and memory optimization' in source
    assert 'Analitika Web 1.1.6 — Responsive mobile layout' in source
    assert 'Analitika Web 1.1.5 — Locked chart interactions' in source
    assert 'Analitika Web 1.1.4 — Release history' in source
    assert 'Analitika Web 1.1.3 — Group small suppliers in pie charts' in source
    assert 'Analitika Web 1.1.2 — Fix chart label clipping' in source
    assert 'В следующих версиях' not in source


def test_section_analytics_are_present():
    source = Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")
    assert "section_divider('Сводка по сети'" in source
    assert "insight_panel('Аналитика по сети'" in source
    assert '"Аналитика по магазину"' in source
    assert '"Аналитика по выбранным параметрам"' in source
    assert '"Аналитика по поставщикам"' in source
