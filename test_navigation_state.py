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
    assert 'def sidebar_navigation(has_report: bool, *, comparison: bool = False)' in source
    assert 'if has_report:' in source
    assert 'side-nav a:visited' in source
    assert 'text-decoration:none !important' in source
    assert '#f1cc85' in source


def test_about_program_content_is_generated_from_single_registry():
    source = Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")
    product_info = Path(__file__).with_name("src").joinpath("product_info.py").read_text(encoding="utf-8")
    changelog = Path(__file__).with_name("CHANGELOG.md").read_text(encoding="utf-8")
    assert "feature_cards_html()" in source
    assert 'release_history_html(Path(__file__).with_name("CHANGELOG.md"))' in source
    assert "Обычная аналитика продаж" in product_info
    assert "Сравнение периодов" in product_info
    assert "Складская аналитика Baserow" in product_info
    assert "Заказ Sonu" in product_info
    assert "Как настроить отчет" not in source
    assert "Как подготовить отчет" not in source
    assert "## 1.2.0 — Baserow warehouse analytics" in changelog
    assert "## 1.2.1 — Fix warehouse chart sizing and label clipping" in changelog


def test_section_analytics_are_present():
    source = Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")
    assert 'section_divider("Сводка по сети"' in source
    assert 'insight_panel("Аналитика по сети"' in source
    assert '"Аналитика по магазину"' in source
    assert '"Аналитика по выбранным параметрам"' in source
    assert '"Аналитика по поставщикам"' in source
