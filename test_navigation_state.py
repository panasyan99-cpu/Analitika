from pathlib import Path


ROOT = Path(__file__).parent


def test_production_navigation_matches_switchable_workspaces():
    text = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    assert 'id="executive"' in text
    assert 'id="workspace"' in text
    assert 'id="comparison-workspace"' in text
    assert 'render_standard_workspace(stores, summary_df, supplier_df)' in text
    assert 'key="comparison_workspace"' in text


def test_about_program_is_a_separate_mode_with_three_levels():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    product = (ROOT / "src" / "product_info.py").read_text(encoding="utf-8")
    assert '"О программе"' in product
    assert 'def render_about_mode()' in app
    assert '("О программе", "Руководство", "История обновлений")' in app
    assert 'USER_GUIDE.md' in app
    assert 'release_history_html' in app


def test_section_analytics_are_switchable_and_tables_collapsible():
    text = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    assert '"Обзор"' in text
    assert '"Магазины"' in text
    assert '"Камни и группы"' in text
    assert '"Поставщики"' in text
    assert '"Исследование данных"' in text
    assert 'with st.expander("Полная таблица по магазинам"' in text
