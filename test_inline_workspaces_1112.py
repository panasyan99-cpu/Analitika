from pathlib import Path

ROOT = Path(__file__).parent


def test_all_workspaces_use_inline_ui_without_active_sidebar_calls():
    sonu = (ROOT / "src" / "sonu.py").read_text(encoding="utf-8")
    warehouse = (ROOT / "src" / "warehouse.py").read_text(encoding="utf-8")
    order = (ROOT / "src" / "order_workflow.py").read_text(encoding="utf-8")

    sonu_body = sonu[sonu.index("def render_sonu_order_dashboard"):]
    warehouse_body = warehouse[warehouse.index("def render_warehouse_dashboard"):]
    order_body = order[order.index("def render_supplier_order_dashboard"):]

    assert "_sonu_sidebar_navigation(" not in sonu_body
    assert "_sonu_mobile_navigation(" not in sonu_body
    assert "status_slot = render_navigation()" not in warehouse_body
    assert "_render_sidebar(parsed, draft)" not in order_body
    assert '"Загрузить другой отчет"' in sonu_body
    assert '"Сохранить сейчас"' in order_body
    warehouse_ui = (ROOT / "src" / "warehouse_management" / "ui.py").read_text(encoding="utf-8")
    assert "Princess Warehouse Online" in warehouse_ui
    assert "загружается только выбранный подраздел" in warehouse_ui


def test_course_and_purities_is_the_only_current_settings_label():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    guide = (ROOT / "USER_GUIDE.md").read_text(encoding="utf-8")
    assert 'with st.expander("⚙️ Курс и пробы", expanded=False)' in app
    assert '"Заказ поставщику",' in app[app.index("def render_report_settings"):app.index("def render_mode_help_page")]
    assert "Настройки отчёта" not in app
    assert "Настройки отчёта" not in guide
    assert "## 3. Курс и пробы" in guide


def test_comparison_upload_has_no_duplicated_intro_panel():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    body = app[app.index("def render_comparison_mode"):app.index("def render_warehouse_mode")]
    assert '<div class="upload-panel"><b>Сравнение двух периодов</b>' not in body
    assert "Сравнение запустится только после отправки сразу двух файлов." not in body
    assert 'with st.form("comparison_upload_form"' in body


def test_management_guide_is_expanded_for_three_operational_modules():
    guide = (ROOT / "USER_GUIDE.md").read_text(encoding="utf-8")
    assert len(guide.splitlines()) > 1150
    for heading in (
        "### Приоритет правил",
        "### Что руководитель должен проверить перед Excel",
        "### Как агрегируются данные Sonu",
        "### Как проверять рекомендацию",
        "### Связь Baserow и Princess Supply Manager",
        "### Проверка качества складских данных",
    ):
        assert heading in guide
