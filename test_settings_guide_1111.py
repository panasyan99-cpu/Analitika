from pathlib import Path

ROOT = Path(__file__).parent


def test_every_operational_mode_has_work_and_help_tabs():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    assert 'options = ("Работа", "Как с этим работать")' in app
    assert 'def render_mode_workspace_tab(mode: str) -> bool:' in app
    for mode, chapter in {
        "Обычный отчет": "1. Общий анализ продаж",
        "Сравнение периодов": "2. Сравнение периодов",
        "Сувениры и касты на складе": "6. Склад Baserow",
        "Заказ Sonu": "5. Заказ Sonu",
        "Заказ поставщику": "4. Заказ поставщику",
        "Управленческий отчет": "7. Управленческий отчет",
    }.items():
        assert f'"{mode}": "{chapter}"' in app


def test_sales_and_comparison_sidebars_are_removed():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    standard = app[app.index('def render_standard_report_mode() -> None:'):app.index('def render_comparison_mode() -> None:')]
    comparison = app[app.index('def render_comparison_mode() -> None:'):app.index('def render_warehouse_mode() -> None:')]
    assert 'sidebar_navigation(' not in standard
    assert 'mobile_navigation(' not in standard
    assert 'sidebar_navigation(' not in comparison
    assert 'mobile_navigation(' not in comparison
    assert '"Загрузить другой отчёт"' in standard
    assert '"Загрузить другие периоды"' in comparison


def test_user_guide_is_detailed_and_shared_with_contextual_help():
    guide = (ROOT / "USER_GUIDE.md").read_text(encoding="utf-8")
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    assert len(guide.splitlines()) > 800
    for heading in (
        "## 1. Общий анализ продаж",
        "## 2. Сравнение периодов",
        "## 4. Заказ поставщику",
        "## 5. Заказ Sonu",
        "## 6. Склад Baserow",
        "## 7. Управленческий отчет",
    ):
        assert heading in guide
    assert 'sections = guide_sections(Path(__file__).with_name("USER_GUIDE.md"))' in app
    assert 'on_click=_open_user_guide' in app


def test_visible_report_chrome_no_longer_repeats_filter_and_rate():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    sonu = (ROOT / "src" / "sonu.py").read_text(encoding="utf-8")
    warehouse = (ROOT / "src" / "warehouse.py").read_text(encoding="utf-8")
    assert 'Фильтр металла применен ко всему обычному отчету' not in app
    assert 'Глобальный фильтр применен ко всей странице' not in app
    assert 'Фильтр материала применен к остаткам' not in warehouse
    assert '· Курс: 1 USD =' not in sonu
