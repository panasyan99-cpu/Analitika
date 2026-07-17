from pathlib import Path


ROOT = Path(__file__).parent


def test_header_changes_for_every_workspace():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    assert "HERO_CONTENT = {" in app
    assert '"Обычный отчет": {' in app
    assert '"Сравнение периодов": {' in app
    assert '"Сувениры и касты на складе": {' in app
    assert '"Заказ Sonu": {' in app
    assert 'def render_hero(mode: str)' in app
    assert 'render_hero(active_mode)' in app
    assert 'Данные, которые' not in app
    assert 'помогают решать' not in app


def test_duplicate_module_heroes_are_removed():
    sonu = (ROOT / "src" / "sonu.py").read_text(encoding="utf-8")
    warehouse = (ROOT / "src" / "warehouse.py").read_text(encoding="utf-8")
    assert 'SONU · ORDER ANALYTICS' not in sonu
    assert '<section class="warehouse-header">' not in warehouse


def test_streamlit_status_and_toolbar_chrome_are_hidden():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    config = (ROOT / ".streamlit" / "config.toml").read_text(encoding="utf-8")
    assert '[data-testid="stStatusWidget"]' in app
    assert '[data-testid="stToolbar"]' in app
    assert 'display:none !important' in app
    assert 'toolbarMode = "minimal"' in config
