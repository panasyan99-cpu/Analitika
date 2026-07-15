from pathlib import Path

ROOT = Path(__file__).parent


def test_warehouse_module_is_restored():
    warehouse = (ROOT / "src" / "warehouse.py").read_text(encoding="utf-8")
    assert "class WarehouseConfig" in warehouse
    assert "def load_bundle" in warehouse
    assert "def render_warehouse_dashboard" in warehouse
    assert "Сувениры" in warehouse
    assert "Касты" in warehouse
    assert "Требует внимания" in warehouse
    assert "Движение" in warehouse
    assert "Поставки" in warehouse
    assert "def render_warehouse_dashboard():\n    pass" not in warehouse


def test_sonu_has_six_branded_navigation_blocks_without_blue_state():
    sonu = (ROOT / "src" / "sonu.py").read_text(encoding="utf-8")
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    for label in [
        "Продажи по магазинам",
        "Средние продажи",
        "Браслеты",
        "Прогноз продаж",
        "Матрица заказа",
        "Рекомендации",
    ]:
        assert label in sonu
    assert 'key="sonu_section"' in sonu
    assert '[aria-pressed="true"]' in app
    assert '#f2cf8c' in app
    assert 'background: linear-gradient(135deg, #17130e 0%, #332515 100%)' in app


def test_about_modes_come_from_feature_registry():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    info = (ROOT / "src" / "product_info.py").read_text(encoding="utf-8")
    assert "REPORT_MODES" in app
    assert "PRODUCT_FEATURES" in info
    assert "release_history_html" in info
