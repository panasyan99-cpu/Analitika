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


def test_sonu_has_final_navigation_blocks_and_shared_control_style():
    sonu = (ROOT / "src" / "sonu.py").read_text(encoding="utf-8")
    for label in [
        "Основной отчет", "Общие выводы", "Серьги", "Кольца", "Подвески",
        "Браслеты не полный круг", "Браслеты полный круг", "Полная выгрузка",
    ]:
        assert label in sonu

def test_about_modes_come_from_feature_registry():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    info = (ROOT / "src" / "product_info.py").read_text(encoding="utf-8")
    assert "REPORT_MODES" in app
    assert "PRODUCT_FEATURES" in info
    assert "release_history_html" in info
