from pathlib import Path

ROOT = Path(__file__).parent


def test_fifth_workspace_is_registered_everywhere():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    product = (ROOT / "src" / "product_info.py").read_text(encoding="utf-8")
    module = (ROOT / "src" / "order_workflow.py").read_text(encoding="utf-8")
    assert '"Заказ поставщику": {' in app
    assert 'elif mode == "Заказ поставщику"' in app
    assert "render_supplier_order_dashboard()" in app
    assert 'if mode != "Заказ поставщику"' in app
    assert 'mode="Заказ поставщику"' in product
    assert "ORDER_MODE_STONES" in module
    assert "ORDER_MODE_PEARLS" in module


def test_order_module_contains_required_business_rules():
    source = (ROOT / "src" / "order_workflow.py").read_text(encoding="utf-8")
    assert "working_raw = total - stock_63 - stock_20" in source
    assert "ntr2" in source.lower()
    assert "Ошибка ТВП" in source
    assert "RING_SIZES = tuple(range(15, 25))" in source
    assert "С остатком сверился" in source
    assert "Скачать заказ в Excel" in source
    assert "Фото", "Артикул"  # readability marker for exact export contract below
    assert 'headers = ["Фото", "Артикул", "Камень", "Группа", "Количество к заказу", "Размеры"]' in source


def test_version_is_180():
    version = (ROOT / "version.json").read_text(encoding="utf-8")
    assert '"version": "1.8.0"' in version
