from types import SimpleNamespace

from src.warehouse_management.client import WarehouseClient
from src.warehouse_management.models import Product
from src.warehouse_management.service import WarehouseService


def _service() -> WarehouseService:
    config = SimpleNamespace(
        base_url="https://example.invalid",
        token="test",
        souvenirs_table_id=1,
        components_table_id=2,
        operations_table_id=3,
        supplies_table_id=4,
        supply_lines_table_id=5,
        email="",
        password="",
    )
    return WarehouseService(WarehouseClient(config))


def _product(*, silver: bool) -> Product:
    return Product(
        number=1,
        boxes="",
        sku="SIL000001" if silver else "ACC001",
        qty_document=10,
        description="Test item",
        category="Аксессуары",
        material="Silver" if silver else "Brass",
        stone="",
        color="",
        unit_weight_kg=None,
        image_path="",
        received=False,
        actual_manual=None,
        comment="",
        checked=True,
        recognition="",
        name="Silver test" if silver else "Accessory test",
        silver_category="Бусины" if silver else "",
        silver_925=silver,
    )


def test_silver_payload_does_not_send_unsupported_legacy_category() -> None:
    service = _service()
    payload = service._product_payload(
        _product(silver=True),
        section="Комплектующие",
        supply_row_id=10,
        existing=None,
        photo=None,
    )
    assert "Категория" not in payload
    assert payload["Серебряная категория"] == "Бусины"
    assert payload["Серебро 925"] is True


def test_regular_component_payload_keeps_legacy_category() -> None:
    service = _service()
    payload = service._product_payload(
        _product(silver=False),
        section="Комплектующие",
        supply_row_id=10,
        existing=None,
        photo=None,
    )
    assert payload["Категория"] == "Аксессуары"
    assert "Серебряная категория" not in payload
