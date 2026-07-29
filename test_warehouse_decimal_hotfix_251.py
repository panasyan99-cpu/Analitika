from types import SimpleNamespace

import pytest

from src.warehouse_management.client import WarehouseClient, WarehouseClientError


def _client() -> WarehouseClient:
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
    client = WarehouseClient(config)
    client._fields[5] = [
        {"name": "Закупка USD/ед.", "type": "number", "number_decimal_places": 6},
        {"name": "Продажа VND при импорте", "type": "number", "number_decimal_places": 0},
    ]
    return client


def test_baserow_number_payload_is_rounded_to_field_precision() -> None:
    clean = _client().normalize_payload(
        5,
        {
            "Закупка USD/ед.": 0.8314798076428942,
            "Продажа VND при импорте": 220339.9,
        },
    )
    assert clean["Закупка USD/ед."] == pytest.approx(0.831480)
    assert clean["Продажа VND при импорте"] == 220340
    assert len(str(clean["Закупка USD/ед."]).split(".")[-1]) <= 6


def test_baserow_number_payload_rejects_non_finite_values() -> None:
    with pytest.raises(WarehouseClientError, match="недопустимое число"):
        _client().normalize_payload(5, {"Закупка USD/ед.": float("nan")})
