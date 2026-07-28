from dataclasses import dataclass

import pytest

from src.warehouse_management.models import SupplySummary
from src.warehouse_management.service import WarehouseService, WarehouseServiceError


@dataclass
class FakeConfig:
    souvenirs_table_id: int = 642
    components_table_id: int = 643
    operations_table_id: int = 644
    supplies_table_id: int = 645
    supply_lines_table_id: int = 0


class FakeClient:
    def __init__(self):
        self.config = FakeConfig()
        self.created_operations = []
        self.rows = {
            642: [
                {
                    "id": 1,
                    "Артикул": "SKU-A",
                    "Остаток": 20,
                    "Поставки": [{"id": 10, "value": "SUP-1"}],
                    "По документу, шт.": 12,
                    "Получено по поставке, шт.": 12,
                }
            ],
            643: [],
            644: [
                {
                    "id": 100,
                    "Тип операции": {"value": "Передача в бухгалтерию"},
                    "Товар сувенирки": [{"id": 1, "value": "SKU-A"}],
                    "Количество": 5,
                    "ID поставки": "SUP-1",
                    "Batch ID": "ACC-OLD",
                }
            ],
            645: [
                {
                    "id": 10,
                    "№ поставки": "SUP-1",
                    "Статус": {"value": "Получена полностью"},
                }
            ],
        }

    def list_rows(self, table_id):
        return [dict(row) for row in self.rows.get(table_id, [])]

    def batch_id(self, prefix):
        return f"{prefix}-TEST"

    def create_operations(self, items, *, batch_id):
        self.created_operations.extend(items)
        return items

    def batch_update(self, table_id, items):
        return None


def supply():
    return SupplySummary(
        row_id=10,
        supply_id="SUP-1",
        date="2026-07-17",
        supplier="",
        status="Получена полностью",
        sku_total=1,
        sku_received=1,
        qty_document=12,
        qty_received=12,
        qty_waiting=0,
        raw={},
    )


def test_transfer_uses_quantity_available_from_selected_supply():
    client = FakeClient()
    service = WarehouseService(client)
    report = service.transfer_supply(supply(), {1: 7})
    assert report["quantity"] == 7
    assert client.created_operations[0]["Количество"] == 7
    assert client.created_operations[0]["ID поставки"] == "SUP-1"


def test_transfer_blocks_more_than_received_minus_already_transferred():
    client = FakeClient()
    service = WarehouseService(client)
    with pytest.raises(WarehouseServiceError, match="не более 7"):
        service.transfer_supply(supply(), {1: 8})


def test_incoming_correction_creates_outgoing_reverse_operation():
    client = FakeClient()
    service = WarehouseService(client)
    operation = {
        "id": 200,
        "Операция": "REC-1 — SKU-A",
        "Тип операции": {"value": "Приход"},
        "Раздел": {"value": "Сувенирка"},
        "Товар сувенирки": [{"id": 1, "value": "SKU-A"}],
        "Количество": 4,
        "Batch ID": "REC-1",
    }
    service.correct_operation(operation, quantity=4, comment="Ошибка")
    created = client.created_operations[0]
    assert created["Тип операции"] == "Расход"
    assert created["Количество"] == 4


def test_outgoing_correction_creates_return():
    client = FakeClient()
    service = WarehouseService(client)
    operation = {
        "id": 201,
        "Операция": "ACC-1 — SKU-A",
        "Тип операции": {"value": "Передача в бухгалтерию"},
        "Раздел": {"value": "Сувенирка"},
        "Товар сувенирки": [{"id": 1, "value": "SKU-A"}],
        "Количество": 3,
        "Batch ID": "ACC-1",
    }
    service.correct_operation(operation, quantity=2, comment="Возврат")
    created = client.created_operations[0]
    assert created["Тип операции"] == "Возврат"
    assert created["Количество"] == 2
