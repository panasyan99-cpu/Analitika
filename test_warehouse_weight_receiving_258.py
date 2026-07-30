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
    supply_lines_table_id: int = 646


class FakeClient:
    def __init__(self, *, received: int = 0) -> None:
        self.config = FakeConfig()
        self.rows = {
            642: [],
            643: [
                {
                    "id": 2,
                    "Артикул": "SIL000001",
                    "Название": "Тестовая бусина",
                    "Остаток": 0,
                    "Поставки": [{"id": 10}],
                    "Активно": True,
                    "Фото": [],
                }
            ],
            644: [],
            645: [{"id": 10, "№ поставки": "SIL-TEST", "Статус": "Ожидается"}],
            646: [
                {
                    "id": 500,
                    "Поставка": [{"id": 10}],
                    "Комплектующее": [{"id": 2}],
                    "По документу, шт.": 100,
                    "Принято, шт.": received,
                    "Передано в бухгалтерию, шт.": 0,
                    "Статус": "Ожидается" if not received else "Частично получена",
                    "Единица учёта": "шт.",
                    "Вес партии, г": 50.0,
                    "Вес единицы, г": 0.5,
                    "Активна": True,
                }
            ],
        }
        self._next_operation_id = 900

    def list_rows(self, table_id: int, *, refresh: bool = False):
        return [dict(row) for row in self.rows.get(table_id, [])]

    def batch_update(self, table_id: int, payloads):
        for payload in payloads:
            target = next(row for row in self.rows[table_id] if int(row["id"]) == int(payload["id"]))
            target.update(payload)

    def batch_id(self, prefix: str) -> str:
        return f"{prefix}-TEST"

    def create_operations(self, items, *, batch_id: str, command_id: str = ""):
        created = []
        for item in items:
            row = {
                "id": self._next_operation_id,
                **item,
                "Статус документа": "Создаётся",
            }
            self._next_operation_id += 1
            self.rows[644].append(row)
            created.append(dict(row))
        return created

    def mark_operations_status(self, rows, status: str):
        ids = {int(row["id"]) for row in rows}
        for row in self.rows[644]:
            if int(row["id"]) in ids:
                row["Статус документа"] = status


def supply() -> SupplySummary:
    return SupplySummary(
        row_id=10,
        supply_id="SIL-TEST",
        date="2026-07-30",
        supplier="",
        status="Ожидается",
        sku_total=1,
        sku_received=0,
        qty_document=100,
        qty_received=0,
        qty_waiting=100,
        raw={},
    )


def test_weight_quantity_uses_half_up_rounding_and_document_cap():
    assert WarehouseService.estimate_quantity_from_weight(1.25, 0.5) == 3
    assert WarehouseService.estimate_quantity_from_weight(100, 0.5, maximum=100) == 100
    assert WarehouseService.estimate_quantity_from_weight(0, 0.5, maximum=100) == 0


def test_weight_receiving_posts_full_receipt_and_pre_accounting_expense():
    client = FakeClient()
    service = WarehouseService(client)

    result = service.receive_existing_supply_by_weight(
        supply(),
        {500: {"weight_g": 20.0, "quantity": 40}},
        command_id="CMD-RECW-TEST",
    )

    assert result == {
        "batch_id": "RECW-TEST",
        "command_id": "CMD-RECW-TEST",
        "sku": 1,
        "received": 100,
        "current": 40,
        "written_off": 60,
    }
    operations = client.rows[644]
    assert [row["Тип операции"] for row in operations] == ["Приход", "Расход"]
    assert [row["Количество"] for row in operations] == [100, 60]
    assert all(row["Статус документа"] == "Проведена" for row in operations)

    line = client.rows[646][0]
    assert line["Принято, шт."] == 100
    assert line["Статус"] == "Получена полностью"
    assert line["Способ приёмки"] == "По весу — товар уже в работе"
    assert line["Вес при приёмке, г"] == 20.0
    assert line["Расчётное количество по весу"] == 40
    assert line["Погрешность веса, г"] == 0.0
    assert client.rows[645][0]["Статус"] == "Получена полностью"


def test_weight_receiving_is_blocked_after_any_previous_receipt():
    client = FakeClient(received=1)
    service = WarehouseService(client)

    with pytest.raises(WarehouseServiceError, match="только до первой приёмки"):
        service.receive_existing_supply_by_weight(
            supply(),
            {500: {"weight_g": 20.0, "quantity": 40}},
        )
