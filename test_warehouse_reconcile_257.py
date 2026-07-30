from dataclasses import dataclass

from src.warehouse_management.models import SupplySummary
from src.warehouse_management.service import WarehouseService


@dataclass
class FakeConfig:
    souvenirs_table_id: int = 642
    components_table_id: int = 643
    operations_table_id: int = 644
    supplies_table_id: int = 645
    supply_lines_table_id: int = 646


class FakeClient:
    def __init__(self) -> None:
        self.config = FakeConfig()
        self.rows = {
            642: [],
            643: [
                {
                    "id": 2,
                    "Артикул": "OLD-SKU",
                    "Остаток": 0,
                    "Поставки": [{"id": 10}],
                    "Активно": True,
                },
                {
                    "id": 3,
                    "Артикул": "HIST-SKU",
                    "Остаток": 0,
                    "Поставки": [{"id": 10}],
                    "Активно": True,
                },
                {
                    "id": 4,
                    "Артикул": "LIVE-SKU",
                    "Остаток": 5,
                    "Поставки": [{"id": 99}],
                    "Активно": True,
                },
            ],
            644: [
                {
                    "id": 800,
                    "Тип операции": "Приход",
                    "Статус документа": "Проведена",
                    "Комплектующее": [{"id": 3}],
                    "Количество": 1,
                    "Поставка": [{"id": 99}],
                },
                {
                    "id": 801,
                    "Тип операции": "Приход",
                    "Статус документа": "Проведена",
                    "Комплектующее": [{"id": 4}],
                    "Позиция поставки": [{"id": 501}],
                    "Поставка": [{"id": 10}],
                    "Количество": 7,
                },
                {
                    "id": 802,
                    "Тип операции": "Расход",
                    "Статус документа": "Проведена",
                    "Комплектующее": [{"id": 4}],
                    "Позиция поставки": [{"id": 501}],
                    "Поставка": [{"id": 10}],
                    "Количество": 1,
                },
                {
                    "id": 803,
                    "Тип операции": "Передача в бухгалтерию",
                    "Статус документа": "Проведена",
                    "Комплектующее": [{"id": 4}],
                    "Позиция поставки": [{"id": 501}],
                    "Поставка": [{"id": 10}],
                    "Количество": 2,
                },
                {
                    "id": 804,
                    "Тип операции": "Возврат",
                    "Статус документа": "Проведена",
                    "Комплектующее": [{"id": 4}],
                    "Позиция поставки": [{"id": 501}],
                    "Поставка": [{"id": 10}],
                    "Количество": 1,
                },
                {
                    "id": 805,
                    "Тип операции": "Приход",
                    "Статус документа": "Создаётся",
                    "Комплектующее": [{"id": 4}],
                    "Позиция поставки": [{"id": 501}],
                    "Поставка": [{"id": 10}],
                    "Количество": 100,
                },
            ],
            645: [
                {"id": 10, "№ поставки": "SUP-10", "Статус": "Ожидается"},
            ],
            646: [
                {
                    "id": 500,
                    "Поставка": [{"id": 10}],
                    "Комплектующее": [{"id": 2}],
                    "По документу, шт.": 5,
                    "Принято, шт.": 0,
                    "Передано в бухгалтерию, шт.": 0,
                    "Статус": "Ожидается",
                    "Активна": True,
                },
                {
                    "id": 501,
                    "Поставка": [{"id": 10}],
                    "Комплектующее": [{"id": 4}],
                    "По документу, шт.": 10,
                    "Принято, шт.": 0,
                    "Передано в бухгалтерию, шт.": 0,
                    "Статус": "Ожидается",
                    "Активна": True,
                },
            ],
        }

    def list_rows(self, table_id: int, *, refresh: bool = False):
        return [dict(row) for row in self.rows.get(table_id, [])]

    def batch_update(self, table_id: int, payloads):
        for payload in payloads:
            target = next(row for row in self.rows[table_id] if int(row["id"]) == int(payload["id"]))
            target.update(payload)

    def delete_row(self, table_id: int, row_id: int):
        self.rows[table_id] = [row for row in self.rows[table_id] if int(row["id"]) != int(row_id)]


def _supply() -> SupplySummary:
    return SupplySummary(
        row_id=10,
        supply_id="SUP-10",
        date="2026-07-30",
        supplier="",
        status="Ожидается",
        sku_total=2,
        sku_received=0,
        qty_document=15,
        qty_received=0,
        qty_waiting=15,
        raw={},
    )


def test_removing_never_received_line_deletes_zero_stock_orphan_catalog_row():
    client = FakeClient()
    service = WarehouseService(client)

    removed = service.remove_waiting_from_supply(_supply(), [2])

    assert removed == 1
    assert all(int(row["id"]) != 500 for row in client.rows[646])
    assert all(int(row["id"]) != 2 for row in client.rows[643])
    assert any(int(row["id"]) == 4 for row in client.rows[643])


def test_baserow_reconcile_repairs_receipts_status_links_and_old_orphans():
    client = FakeClient()
    # Simulate the old bug: line 500 was already deleted, but OLD-SKU remained linked.
    client.rows[646] = [row for row in client.rows[646] if int(row["id"]) != 500]
    service = WarehouseService(client)

    report = service.synchronize_baserow_from_documents()

    line = next(row for row in client.rows[646] if int(row["id"]) == 501)
    assert line["Принято, шт."] == 6
    assert line["Передано в бухгалтерию, шт."] == 1
    assert line["Статус"] == "Частично передана"
    assert client.rows[645][0]["Статус"] == "Частично получена"

    assert all(int(row["id"]) != 2 for row in client.rows[643])
    historical = next(row for row in client.rows[643] if int(row["id"]) == 3)
    assert historical["Поставки"] == []
    assert historical["Активно"] is False
    live = next(row for row in client.rows[643] if int(row["id"]) == 4)
    assert live["Поставки"] == [10]

    assert report == {
        "lines_updated": 1,
        "supplies_updated": 1,
        "catalog_relinked": 2,
        "catalog_deleted": 1,
        "catalog_deactivated": 1,
    }
