from dataclasses import dataclass

from src.warehouse_management.service import WarehouseService


@dataclass
class FakeConfig:
    souvenirs_table_id: int = 642
    components_table_id: int = 643
    operations_table_id: int = 644
    supplies_table_id: int = 645
    supply_lines_table_id: int = 646


class FakeClient:
    def __init__(self):
        self.config = FakeConfig()
        self.rows = {
            642: [],
            643: [
                {
                    "id": 20,
                    "Артикул": "SIL-NEW",
                    "Поставки": [{"id": 11}],
                    "По документу, шт.": 25,
                    "Получено по поставке, шт.": 0,
                    "Активно": True,
                }
            ],
            644: [],
            645: [
                {"id": 10, "№ поставки": "NEW-LINES", "Дата": "2026-07-29", "Статус": "Ожидается"},
                {"id": 11, "№ поставки": "LEGACY-COMP", "Дата": "2026-07-30", "Статус": "Ожидается"},
                {"id": 12, "№ поставки": "HEADER-ONLY", "Дата": "2026-07-30", "Статус": "Ожидается", "Ожидается": 7},
            ],
            646: [
                {
                    "id": 500,
                    "Поставка": [{"id": 10}],
                    "Комплектующее": [{"id": 20}],
                    "По документу, шт.": 10,
                    "Принято, шт.": 0,
                    "Статус": "Ожидается",
                }
            ],
        }

    def list_rows(self, table_id, **kwargs):
        return [dict(row) for row in self.rows.get(table_id, [])]


def test_supply_registry_keeps_new_lines_legacy_components_and_header_only_rows_visible():
    service = WarehouseService(FakeClient())
    summaries = {item.supply_id: item for item in service.supply_summaries()}

    assert set(summaries) == {"NEW-LINES", "LEGACY-COMP", "HEADER-ONLY"}
    assert summaries["NEW-LINES"].qty_waiting == 10
    assert summaries["LEGACY-COMP"].qty_waiting == 25
    assert summaries["HEADER-ONLY"].qty_waiting == 7


def test_supply_detail_falls_back_to_legacy_component_links_when_no_line_exists():
    service = WarehouseService(FakeClient())
    rows = service.supply_products(11)

    assert len(rows) == 1
    assert rows[0]["Артикул"] == "SIL-NEW"
    assert rows[0]["_section"] == "Комплектующие"
    assert rows[0]["_document"] == 25
