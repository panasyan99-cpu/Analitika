from dataclasses import dataclass
from pathlib import Path

from src.warehouse_management.models import SupplySummary
from src.warehouse_management.schema import SUPPLY_LINES_TABLE_NAME
from src.warehouse_management.service import WarehouseService


@dataclass
class MixedConfig:
    souvenirs_table_id: int = 642
    components_table_id: int = 643
    operations_table_id: int = 644
    supplies_table_id: int = 645
    supply_lines_table_id: int = 646


class MixedClient:
    def __init__(self) -> None:
        self.config = MixedConfig()
        self.created_operations = []
        self.rows = {
            642: [
                {"id": 1, "Артикул": "SOUV-1", "Остаток": 9, "Активный SKU": True},
                {"id": 2, "Артикул": "ARCHIVE", "Остаток": 4, "Активный SKU": False},
            ],
            # Row IDs can overlap across different Baserow tables.
            643: [{"id": 1, "Артикул": "COMP-1", "Остаток": 5, "Активно": True}],
            644: [],
            645: [{"id": 10, "№ поставки": "SUP-MIX", "Статус": {"value": "Получена полностью"}}],
            646: [
                {
                    "id": 501,
                    "Строка поставки": "SUP-MIX — SOUV-1",
                    "Поставка": [{"id": 10}],
                    "Товар сувенирки": [{"id": 1}],
                    "По документу, шт.": 4,
                    "Принято, шт.": 4,
                    "Передано в бухгалтерию, шт.": 0,
                },
                {
                    "id": 502,
                    "Строка поставки": "SUP-MIX — COMP-1",
                    "Поставка": [{"id": 10}],
                    "Комплектующее": [{"id": 1}],
                    "По документу, шт.": 3,
                    "Принято, шт.": 3,
                    "Передано в бухгалтерию, шт.": 0,
                },
            ],
        }

    def list_rows(self, table_id):
        return [dict(row) for row in self.rows.get(table_id, [])]

    def batch_id(self, prefix):
        return f"{prefix}-TEST"

    def create_operations(self, items, *, batch_id, command_id=""):
        self.created_operations = [{"id": 900 + index, **item} for index, item in enumerate(items)]
        return self.created_operations

    def mark_operations_status(self, rows, status):
        return None

    def batch_update(self, table_id, items):
        return None


def _supply() -> SupplySummary:
    return SupplySummary(
        row_id=10,
        supply_id="SUP-MIX",
        date="2026-07-28",
        supplier="",
        status="Получена полностью",
        sku_total=2,
        sku_received=2,
        qty_document=7,
        qty_received=7,
        qty_waiting=0,
        raw={},
    )


def test_mixed_supply_uses_supply_line_ids_and_correct_link_fields() -> None:
    client = MixedClient()
    service = WarehouseService(client)
    result = service.transfer_supply(_supply(), {501: 4, 502: 3})
    assert result["quantity"] == 7
    by_sku = {row["Операция"].split(" — ", 1)[1]: row for row in client.created_operations}
    assert by_sku["SOUV-1"]["Раздел"] == "Сувенирка"
    assert by_sku["SOUV-1"]["Товар сувенирки"] == [1]
    assert by_sku["COMP-1"]["Раздел"] == "Комплектующие"
    assert by_sku["COMP-1"]["Комплектующее"] == [1]
    assert by_sku["SOUV-1"]["Позиция поставки"] == [501]
    assert by_sku["COMP-1"]["Позиция поставки"] == [502]


def test_archived_cards_are_hidden_by_default() -> None:
    service = WarehouseService(MixedClient())
    assert [item.sku for item in service.catalog("Сувенирка")] == ["SOUV-1"]
    assert {item.sku for item in service.catalog("Сувенирка", include_inactive=True)} == {
        "SOUV-1",
        "ARCHIVE",
    }


def test_release_contains_safe_schema_and_strict_excel_validation() -> None:
    schema = Path("src/warehouse_management/schema.py").read_text(encoding="utf-8")
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    assert SUPPLY_LINES_TABLE_NAME == "Позиции поставок"
    assert '"first_row_header": True' in schema
    assert '"Позиция поставки"' in schema
    assert "В Excel отсутствуют обязательные колонки" in ui
    assert '_auto_prepare_safe_schema' in ui
    assert 'HISTORY_WORKSPACES = ("Операции",)' in ui
    assert "st.cache_data.clear()" not in ui


def test_cumulative_release_history_is_documented() -> None:
    changelog = Path("CHANGELOG.md").read_text(encoding="utf-8")
    cumulative = Path("CUMULATIVE_UPDATES_2.1_TO_2.4.md").read_text(encoding="utf-8")
    for version in ("2.1.0", "2.2.0", "2.3.0", "2.4.0", "2.4.1", "2.4.2", "2.4.3", "2.4.4"):
        assert version in cumulative
        assert version in changelog
