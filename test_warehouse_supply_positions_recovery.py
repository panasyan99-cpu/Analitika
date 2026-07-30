from dataclasses import dataclass
from types import SimpleNamespace

from src.warehouse_management.schema import BaserowSchemaManager
from src.warehouse_management import ui


@dataclass(frozen=True)
class Config:
    base_url: str = "https://example.invalid"
    token: str = "token"
    database_id: int = 148
    souvenirs_table_id: int = 642
    components_table_id: int = 643
    operations_table_id: int = 644
    supplies_table_id: int = 645
    supply_lines_table_id: int = 0
    email: str = ""
    password: str = ""


def test_resolved_config_probes_known_table_when_metadata_listing_is_restricted(monkeypatch):
    class FakeClient:
        def __init__(self, config):
            self.config = config

        def table_is_accessible(self, table_id):
            return int(table_id) == 646

        def discover_table_id(self, name):
            raise AssertionError("metadata discovery must not be needed for known table 646")

    monkeypatch.setattr(ui, "WarehouseClient", FakeClient)
    monkeypatch.setattr(ui, "st", SimpleNamespace(session_state={}))

    resolved = ui._resolved_config(Config())

    assert resolved.supply_lines_table_id == 646
    assert ui.st.session_state["warehouse_supply_lines_table_id"] == 646


def test_migration_restores_missing_component_and_silver_supply_lines():
    manager = object.__new__(BaserowSchemaManager)
    rows = {
        646: [],
        642: [],
        643: [
            {
                "id": 20,
                "Артикул": "SIL000020",
                "Поставки": [{"id": 11}],
                "По документу, шт.": 25,
                "Получено по поставке, шт.": 0,
                "Номера коробок": "S-01",
                "Серебро 925": True,
            }
        ],
        644: [
            {
                "id": 90,
                "Тип операции": {"value": "Приход"},
                "Количество": 2,
                "Поставка": [{"id": 11}],
                "Комплектующее": [{"id": 20}],
            }
        ],
        645: [{"id": 11, "№ поставки": "SIL-20260730-001"}],
    }
    created = []
    manager.list_rows = lambda table_id: [dict(row) for row in rows.get(int(table_id), [])]
    manager.create_rows = lambda table_id, items: created.extend(list(items)) or created

    migrated, skipped, ambiguous = manager.migrate_legacy_supply_lines(
        table_id=646,
        souvenirs_table_id=642,
        components_table_id=643,
        operations_table_id=644,
        supplies_table_id=645,
    )

    assert migrated == 1
    assert skipped == 0
    assert ambiguous == []
    assert created[0]["Поставка"] == [11]
    assert created[0]["Комплектующее"] == [20]
    assert created[0]["По документу, шт."] == 25
    assert created[0]["Принято, шт."] == 2
    assert created[0]["Статус"] == "Частично получена"
