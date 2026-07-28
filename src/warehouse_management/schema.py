from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable

import requests

from .client import as_int, link_ids, select_text


SUPPLY_LINES_TABLE_NAME = "Позиции поставок"
SUPPLY_LINE_STATUSES = (
    "Ожидается",
    "Частично получена",
    "Получена полностью",
    "Частично передана",
    "Передана полностью",
    "Требует проверки",
    "Отменена",
)


class WarehouseSchemaError(RuntimeError):
    pass


@dataclass(frozen=True)
class SchemaMigrationReport:
    table_id: int
    created_table: bool
    created_fields: tuple[str, ...]
    migrated_lines: int
    skipped_lines: int
    ambiguous_skus: tuple[str, ...]
    added_operation_fields: tuple[str, ...]
    added_supply_fields: tuple[str, ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "table_id": self.table_id,
            "created_table": self.created_table,
            "created_fields": list(self.created_fields),
            "migrated_lines": self.migrated_lines,
            "skipped_lines": self.skipped_lines,
            "ambiguous_skus": list(self.ambiguous_skus),
            "added_operation_fields": list(self.added_operation_fields),
            "added_supply_fields": list(self.added_supply_fields),
        }


class BaserowSchemaManager:
    """Baserow schema/migration client authenticated with a short-lived JWT.

    Database tokens deliberately cannot change table structure. This manager is
    only used from the protected maintenance page or the bootstrap CLI and does
    not persist the user's Baserow password.
    """

    def __init__(self, base_url: str, email: str, password: str, *, timeout: int = 60) -> None:
        self.base_url = str(base_url).rstrip("/")
        self.email = str(email).strip()
        self.password = str(password)
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Accept": "application/json",
                "User-Agent": "Princess-Analitika-Warehouse-Schema/2.4.1",
            }
        )
        self._authenticate()

    def _authenticate(self) -> None:
        if not self.email or not self.password:
            raise WarehouseSchemaError("Нужны email и пароль пользователя Baserow с правами Builder/Admin.")
        try:
            response = self.session.post(
                f"{self.base_url}/api/user/token-auth/",
                json={"email": self.email, "password": self.password},
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            raise WarehouseSchemaError(f"Не удалось подключиться к Baserow: {exc}") from exc
        if not response.ok:
            raise WarehouseSchemaError(
                f"Baserow не выдал JWT ({response.status_code}). Проверьте email, пароль и права."
            )
        payload = response.json()
        token = str(payload.get("access_token") or payload.get("token") or "").strip()
        if not token:
            raise WarehouseSchemaError("Baserow не вернул JWT-токен.")
        self.session.headers["Authorization"] = f"JWT {token}"

    def request(
        self,
        method: str,
        path: str,
        *,
        payload: Any = None,
        params: dict[str, Any] | None = None,
    ) -> Any:
        try:
            response = self.session.request(
                method,
                f"{self.base_url}{path}",
                json=payload,
                params=params,
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            raise WarehouseSchemaError(f"Ошибка соединения с Baserow: {exc}") from exc
        if not response.ok:
            raise WarehouseSchemaError(
                f"Baserow HTTP {response.status_code}: {response.text[:1000]}"
            )
        if not response.content:
            return None
        return response.json()

    def list_tables(self, database_id: int) -> list[dict[str, Any]]:
        result = self.request("GET", f"/api/database/tables/database/{int(database_id)}/")
        return list(result or [])

    def fields(self, table_id: int) -> list[dict[str, Any]]:
        result = self.request("GET", f"/api/database/fields/table/{int(table_id)}/")
        return list(result or [])

    def list_rows(self, table_id: int) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        page = 1
        while True:
            payload = self.request(
                "GET",
                f"/api/database/rows/table/{int(table_id)}/",
                params={"user_field_names": "true", "size": 200, "page": page},
            )
            result.extend(payload.get("results", []))
            if not payload.get("next"):
                return result
            page += 1

    def create_table(self, database_id: int, name: str) -> dict[str, Any]:
        return self.request(
            "POST",
            f"/api/database/tables/database/{int(database_id)}/",
            payload={
                "name": name,
                "data": [["Строка поставки"]],
                "first_row_header": True,
            },
        )

    def create_field(self, table_id: int, payload: dict[str, Any]) -> dict[str, Any]:
        return self.request(
            "POST",
            f"/api/database/fields/table/{int(table_id)}/",
            payload=payload,
        )

    def update_field(self, field_id: int, payload: dict[str, Any]) -> dict[str, Any]:
        return self.request(
            "PATCH",
            f"/api/database/fields/{int(field_id)}/",
            payload=payload,
        )

    def create_rows(self, table_id: int, items: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
        source = list(items)
        created: list[dict[str, Any]] = []
        for start in range(0, len(source), 100):
            response = self.request(
                "POST",
                f"/api/database/rows/table/{int(table_id)}/batch/",
                params={"user_field_names": "true"},
                payload={"items": source[start : start + 100]},
            )
            if isinstance(response, dict):
                created.extend(response.get("items", []))
            elif isinstance(response, list):
                created.extend(response)
        return created

    def _ensure_field(
        self,
        table_id: int,
        name: str,
        field_type: str,
        **extra: Any,
    ) -> bool:
        current = {str(field.get("name") or ""): field for field in self.fields(table_id)}
        if name in current:
            return False
        self.create_field(table_id, {"name": name, "type": field_type, **extra})
        return True

    def ensure_supply_lines_table(
        self,
        *,
        database_id: int,
        souvenirs_table_id: int,
        components_table_id: int,
        operations_table_id: int,
        supplies_table_id: int,
    ) -> tuple[int, bool, list[str], list[str], list[str]]:
        tables = self.list_tables(database_id)
        existing = next(
            (table for table in tables if str(table.get("name") or "").strip().casefold() == SUPPLY_LINES_TABLE_NAME.casefold()),
            None,
        )
        created_table = existing is None
        table = self.create_table(database_id, SUPPLY_LINES_TABLE_NAME) if existing is None else existing
        table_id = int(table["id"])

        fields = self.fields(table_id)
        if fields:
            primary = next((field for field in fields if field.get("primary")), fields[0])
            if str(primary.get("name") or "") != "Строка поставки":
                self.update_field(int(primary["id"]), {"name": "Строка поставки"})

        created_fields: list[str] = []

        def ensure(name: str, kind: str, **extra: Any) -> None:
            if self._ensure_field(table_id, name, kind, **extra):
                created_fields.append(name)

        ensure("Поставка", "link_row", link_row_table_id=int(supplies_table_id))
        ensure("Товар сувенирки", "link_row", link_row_table_id=int(souvenirs_table_id))
        ensure("Комплектующее", "link_row", link_row_table_id=int(components_table_id))
        ensure("По документу, шт.", "number", number_decimal_places=0, number_negative=False)
        ensure("Принято, шт.", "number", number_decimal_places=0, number_negative=False)
        ensure("Передано в бухгалтерию, шт.", "number", number_decimal_places=0, number_negative=False)
        ensure("Номера коробок", "long_text")
        ensure(
            "Статус",
            "single_select",
            select_options=[
                {"value": value, "color": color}
                for value, color in zip(
                    SUPPLY_LINE_STATUSES,
                    ("light-gray", "light-orange", "light-green", "light-blue", "dark-green", "light-red", "dark-gray"),
                )
            ],
        )
        ensure("Комментарий", "long_text")
        ensure("Версия", "number", number_decimal_places=0, number_negative=False)
        ensure("Активна", "boolean")
        ensure("Command ID", "text")
        ensure("Создано из импорта", "text")

        operation_fields: list[str] = []

        def ensure_operation(name: str, kind: str, **extra: Any) -> None:
            if self._ensure_field(operations_table_id, name, kind, **extra):
                operation_fields.append(name)

        ensure_operation("Позиция поставки", "link_row", link_row_table_id=table_id)
        ensure_operation("Исходная операция", "link_row", link_row_table_id=int(operations_table_id))
        ensure_operation("Command ID", "text")
        ensure_operation("Причина корректировки", "long_text")
        ensure_operation(
            "Статус документа",
            "single_select",
            select_options=[
                {"value": "Создаётся", "color": "light-orange"},
                {"value": "Проведена", "color": "light-green"},
                {"value": "Требует восстановления", "color": "light-red"},
                {"value": "Отменена", "color": "dark-gray"},
            ],
        )

        supply_fields: list[str] = []

        def ensure_supply(name: str, kind: str, **extra: Any) -> None:
            if self._ensure_field(supplies_table_id, name, kind, **extra):
                supply_fields.append(name)

        ensure_supply("Import ID", "text")
        ensure_supply(
            "Статус импорта",
            "single_select",
            select_options=[
                {"value": "В процессе", "color": "light-orange"},
                {"value": "Завершён", "color": "light-green"},
                {"value": "Ошибка", "color": "light-red"},
            ],
        )

        return table_id, created_table, created_fields, operation_fields, supply_fields

    @staticmethod
    def _operation_direction(operation_type: str) -> tuple[int, int]:
        value = str(operation_type or "").casefold()
        if value == "приход":
            return 1, 0
        if value == "расход":
            return -1, 0
        if "передач" in value:
            return 0, 1
        if value == "возврат":
            return 0, -1
        return 0, 0

    def migrate_legacy_supply_lines(
        self,
        *,
        table_id: int,
        souvenirs_table_id: int,
        components_table_id: int,
        operations_table_id: int,
        supplies_table_id: int,
    ) -> tuple[int, int, list[str]]:
        existing_lines = self.list_rows(table_id)
        existing_keys = {
            (
                (link_ids(row.get("Поставка")) or [0])[0],
                (link_ids(row.get("Товар сувенирки")) or [0])[0],
                (link_ids(row.get("Комплектующее")) or [0])[0],
            )
            for row in existing_lines
        }
        supplies = self.list_rows(supplies_table_id)
        supply_id_by_row = {
            int(row["id"]): str(row.get("№ поставки") or "").strip()
            for row in supplies
        }
        operations = self.list_rows(operations_table_id)

        operation_totals: dict[tuple[int, int], dict[str, int]] = {}
        for operation in operations:
            linked_supply = (link_ids(operation.get("Поставка")) or [0])[0]
            supply_text = str(operation.get("ID поставки") or "").strip()
            if not linked_supply and supply_text:
                linked_supply = next(
                    (row_id for row_id, value in supply_id_by_row.items() if value == supply_text),
                    0,
                )
            if not linked_supply:
                continue
            received_direction, transfer_direction = self._operation_direction(
                select_text(operation.get("Тип операции"))
            )
            quantity = as_int(operation.get("Количество"))
            for product_id in link_ids(operation.get("Товар сувенирки")):
                totals = operation_totals.setdefault((linked_supply, product_id), {"received": 0, "transferred": 0})
                totals["received"] = max(totals["received"] + received_direction * quantity, 0)
                totals["transferred"] = max(totals["transferred"] + transfer_direction * quantity, 0)

        items: list[dict[str, Any]] = []
        ambiguous: list[str] = []
        skipped = 0
        for product in self.list_rows(souvenirs_table_id):
            product_id = int(product["id"])
            links = link_ids(product.get("Поставки"))
            if not links:
                continue
            old_document = as_int(product.get("По документу, шт."))
            old_received = as_int(product.get("Получено по поставке, шт."))
            boxes = str(product.get("Номера коробок") or "")
            sku = str(product.get("Артикул") or f"row-{product_id}")
            multi = len(links) > 1
            if multi:
                ambiguous.append(sku)

            received_values = {
                supply_row_id: operation_totals.get((supply_row_id, product_id), {}).get("received", 0)
                for supply_row_id in links
            }
            if len(links) == 1 and received_values[links[0]] == 0:
                received_values[links[0]] = old_received
            received_sum = sum(received_values.values())
            remaining_document = max(old_document - received_sum, 0)
            latest_supply = max(links)

            for supply_row_id in links:
                key = (supply_row_id, product_id, 0)
                if key in existing_keys:
                    skipped += 1
                    continue
                received = received_values[supply_row_id]
                transferred = operation_totals.get((supply_row_id, product_id), {}).get("transferred", 0)
                if len(links) == 1:
                    document = max(old_document, received)
                else:
                    document = received + (remaining_document if supply_row_id == latest_supply else 0)
                    document = max(document, received)
                if multi:
                    status = "Требует проверки"
                    comment = (
                        "Автомиграция: SKU был связан с несколькими поставками. "
                        "Принятое восстановлено по операциям; количество по документу требует проверки."
                    )
                elif received >= document and document > 0:
                    status = "Получена полностью"
                    comment = "Автомиграция из старой схемы."
                elif received > 0:
                    status = "Частично получена"
                    comment = "Автомиграция из старой схемы."
                else:
                    status = "Ожидается"
                    comment = "Автомиграция из старой схемы."
                items.append(
                    {
                        "Строка поставки": f"{supply_id_by_row.get(supply_row_id, supply_row_id)} — {sku}",
                        "Поставка": [supply_row_id],
                        "Товар сувенирки": [product_id],
                        "По документу, шт.": document,
                        "Принято, шт.": received,
                        "Передано в бухгалтерию, шт.": transferred,
                        "Номера коробок": boxes if len(links) == 1 or supply_row_id == latest_supply else "",
                        "Статус": status,
                        "Комментарий": comment,
                        "Версия": 1,
                        "Активна": True,
                        "Command ID": f"MIG-{supply_row_id}-{product_id}",
                        "Создано из импорта": "legacy-2.4.0",
                    }
                )

        # Components did not have a reliable legacy supply relation in the old
        # schema. They are intentionally not invented during migration.
        del components_table_id
        self.create_rows(table_id, items)
        return len(items), skipped, sorted(set(ambiguous), key=str.casefold)

    def ensure_and_migrate(
        self,
        *,
        database_id: int,
        souvenirs_table_id: int,
        components_table_id: int,
        operations_table_id: int,
        supplies_table_id: int,
    ) -> SchemaMigrationReport:
        (
            table_id,
            created_table,
            created_fields,
            operation_fields,
            supply_fields,
        ) = self.ensure_supply_lines_table(
            database_id=database_id,
            souvenirs_table_id=souvenirs_table_id,
            components_table_id=components_table_id,
            operations_table_id=operations_table_id,
            supplies_table_id=supplies_table_id,
        )
        migrated, skipped, ambiguous = self.migrate_legacy_supply_lines(
            table_id=table_id,
            souvenirs_table_id=souvenirs_table_id,
            components_table_id=components_table_id,
            operations_table_id=operations_table_id,
            supplies_table_id=supplies_table_id,
        )
        return SchemaMigrationReport(
            table_id=table_id,
            created_table=created_table,
            created_fields=tuple(created_fields),
            migrated_lines=migrated,
            skipped_lines=skipped,
            ambiguous_skus=tuple(ambiguous),
            added_operation_fields=tuple(operation_fields),
            added_supply_fields=tuple(supply_fields),
        )
