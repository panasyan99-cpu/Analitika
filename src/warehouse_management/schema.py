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

SUPPLY_LINE_REQUIRED_FIELDS = (
    "Строка поставки", "Поставка", "Товар сувенирки", "Комплектующее",
    "По документу, шт.", "Принято, шт.", "Передано в бухгалтерию, шт.",
    "Номера коробок", "Статус", "Комментарий", "Версия", "Активна",
    "Command ID", "Создано из импорта",
)
OPERATION_REQUIRED_FIELDS = (
    "Позиция поставки", "Исходная операция", "Command ID",
    "Причина корректировки", "Статус документа",
)
SUPPLY_REQUIRED_FIELDS = (
    "Import ID", "Статус импорта", "Импорт обработано SKU",
    "Импорт всего SKU", "Последняя ошибка импорта",
)
SILVER_COMPONENT_FIELDS = (
    "Название", "Серебряная категория", "Серебро 925", "Покрытие",
    "Размер", "Единица учёта", "Продаётся отдельно", "Закупка USD/ед.",
)
SILVER_LINE_FIELDS = (
    "Оригинальное название", "Название", "Серебряная категория", "Покрытие",
    "Размер", "Единица учёта", "Вес партии, г", "Вес единицы, г",
    "Способ приёмки", "Вес при приёмке, г", "Расчётное количество по весу",
    "Погрешность веса, г",
    "Серебро RMB/г", "Работа RMB/г", "Цена RMB/г", "Сумма RMB",
    "Курс USD/RMB", "CIF, %", "Закупка USD/ед.",
    "Продажа USD при импорте", "Курс USD/VND при импорте",
    "Коэффициент при импорте", "Продажа VND при импорте",
    "Серебро 925", "Продаётся отдельно",
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


@dataclass(frozen=True)
class SchemaInspection:
    table_id: int
    table_exists: bool
    missing_supply_line_fields: tuple[str, ...] = ()
    missing_operation_fields: tuple[str, ...] = ()
    missing_supply_fields: tuple[str, ...] = ()
    missing_silver_component_fields: tuple[str, ...] = ()
    missing_silver_line_fields: tuple[str, ...] = ()
    error: str = ""

    @property
    def ready(self) -> bool:
        return bool(
            self.table_exists
            and self.table_id > 0
            and not self.missing_supply_line_fields
            and not self.missing_operation_fields
            and not self.missing_supply_fields
            and not self.missing_silver_component_fields
            and not self.missing_silver_line_fields
            and not self.error
        )

    @property
    def change_count(self) -> int:
        return (
            (0 if self.table_exists else 1)
            + len(self.missing_supply_line_fields)
            + len(self.missing_operation_fields)
            + len(self.missing_supply_fields)
            + len(self.missing_silver_component_fields)
            + len(self.missing_silver_line_fields)
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "table_id": self.table_id,
            "table_exists": self.table_exists,
            "ready": self.ready,
            "change_count": self.change_count,
            "missing_supply_line_fields": list(self.missing_supply_line_fields),
            "missing_operation_fields": list(self.missing_operation_fields),
            "missing_supply_fields": list(self.missing_supply_fields),
            "missing_silver_component_fields": list(self.missing_silver_component_fields),
            "missing_silver_line_fields": list(self.missing_silver_line_fields),
            "error": self.error,
        }


class BaserowSchemaManager:
    """Baserow schema/migration client authenticated with a short-lived JWT.

    Database tokens deliberately cannot change table structure. This manager is
    used automatically by the private server account or from the bootstrap CLI.
    Credentials are read from server-side Secrets and never rendered in the UI.
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
                "User-Agent": "Princess-Analitika-Warehouse-Schema/2.5.0",
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

    def inspect_schema(
        self,
        *,
        database_id: int,
        souvenirs_table_id: int,
        components_table_id: int,
        operations_table_id: int,
        supplies_table_id: int,
        known_supply_lines_table_id: int = 0,
    ) -> SchemaInspection:
        """Read the warehouse schema and return a no-write repair plan."""
        try:
            tables = self.list_tables(database_id)
            existing = next(
                (table for table in tables if str(table.get("name") or "").strip().casefold() == SUPPLY_LINES_TABLE_NAME.casefold()),
                None,
            )
            table_id = int(existing.get("id") or 0) if existing else int(known_supply_lines_table_id or 0)
            table_exists = bool(existing or table_id)
            if not table_exists:
                return SchemaInspection(
                    table_id=0,
                    table_exists=False,
                    missing_supply_line_fields=SUPPLY_LINE_REQUIRED_FIELDS,
                    missing_operation_fields=OPERATION_REQUIRED_FIELDS,
                    missing_supply_fields=SUPPLY_REQUIRED_FIELDS,
                    missing_silver_component_fields=SILVER_COMPONENT_FIELDS,
                    missing_silver_line_fields=SILVER_LINE_FIELDS,
                )

            def missing(current_table_id: int, required: tuple[str, ...]) -> tuple[str, ...]:
                names = {str(field.get("name") or "") for field in self.fields(current_table_id)}
                return tuple(name for name in required if name not in names)

            return SchemaInspection(
                table_id=table_id,
                table_exists=True,
                missing_supply_line_fields=missing(table_id, SUPPLY_LINE_REQUIRED_FIELDS),
                missing_operation_fields=missing(int(operations_table_id), OPERATION_REQUIRED_FIELDS),
                missing_supply_fields=missing(int(supplies_table_id), SUPPLY_REQUIRED_FIELDS),
                missing_silver_component_fields=missing(int(components_table_id), SILVER_COMPONENT_FIELDS),
                missing_silver_line_fields=missing(table_id, SILVER_LINE_FIELDS),
            )
        except Exception as exc:
            return SchemaInspection(
                table_id=int(known_supply_lines_table_id or 0),
                table_exists=bool(known_supply_lines_table_id),
                error=str(exc),
            )

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
        ensure_supply("Импорт обработано SKU", "number", number_decimal_places=0, number_negative=False)
        ensure_supply("Импорт всего SKU", "number", number_decimal_places=0, number_negative=False)
        ensure_supply("Последняя ошибка импорта", "long_text")

        return table_id, created_table, created_fields, operation_fields, supply_fields

    def ensure_silver_fields(
        self,
        *,
        components_table_id: int,
        supply_lines_table_id: int,
        supplies_table_id: int | None = None,
    ) -> dict[str, list[str]]:
        """Create the additive 2.5.0 silver fields without changing legacy rows."""
        created_components: list[str] = []
        created_lines: list[str] = []
        component_names = {
            str(field.get("name") or "")
            for field in self.fields(int(components_table_id))
        }
        line_names = {
            str(field.get("name") or "")
            for field in self.fields(int(supply_lines_table_id))
        }

        def component(name: str, kind: str, **extra: Any) -> None:
            if name in component_names:
                return
            self.create_field(int(components_table_id), {"name": name, "type": kind, **extra})
            component_names.add(name)
            created_components.append(name)

        component("Название", "text")
        component("Серебряная категория", "text")
        component("Серебро 925", "boolean")
        component("Покрытие", "text")
        component("Размер", "text")
        component("Единица учёта", "text")
        component("Продаётся отдельно", "boolean")
        component("Закупка USD/ед.", "number", number_decimal_places=6, number_negative=False)

        def line(name: str, kind: str, **extra: Any) -> None:
            if name in line_names:
                return
            self.create_field(int(supply_lines_table_id), {"name": name, "type": kind, **extra})
            line_names.add(name)
            created_lines.append(name)

        line("Оригинальное название", "long_text")
        line("Название", "text")
        line("Серебряная категория", "text")
        line("Покрытие", "text")
        line("Размер", "text")
        line("Единица учёта", "text")
        line("Вес партии, г", "number", number_decimal_places=4, number_negative=False)
        line("Вес единицы, г", "number", number_decimal_places=6, number_negative=False)
        line("Способ приёмки", "text")
        line("Вес при приёмке, г", "number", number_decimal_places=4, number_negative=False)
        line("Расчётное количество по весу", "number", number_decimal_places=0, number_negative=False)
        line("Погрешность веса, г", "number", number_decimal_places=6, number_negative=False)
        line("Серебро RMB/г", "number", number_decimal_places=4, number_negative=False)
        line("Работа RMB/г", "number", number_decimal_places=4, number_negative=False)
        line("Цена RMB/г", "number", number_decimal_places=4, number_negative=False)
        line("Сумма RMB", "number", number_decimal_places=4, number_negative=False)
        line("Курс USD/RMB", "number", number_decimal_places=4, number_negative=False)
        line("CIF, %", "number", number_decimal_places=2, number_negative=False)
        line("Закупка USD/ед.", "number", number_decimal_places=6, number_negative=False)
        line("Продажа USD при импорте", "number", number_decimal_places=6, number_negative=False)
        line("Курс USD/VND при импорте", "number", number_decimal_places=0, number_negative=False)
        line("Коэффициент при импорте", "number", number_decimal_places=4, number_negative=False)
        line("Продажа VND при импорте", "number", number_decimal_places=0, number_negative=False)
        line("Серебро 925", "boolean")
        line("Продаётся отдельно", "boolean")

        # A supply can be registered before the physical goods arrive. Keep the
        # existing select options and add the explicit waiting state only once.
        if supplies_table_id:
            fields = self.fields(int(supplies_table_id))
            status_field = next(
                (field for field in fields if str(field.get("name") or "") == "Статус"),
                None,
            )
            if status_field and str(status_field.get("type") or "") == "single_select":
                options = list(status_field.get("select_options") or [])
                values = {str(option.get("value") or "").strip().casefold() for option in options}
                if "ожидается" not in values:
                    options.append({"value": "Ожидается", "color": "light-gray"})
                    self.update_field(int(status_field["id"]), {"select_options": options})

        return {"components": created_components, "supply_lines": created_lines}

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
        """Idempotently restore missing detail rows from both product catalogs.

        Historical versions stored the supply relation on the permanent product
        card.  The first migration only processed souvenirs, which left component
        and Silver 925 supplies with a visible header but no detail rows.
        """
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

        # Key: (supply row, section, product row).
        operation_totals: dict[tuple[int, str, int], dict[str, int]] = {}
        operation_fields = (
            ("Сувенирка", "Товар сувенирки"),
            ("Комплектующие", "Комплектующее"),
        )
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
            for section, field_name in operation_fields:
                for product_id in link_ids(operation.get(field_name)):
                    totals = operation_totals.setdefault(
                        (linked_supply, section, product_id),
                        {"received": 0, "transferred": 0},
                    )
                    totals["received"] = max(
                        totals["received"] + received_direction * quantity, 0
                    )
                    totals["transferred"] = max(
                        totals["transferred"] + transfer_direction * quantity, 0
                    )

        def quantity_value(product: dict[str, Any], names: tuple[str, ...]) -> int:
            for name in names:
                value = product.get(name)
                if value not in (None, ""):
                    return as_int(value)
            return 0

        items: list[dict[str, Any]] = []
        ambiguous: list[str] = []
        skipped = 0
        catalogs = (
            ("Сувенирка", int(souvenirs_table_id), "Товар сувенирки"),
            ("Комплектующие", int(components_table_id), "Комплектующее"),
        )
        for section, catalog_table_id, line_link_field in catalogs:
            for product in self.list_rows(catalog_table_id):
                product_id = int(product["id"])
                links = link_ids(product.get("Поставки")) or link_ids(product.get("Поставка"))
                if not links:
                    continue
                old_document = quantity_value(
                    product,
                    ("По документу, шт.", "Количество по документу", "Количество", "Qty", "QTY"),
                )
                old_received = quantity_value(
                    product,
                    ("Получено по поставке, шт.", "Принято, шт.", "Получено", "Факт"),
                )
                boxes = str(product.get("Номера коробок") or product.get("Коробки") or "")
                sku = str(product.get("Артикул") or f"row-{product_id}")
                multi = len(links) > 1
                if multi:
                    ambiguous.append(sku)

                received_values = {
                    supply_row_id: operation_totals.get(
                        (supply_row_id, section, product_id), {}
                    ).get("received", 0)
                    for supply_row_id in links
                }
                if len(links) == 1 and received_values[links[0]] == 0:
                    received_values[links[0]] = old_received
                received_sum = sum(received_values.values())
                remaining_document = max(old_document - received_sum, 0)
                latest_supply = max(links)

                for supply_row_id in links:
                    key = (
                        supply_row_id,
                        product_id if section == "Сувенирка" else 0,
                        product_id if section == "Комплектующие" else 0,
                    )
                    if key in existing_keys:
                        skipped += 1
                        continue
                    received = received_values[supply_row_id]
                    transferred = operation_totals.get(
                        (supply_row_id, section, product_id), {}
                    ).get("transferred", 0)
                    if len(links) == 1:
                        document = max(old_document, received)
                    else:
                        document = received + (remaining_document if supply_row_id == latest_supply else 0)
                        document = max(document, received)

                    quantity_missing = document <= 0
                    if multi or quantity_missing:
                        status = "Требует проверки"
                        reasons = []
                        if multi:
                            reasons.append("SKU связан с несколькими поставками")
                        if quantity_missing:
                            reasons.append("количество по документу не удалось восстановить")
                        comment = "Автомиграция: " + "; ".join(reasons) + "."
                    elif received >= document:
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
                            line_link_field: [product_id],
                            "По документу, шт.": document,
                            "Принято, шт.": received,
                            "Передано в бухгалтерию, шт.": transferred,
                            "Номера коробок": boxes if len(links) == 1 or supply_row_id == latest_supply else "",
                            "Статус": status,
                            "Комментарий": comment,
                            "Версия": 1,
                            "Активна": True,
                            "Command ID": f"MIG-{section[:3].upper()}-{supply_row_id}-{product_id}",
                            "Создано из импорта": "legacy-recovery-2.6.0",
                        }
                    )
                    existing_keys.add(key)

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
