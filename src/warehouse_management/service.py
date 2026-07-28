from __future__ import annotations

from dataclasses import asdict
from datetime import date
from pathlib import Path
from typing import Any, Iterable

from .client import WarehouseClient, WarehouseClientError, as_int, link_ids, select_text
from .models import (
    CatalogItem,
    Product,
    SupplySummary,
    normalize_stone_names,
    split_multi_values,
)


class WarehouseServiceError(RuntimeError):
    pass


class WarehouseService:
    """Business layer shared by all Streamlit warehouse workspaces."""

    def __init__(self, client: WarehouseClient) -> None:
        self.client = client
        self.config = client.config

    def table_id(self, section: str) -> int:
        return (
            self.config.souvenirs_table_id
            if section == "Сувенирка"
            else self.config.components_table_id
        )

    def catalog_rows(self, section: str) -> list[dict[str, Any]]:
        return self.client.list_rows(self.table_id(section))

    def catalog(self, section: str) -> list[CatalogItem]:
        result: list[CatalogItem] = []
        for row in self.catalog_rows(section):
            sku = str(row.get("Артикул") or "").strip()
            if not sku:
                continue
            result.append(
                CatalogItem(
                    row_id=int(row["id"]),
                    sku=sku,
                    section=section,
                    balance=as_int(row.get("Остаток")),
                    boxes=str(row.get("Номера коробок") or ""),
                    category=select_text(row.get("Категория")),
                    material=select_text(row.get("Материал")),
                    stone=select_text(row.get("Камень")),
                    color=select_text(row.get("Цвет")),
                    photo=row.get("Фото"),
                    min_balance=as_int(row.get("Минимальный остаток"), 10) or 10,
                    raw=row,
                )
            )
        result.sort(key=lambda item: (item.balance <= 0, item.sku.casefold()))
        return result

    def _supply_rows(self) -> list[dict[str, Any]]:
        return self.client.list_rows(self.config.supplies_table_id)

    def find_supply(self, supply_id: str) -> dict[str, Any] | None:
        target = str(supply_id or "").strip().casefold()
        return next(
            (
                row
                for row in self._supply_rows()
                if str(row.get("№ поставки") or "").strip().casefold() == target
            ),
            None,
        )

    def find_or_create_supply(
        self,
        supply_id: str,
        *,
        supplier: str = "",
        invoice: str = "",
        comment: str = "",
    ) -> dict[str, Any]:
        supply_id = str(supply_id or "").strip()
        if not supply_id:
            raise WarehouseServiceError("Не указан номер поставки.")
        existing = self.find_supply(supply_id)
        if existing:
            return existing
        return self.client.create_row(
            self.config.supplies_table_id,
            {
                "№ поставки": supply_id,
                "Дата": str(date.today()),
                "Поставщик": supplier,
                "Invoice": invoice,
                "Статус": "Частично получена",
                "Комментарий": comment or "Создано в Analitika Web",
            },
        )

    @property
    def has_supply_lines(self) -> bool:
        return int(getattr(self.config, "supply_lines_table_id", 0) or 0) > 0

    def _supply_line_rows(self) -> list[dict[str, Any]]:
        if not self.has_supply_lines:
            return []
        return self.client.list_rows(int(self.config.supply_lines_table_id))

    @staticmethod
    def _linked_row_id(row: dict[str, Any], names: Iterable[str]) -> int:
        for name in names:
            ids = link_ids(row.get(name))
            if ids:
                return ids[0]
        return 0

    def supply_summaries(self) -> list[SupplySummary]:
        supplies = self._supply_rows()
        result: list[SupplySummary] = []

        if self.has_supply_lines:
            by_supply: dict[int, list[dict[str, Any]]] = {}
            for line in self._supply_line_rows():
                supply_row_id = self._linked_row_id(line, ("Поставка",))
                if supply_row_id:
                    by_supply.setdefault(supply_row_id, []).append(line)
            for supply in supplies:
                lines = by_supply.get(int(supply["id"]), [])
                if not lines:
                    continue
                document = sum(as_int(line.get("По документу, шт.")) for line in lines)
                received = sum(as_int(line.get("Принято, шт.")) for line in lines)
                result.append(
                    SupplySummary(
                        row_id=int(supply["id"]),
                        supply_id=str(supply.get("№ поставки") or ""),
                        date=str(supply.get("Дата") or supply.get("Дата создания") or ""),
                        supplier=str(supply.get("Поставщик") or ""),
                        status=select_text(supply.get("Статус")),
                        sku_total=len(lines),
                        sku_received=sum(as_int(line.get("Принято, шт.")) > 0 for line in lines),
                        qty_document=document,
                        qty_received=received,
                        qty_waiting=max(document - received, 0),
                        raw=supply,
                    )
                )
        else:
            products = self.catalog_rows("Сувенирка")
            by_supply: dict[int, list[dict[str, Any]]] = {}
            for row in products:
                for supply_row_id in link_ids(row.get("Поставки")):
                    by_supply.setdefault(supply_row_id, []).append(row)
            for supply in supplies:
                linked = by_supply.get(int(supply["id"]), [])
                if not linked:
                    continue
                document = sum(as_int(row.get("По документу, шт.")) for row in linked)
                received = sum(as_int(row.get("Получено по поставке, шт.")) for row in linked)
                result.append(
                    SupplySummary(
                        row_id=int(supply["id"]),
                        supply_id=str(supply.get("№ поставки") or ""),
                        date=str(supply.get("Дата") or supply.get("Дата создания") or ""),
                        supplier=str(supply.get("Поставщик") or ""),
                        status=select_text(supply.get("Статус")),
                        sku_total=len(linked),
                        sku_received=sum(as_int(row.get("Получено по поставке, шт.")) > 0 for row in linked),
                        qty_document=document,
                        qty_received=received,
                        qty_waiting=max(document - received, 0),
                        raw=supply,
                    )
                )
        result.sort(key=lambda item: (item.qty_waiting <= 0, item.date, item.supply_id), reverse=False)
        return result

    def supply_products(self, supply: SupplySummary | int) -> list[dict[str, Any]]:
        supply_row_id = supply.row_id if isinstance(supply, SupplySummary) else int(supply)
        catalog = {int(row["id"]): row for row in self.catalog_rows("Сувенирка")}
        if self.has_supply_lines:
            result: list[dict[str, Any]] = []
            for line in self._supply_line_rows():
                if supply_row_id not in link_ids(line.get("Поставка")):
                    continue
                product_id = self._linked_row_id(line, ("Товар сувенирки", "Товар"))
                product = catalog.get(product_id, {})
                result.append(
                    {
                        **product,
                        "_line_id": int(line["id"]),
                        "_supply_row_id": supply_row_id,
                        "_document": as_int(line.get("По документу, шт.")),
                        "_received": as_int(line.get("Принято, шт.")),
                        "_transferred": as_int(line.get("Передано в бухгалтерию, шт.")),
                        "_boxes": str(line.get("Номера коробок") or product.get("Номера коробок") or ""),
                    }
                )
            return sorted(result, key=lambda row: str(row.get("Артикул") or ""))

        result = []
        for row in catalog.values():
            if supply_row_id in link_ids(row.get("Поставки")):
                result.append(
                    {
                        **row,
                        "_line_id": 0,
                        "_supply_row_id": supply_row_id,
                        "_document": as_int(row.get("По документу, шт.")),
                        "_received": as_int(row.get("Получено по поставке, шт.")),
                        "_transferred": 0,
                        "_boxes": str(row.get("Номера коробок") or ""),
                    }
                )
        return sorted(result, key=lambda row: str(row.get("Артикул") or ""))

    def transferred_by_supply(self, supply_id: str) -> dict[int, int]:
        result: dict[int, int] = {}
        for operation in self.client.list_rows(self.config.operations_table_id):
            if str(operation.get("ID поставки") or "").strip() != supply_id:
                continue
            operation_type = select_text(operation.get("Тип операции")).casefold()
            if "передач" in operation_type:
                direction = 1
            elif "возврат" in operation_type:
                direction = -1
            else:
                continue
            quantity = as_int(operation.get("Количество")) * direction
            for product_id in link_ids(operation.get("Товар сувенирки")):
                result[product_id] = max(result.get(product_id, 0) + quantity, 0)
        return result

    def _product_payload(
        self,
        product: Product,
        *,
        supply_row_id: int,
        existing: dict[str, Any] | None,
        photo: dict[str, Any] | None,
    ) -> dict[str, Any]:
        supply_links = list(
            dict.fromkeys(
                (link_ids(existing.get("Поставки")) if existing else []) + [supply_row_id]
            )
        )
        payload: dict[str, Any] = {
            "Артикул": product.sku,
            "Категория": product.category or None,
            "Материал": split_multi_values(product.material),
            "Камень": split_multi_values(normalize_stone_names(product.stone)),
            "Цвет": split_multi_values(product.color),
            "Вес 1 шт. (кг)": product.unit_weight_kg,
            "Номера коробок": product.boxes,
            "Поставки": supply_links,
            "Минимальный остаток": 10,
            "Активный SKU": True,
            "Комментарий": product.comment,
        }
        if not self.has_supply_lines:
            payload["По документу, шт."] = product.qty_document
            payload["Получено по поставке, шт."] = product.actual_qty or 0
        if photo:
            payload["Фото"] = [photo]
        return payload

    def create_supply_from_products(
        self,
        *,
        supply_id: str,
        supplier: str,
        invoice: str,
        comment: str,
        products: list[Product],
        allow_reused_sku_compatibility: bool = False,
    ) -> dict[str, Any]:
        products = [product for product in products if product.sku and product.qty_document > 0]
        if not products:
            raise WarehouseServiceError("В поставке нет корректных строк с SKU и количеством.")
        if len({product.sku.casefold() for product in products}) != len(products):
            raise WarehouseServiceError("В файле есть повторяющиеся SKU после нормализации.")

        existing_supply = self.find_supply(supply_id)
        if existing_supply:
            existing_products = self.supply_products(int(existing_supply["id"]))
            existing_operations = [
                row
                for row in self.client.list_rows(self.config.operations_table_id)
                if str(row.get("ID поставки") or "").strip() == str(supply_id).strip()
            ]
            if existing_products or existing_operations:
                raise WarehouseServiceError(
                    f"Поставка {supply_id} уже содержит позиции или операции. "
                    "Повторное создание заблокировано. Используйте раздел «Поставки» для исправления."
                )

        supply = self.find_or_create_supply(
            supply_id,
            supplier=supplier,
            invoice=invoice,
            comment=comment,
        )
        supply_row_id = int(supply["id"])
        table_id = self.config.souvenirs_table_id
        rows = self.client.list_rows(table_id)
        by_sku = {
            str(row.get("Артикул") or "").strip().casefold(): row
            for row in rows
            if str(row.get("Артикул") or "").strip()
        }

        reused = [
            product.sku
            for product in products
            if (
                (existing := by_sku.get(product.sku.casefold()))
                and any(link != supply_row_id for link in link_ids(existing.get("Поставки")))
            )
        ]
        if reused and not self.has_supply_lines and not allow_reused_sku_compatibility:
            raise WarehouseServiceError(
                "Эти SKU уже связаны с другой поставкой: "
                + ", ".join(reused[:15])
                + ("…" if len(reused) > 15 else "")
                + ". Для безопасной истории создайте таблицу «Позиции поставок» "
                "или подтвердите режим совместимости."
            )

        batch_id = self.client.batch_id("SUP")
        operations: list[dict[str, Any]] = []
        line_payloads: list[dict[str, Any]] = []
        created = 0
        updated = 0
        photos = 0
        received = 0

        for product in products:
            existing = by_sku.get(product.sku.casefold())
            photo = None
            if product.image_path and Path(product.image_path).exists():
                try:
                    photo = self.client.upload_file(Path(product.image_path))
                    photos += 1
                except WarehouseClientError:
                    # Data import must not fail completely because one image upload failed.
                    photo = None
            payload = self._product_payload(
                product,
                supply_row_id=supply_row_id,
                existing=existing,
                photo=photo,
            )
            if existing:
                self.client.batch_update(table_id, [{"id": int(existing["id"]), **payload}])
                row_id = int(existing["id"])
                updated += 1
            else:
                row = self.client.create_row(table_id, payload)
                row_id = int(row["id"])
                by_sku[product.sku.casefold()] = row
                created += 1

            actual = int(product.actual_qty or 0)
            if self.has_supply_lines:
                line_payloads.append(
                    {
                        "Строка поставки": f"{supply_id} — {product.sku}",
                        "Поставка": [supply_row_id],
                        "Товар сувенирки": [row_id],
                        "По документу, шт.": product.qty_document,
                        "Принято, шт.": actual,
                        "Передано в бухгалтерию, шт.": 0,
                        "Номера коробок": product.boxes,
                        "Статус": "Получена полностью" if actual >= product.qty_document else "Частично получена" if actual else "Ожидается",
                        "Комментарий": product.comment,
                        "Версия": 1,
                        "Активна": True,
                    }
                )
            if actual > 0:
                operations.append(
                    {
                        "Операция": f"{batch_id} — {product.sku}",
                        "Тип операции": "Приход",
                        "Раздел": "Сувенирка",
                        "Товар сувенирки": [row_id],
                        "Количество": actual,
                        "Поставка": [supply_row_id],
                        "ID поставки": supply_id,
                        "Batch ID": batch_id,
                        "Комментарий": "Импорт поставки из Analitika Web",
                    }
                )
                received += actual

        if line_payloads:
            self.client.batch_create(int(self.config.supply_lines_table_id), line_payloads)
        if operations:
            self.client.create_operations(operations, batch_id=batch_id)

        complete = all((product.actual_qty or 0) >= product.qty_document for product in products)
        self.client.batch_update(
            self.config.supplies_table_id,
            [
                {
                    "id": supply_row_id,
                    "Статус": "Получена полностью" if complete else "Частично получена",
                }
            ],
        )
        return {
            "supply_id": supply_id,
            "batch_id": batch_id if operations else "",
            "sku": len(products),
            "created": created,
            "updated": updated,
            "photos": photos,
            "received": received,
            "waiting": sum(product.waiting_qty for product in products),
            "compatibility_mode": not self.has_supply_lines,
        }

    def receive_supply(
        self,
        supply: SupplySummary,
        quantities: dict[int, int],
    ) -> dict[str, Any]:
        rows = self.supply_products(supply)
        selected = [row for row in rows if as_int(quantities.get(int(row["id"]))) > 0]
        if not selected:
            raise WarehouseServiceError("Не указано количество для приёмки.")
        batch_id = self.client.batch_id("REC")
        operations: list[dict[str, Any]] = []
        product_updates: list[dict[str, Any]] = []
        line_updates: list[dict[str, Any]] = []
        total = 0
        for row in selected:
            row_id = int(row["id"])
            quantity = as_int(quantities.get(row_id))
            waiting = max(as_int(row.get("_document")) - as_int(row.get("_received")), 0)
            if quantity > waiting:
                raise WarehouseServiceError(
                    f"{row.get('Артикул')}: можно принять не более {waiting} шт."
                )
            operations.append(
                {
                    "Операция": f"{batch_id} — {row.get('Артикул')}",
                    "Тип операции": "Приход",
                    "Раздел": "Сувенирка",
                    "Товар сувенирки": [row_id],
                    "Количество": quantity,
                    "Поставка": [supply.row_id],
                    "ID поставки": supply.supply_id,
                    "Batch ID": batch_id,
                    "Комментарий": "Доприёмка из Analitika Web",
                }
            )
            new_received = as_int(row.get("_received")) + quantity
            if self.has_supply_lines and as_int(row.get("_line_id")):
                line_updates.append(
                    {
                        "id": as_int(row.get("_line_id")),
                        "Принято, шт.": new_received,
                        "Статус": "Получена полностью" if new_received >= as_int(row.get("_document")) else "Частично получена",
                    }
                )
            else:
                product_updates.append(
                    {
                        "id": row_id,
                        "Получено по поставке, шт.": new_received,
                    }
                )
            total += quantity

        self.client.create_operations(operations, batch_id=batch_id)
        if product_updates:
            self.client.batch_update(self.config.souvenirs_table_id, product_updates)
        if line_updates:
            self.client.batch_update(int(self.config.supply_lines_table_id), line_updates)

        refreshed = self.supply_products(supply)
        complete = bool(refreshed) and all(
            as_int(row.get("_received")) >= as_int(row.get("_document")) for row in refreshed
        )
        self.client.batch_update(
            self.config.supplies_table_id,
            [{"id": supply.row_id, "Статус": "Получена полностью" if complete else "Частично получена"}],
        )
        return {"batch_id": batch_id, "sku": len(operations), "quantity": total}

    def transfer_supply(
        self,
        supply: SupplySummary,
        quantities: dict[int, int],
        *,
        comment: str = "",
    ) -> dict[str, Any]:
        rows = self.supply_products(supply)
        already = self.transferred_by_supply(supply.supply_id)
        selected: list[tuple[dict[str, Any], int]] = []
        for row in rows:
            row_id = int(row["id"])
            quantity = as_int(quantities.get(row_id))
            if quantity <= 0:
                continue
            received = as_int(row.get("_received"))
            transferred = (
                as_int(row.get("_transferred"))
                if self.has_supply_lines
                else already.get(row_id, 0)
            )
            available_from_supply = max(received - transferred, 0)
            stock = as_int(row.get("Остаток"))
            maximum = min(available_from_supply, stock)
            if quantity > maximum:
                raise WarehouseServiceError(
                    f"{row.get('Артикул')}: можно передать не более {maximum} шт."
                )
            selected.append((row, quantity))
        if not selected:
            raise WarehouseServiceError("Не выбраны товары для передачи.")

        batch_id = self.client.batch_id("ACC")
        operations: list[dict[str, Any]] = []
        line_updates: list[dict[str, Any]] = []
        for row, quantity in selected:
            row_id = int(row["id"])
            operations.append(
                {
                    "Операция": f"{batch_id} — {row.get('Артикул')}",
                    "Тип операции": "Передача в бухгалтерию",
                    "Раздел": "Сувенирка",
                    "Товар сувенирки": [row_id],
                    "Количество": quantity,
                    "Поставка": [supply.row_id],
                    "ID поставки": supply.supply_id,
                    "Batch ID": batch_id,
                    "Комментарий": comment or f"Поставка {supply.supply_id}",
                }
            )
            if self.has_supply_lines and as_int(row.get("_line_id")):
                line_updates.append(
                    {
                        "id": as_int(row.get("_line_id")),
                        "Передано в бухгалтерию, шт.": as_int(row.get("_transferred")) + quantity,
                    }
                )
        self.client.create_operations(operations, batch_id=batch_id)
        if line_updates:
            self.client.batch_update(int(self.config.supply_lines_table_id), line_updates)
        return {
            "batch_id": batch_id,
            "sku": len(selected),
            "quantity": sum(quantity for _, quantity in selected),
        }

    def manual_operation(
        self,
        *,
        operation_type: str,
        section: str,
        quantities: dict[int, int],
        comment: str = "",
        supply_id: str = "",
    ) -> dict[str, Any]:
        catalog = {item.row_id: item for item in self.catalog(section)}
        selected = [
            (catalog[row_id], as_int(quantity))
            for row_id, quantity in quantities.items()
            if row_id in catalog and as_int(quantity) > 0
        ]
        if not selected:
            raise WarehouseServiceError("Не выбрано ни одной позиции.")
        if operation_type == "Передача в бухгалтерию":
            for item, quantity in selected:
                if quantity > item.balance:
                    raise WarehouseServiceError(
                        f"{item.sku}: требуется {quantity}, доступно {item.balance}."
                    )
        prefix = "ACC" if operation_type == "Передача в бухгалтерию" else "REC"
        batch_id = self.client.batch_id(prefix)
        link_field = "Товар сувенирки" if section == "Сувенирка" else "Комплектующее"
        operations = [
            {
                "Операция": f"{batch_id} — {item.sku}",
                "Тип операции": operation_type,
                "Раздел": section,
                link_field: [item.row_id],
                "Количество": quantity,
                "ID поставки": supply_id,
                "Batch ID": batch_id,
                "Комментарий": comment,
            }
            for item, quantity in selected
        ]
        self.client.create_operations(operations, batch_id=batch_id)
        return {
            "batch_id": batch_id,
            "sku": len(selected),
            "quantity": sum(quantity for _, quantity in selected),
        }

    def correct_operation(
        self,
        operation: dict[str, Any],
        *,
        quantity: int,
        comment: str,
    ) -> dict[str, Any]:
        quantity = as_int(quantity)
        original_quantity = as_int(operation.get("Количество"))
        if quantity <= 0 or quantity > original_quantity:
            raise WarehouseServiceError(
                f"Количество корректировки должно быть от 1 до {original_quantity}."
            )
        original_type = select_text(operation.get("Тип операции"))
        original_is_incoming = "приход" in original_type.casefold()
        batch_id = self.client.batch_id("COR")
        # Existing Baserow balance formulas classify the direction by operation
        # type. Use a positive quantity with the opposite canonical type rather
        # than relying on a negative number in a generic correction row.
        reverse_type = "Расход" if original_is_incoming else "Возврат"
        payload = {
            "Операция": f"{batch_id} — отмена {operation.get('Операция') or operation.get('id')}",
            "Тип операции": reverse_type,
            "Раздел": select_text(operation.get("Раздел")),
            "Товар сувенирки": link_ids(operation.get("Товар сувенирки")),
            "Комплектующее": link_ids(operation.get("Комплектующее")),
            "Количество": quantity,
            "Поставка": link_ids(operation.get("Поставка")),
            "ID поставки": str(operation.get("ID поставки") or ""),
            "Batch ID": batch_id,
            "Комментарий": comment or f"Корректировка операции {operation.get('Batch ID') or operation.get('id')}",
        }
        self.client.create_operations([payload], batch_id=batch_id)

        supply_id = str(operation.get("ID поставки") or "").strip()
        product_ids = link_ids(operation.get("Товар сувенирки"))
        if supply_id and product_ids:
            supply_row = self.find_supply(supply_id)
            if supply_row:
                summary = next(
                    (item for item in self.supply_summaries() if item.row_id == int(supply_row["id"])),
                    None,
                )
                if summary:
                    rows = {
                        int(row["id"]): row
                        for row in self.supply_products(summary)
                    }
                    if original_is_incoming:
                        product_updates: list[dict[str, Any]] = []
                        line_updates: list[dict[str, Any]] = []
                        for product_id in product_ids:
                            row = rows.get(product_id)
                            if not row:
                                continue
                            updated = max(as_int(row.get("_received")) - quantity, 0)
                            if self.has_supply_lines and as_int(row.get("_line_id")):
                                line_updates.append(
                                    {
                                        "id": as_int(row.get("_line_id")),
                                        "Принято, шт.": updated,
                                        "Статус": (
                                            "Получена полностью"
                                            if updated >= as_int(row.get("_document"))
                                            else "Частично получена"
                                            if updated > 0
                                            else "Ожидается"
                                        ),
                                    }
                                )
                            else:
                                product_updates.append(
                                    {
                                        "id": product_id,
                                        "Получено по поставке, шт.": updated,
                                    }
                                )
                        if product_updates:
                            self.client.batch_update(self.config.souvenirs_table_id, product_updates)
                        if line_updates:
                            self.client.batch_update(int(self.config.supply_lines_table_id), line_updates)
                    elif self.has_supply_lines:
                        line_updates = []
                        for product_id in product_ids:
                            row = rows.get(product_id)
                            if row and as_int(row.get("_line_id")):
                                line_updates.append(
                                    {
                                        "id": as_int(row.get("_line_id")),
                                        "Передано в бухгалтерию, шт.": max(
                                            as_int(row.get("_transferred")) - quantity,
                                            0,
                                        ),
                                    }
                                )
                        if line_updates:
                            self.client.batch_update(int(self.config.supply_lines_table_id), line_updates)

                    refreshed = self.supply_products(summary)
                    complete = bool(refreshed) and all(
                        as_int(row.get("_received")) >= as_int(row.get("_document"))
                        for row in refreshed
                    )
                    self.client.batch_update(
                        self.config.supplies_table_id,
                        [
                            {
                                "id": summary.row_id,
                                "Статус": "Получена полностью" if complete else "Частично получена",
                            }
                        ],
                    )
        return {"batch_id": batch_id, "quantity": quantity}

    def add_catalog_item(
        self,
        *,
        section: str,
        sku: str,
        category: str,
        material: str,
        stone: str,
        color: str,
        boxes: str,
        minimum: int,
        comment: str,
        photo_path: Path | None = None,
    ) -> dict[str, Any]:
        sku = str(sku or "").strip()
        if not sku:
            raise WarehouseServiceError("Артикул обязателен.")
        existing = next((item for item in self.catalog(section) if item.sku.casefold() == sku.casefold()), None)
        if existing:
            raise WarehouseServiceError(f"Артикул {sku} уже существует.")
        payload: dict[str, Any] = {
            "Артикул": sku,
            "Категория": category or None,
            "Материал": split_multi_values(material),
            "Камень": split_multi_values(normalize_stone_names(stone)),
            "Цвет": split_multi_values(color),
            "Номера коробок": boxes,
            "Минимальный остаток": max(as_int(minimum), 1),
            "Активный SKU" if section == "Сувенирка" else "Активно": True,
            "Комментарий": comment,
        }
        if photo_path and photo_path.exists():
            payload["Фото"] = [self.client.upload_file(photo_path)]
        return self.client.create_row(self.table_id(section), payload)

    def update_catalog_item(
        self,
        section: str,
        row_id: int,
        payload: dict[str, Any],
        *,
        photo_path: Path | None = None,
    ) -> None:
        clean = dict(payload)
        if "Камень" in clean:
            clean["Камень"] = split_multi_values(normalize_stone_names(str(clean["Камень"] or "")))
        if "Материал" in clean:
            clean["Материал"] = split_multi_values(str(clean["Материал"] or ""))
        if "Цвет" in clean:
            clean["Цвет"] = split_multi_values(str(clean["Цвет"] or ""))
        if photo_path and photo_path.exists():
            clean["Фото"] = [self.client.upload_file(photo_path)]
        self.client.batch_update(self.table_id(section), [{"id": int(row_id), **clean}])

    def deactivate_or_delete_catalog_item(self, section: str, row_id: int) -> str:
        operation_link = "Товар сувенирки" if section == "Сувенирка" else "Комплектующее"
        has_operations = any(
            int(row_id) in link_ids(row.get(operation_link))
            for row in self.client.list_rows(self.config.operations_table_id)
        )
        if has_operations:
            field = "Активный SKU" if section == "Сувенирка" else "Активно"
            self.client.batch_update(self.table_id(section), [{"id": int(row_id), field: False}])
            return "deactivated"
        self.client.delete_row(self.table_id(section), int(row_id))
        return "deleted"

    def remove_waiting_from_supply(self, supply: SupplySummary, product_ids: list[int]) -> int:
        selected = [row for row in self.supply_products(supply) if int(row["id"]) in product_ids]
        blocked = [str(row.get("Артикул") or "") for row in selected if as_int(row.get("_received")) > 0]
        if blocked:
            raise WarehouseServiceError(
                "Нельзя убрать уже принятые позиции: " + ", ".join(blocked[:12])
            )
        if self.has_supply_lines:
            for row in selected:
                if as_int(row.get("_line_id")):
                    self.client.delete_row(int(self.config.supply_lines_table_id), as_int(row.get("_line_id")))
        else:
            updates = []
            for row in selected:
                links = [value for value in link_ids(row.get("Поставки")) if value != supply.row_id]
                payload: dict[str, Any] = {"id": int(row["id"]), "Поставки": links}
                if not links:
                    payload["По документу, шт."] = 0
                    payload["Получено по поставке, шт."] = 0
                updates.append(payload)
            if updates:
                self.client.batch_update(self.config.souvenirs_table_id, updates)
        return len(selected)

    def delete_empty_supply(self, supply: SupplySummary) -> None:
        if supply.qty_received > 0:
            raise WarehouseServiceError(
                "Поставка уже содержит приёмку. Используйте корректирующие операции, а не удаление."
            )
        rows = self.supply_products(supply)
        self.remove_waiting_from_supply(supply, [int(row["id"]) for row in rows])
        self.client.delete_row(self.config.supplies_table_id, supply.row_id)

    def diagnostics(self) -> dict[str, Any]:
        souvenirs = self.catalog("Сувенирка")
        components = self.catalog("Комплектующие")
        operations = self.client.list_rows(self.config.operations_table_id)
        supplies = self.supply_summaries()
        duplicate_skus: list[str] = []
        for items in (souvenirs, components):
            seen: set[str] = set()
            for item in items:
                key = item.sku.casefold()
                if key in seen and item.sku not in duplicate_skus:
                    duplicate_skus.append(item.sku)
                seen.add(key)
        return {
            "souvenir_sku": len(souvenirs),
            "component_sku": len(components),
            "operations": len(operations),
            "supplies": len(supplies),
            "duplicate_sku": duplicate_skus,
            "without_photo": [item.sku for item in [*souvenirs, *components] if not item.photo],
            "without_category": [item.sku for item in [*souvenirs, *components] if not item.category],
            "supply_lines_mode": self.has_supply_lines,
        }
