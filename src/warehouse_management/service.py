from __future__ import annotations

from collections import defaultdict
from dataclasses import asdict
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
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


def _as_float(value: Any, default: float = 0.0) -> float:
    if value in (None, ""):
        return default
    try:
        return float(str(value).replace(" ", "").replace(",", "."))
    except (TypeError, ValueError):
        return default


class WarehouseService:
    """Business layer shared by all Streamlit warehouse workspaces."""

    @staticmethod
    def estimate_quantity_from_weight(
        weight_g: float,
        unit_weight_g: float,
        *,
        maximum: int | None = None,
    ) -> int:
        """Convert a clean measured weight into units using half-up rounding.

        Decimal arithmetic is intentional here: Python's built-in ``round`` uses
        bankers rounding and can turn an exact .5 into an unexpectedly lower
        quantity for warehouse operators.
        """
        try:
            weight = Decimal(str(weight_g))
            unit = Decimal(str(unit_weight_g))
        except (InvalidOperation, TypeError, ValueError) as exc:
            raise WarehouseServiceError("Вес должен быть числом.") from exc
        if weight < 0:
            raise WarehouseServiceError("Вес не может быть отрицательным.")
        if unit <= 0:
            raise WarehouseServiceError("Для позиции не указан корректный средний вес единицы.")
        quantity = int((weight / unit).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
        if maximum is not None:
            quantity = min(quantity, max(int(maximum), 0))
        return max(quantity, 0)

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

    def catalog(
        self,
        section: str,
        *,
        include_inactive: bool = False,
    ) -> list[CatalogItem]:
        result: list[CatalogItem] = []
        active_field = "Активный SKU" if section == "Сувенирка" else "Активно"
        for row in self.catalog_rows(section):
            sku = str(row.get("Артикул") or "").strip()
            if not sku:
                continue
            active = row.get(active_field) is not False
            if not include_inactive and not active:
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
                    active=active,
                    name=str(row.get("Название") or "").strip(),
                    silver_category=str(row.get("Серебряная категория") or "").strip(),
                    silver_925=bool(row.get("Серебро 925")),
                    plating=str(row.get("Покрытие") or "").strip(),
                    size=str(row.get("Размер") or "").strip(),
                    unit_label=str(row.get("Единица учёта") or "шт.").strip(),
                    sellable=bool(row.get("Продаётся отдельно")),
                    purchase_usd_per_unit=(
                        _as_float(row.get("Закупка USD/ед."))
                        if row.get("Закупка USD/ед.") not in (None, "")
                        else None
                    ),
                    raw=row,
                )
            )
        result.sort(key=lambda item: (not item.active, item.balance <= 0, item.sku.casefold()))
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
                "Статус": "Ожидается",
                "Комментарий": comment or "Создано в Analitika Web",
            },
        )

    @property
    def has_supply_lines(self) -> bool:
        return int(getattr(self.config, "supply_lines_table_id", 0) or 0) > 0

    def require_supply_lines(self) -> None:
        if not self.has_supply_lines:
            raise WarehouseServiceError(
                "Рабочие складские операции заблокированы до создания таблицы "
                "«Позиции поставок». Схема проверяется автоматически рабочим аккаунтом."
            )

    def next_supply_id(self, *, prefix: str = "SUP") -> str:
        today = date.today().strftime("%Y%m%d")
        base = f"{prefix}-{today}-"
        used: list[int] = []
        for row in self._supply_rows():
            value = str(row.get("№ поставки") or "").strip()
            if not value.startswith(base):
                continue
            try:
                used.append(int(value.rsplit("-", 1)[1]))
            except (TypeError, ValueError):
                continue
        return f"{base}{max(used, default=0) + 1:03d}"

    def next_silver_skus(self, count: int) -> list[str]:
        used: list[int] = []
        for row in self.catalog_rows("Комплектующие"):
            sku = str(row.get("Артикул") or "").strip().upper()
            if not sku.startswith("SIL"):
                continue
            try:
                used.append(int(sku[3:]))
            except (TypeError, ValueError):
                continue
        start = max(used, default=0) + 1
        return [f"SIL{value:06d}" for value in range(start, start + max(int(count), 0))]

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
        souvenir_catalog = {int(row["id"]): row for row in self.catalog_rows("Сувенирка")}
        component_catalog = {int(row["id"]): row for row in self.catalog_rows("Комплектующие")}
        if self.has_supply_lines:
            result: list[dict[str, Any]] = []
            for line in self._supply_line_rows():
                if supply_row_id not in link_ids(line.get("Поставка")):
                    continue
                souvenir_id = self._linked_row_id(line, ("Товар сувенирки", "Товар"))
                component_id = self._linked_row_id(line, ("Комплектующее",))
                product = souvenir_catalog.get(souvenir_id) or component_catalog.get(component_id) or {}
                section = "Комплектующие" if component_id else "Сувенирка"
                result.append(
                    {
                        **product,
                        "_section": section,
                        "_line_id": int(line["id"]),
                        "_supply_row_id": supply_row_id,
                        "_document": as_int(line.get("По документу, шт.")),
                        "_received": as_int(line.get("Принято, шт.")),
                        "_transferred": as_int(line.get("Передано в бухгалтерию, шт.")),
                        "_boxes": str(line.get("Номера коробок") or ""),
                        "_line_status": select_text(line.get("Статус")),
                        "_silver_925": bool(line.get("Серебро 925")),
                        "_original_name": str(line.get("Оригинальное название") or ""),
                        "_line_name": str(line.get("Название") or ""),
                        "_silver_category": str(line.get("Серебряная категория") or ""),
                        "_plating": str(line.get("Покрытие") or ""),
                        "_size": str(line.get("Размер") or ""),
                        "_unit_label": str(line.get("Единица учёта") or "шт."),
                        "_total_weight_g": _as_float(line.get("Вес партии, г")),
                        "_unit_weight_g": _as_float(line.get("Вес единицы, г")),
                        "_receiving_method": str(line.get("Способ приёмки") or ""),
                        "_receiving_weight_g": _as_float(line.get("Вес при приёмке, г")),
                        "_weight_estimated_qty": as_int(line.get("Расчётное количество по весу")),
                        "_weight_error_g": _as_float(line.get("Погрешность веса, г")),
                        "_silver_rmb_per_g": _as_float(line.get("Серебро RMB/г")),
                        "_labour_rmb_per_g": _as_float(line.get("Работа RMB/г")),
                        "_price_rmb_per_g": _as_float(line.get("Цена RMB/г")),
                        "_amount_rmb": _as_float(line.get("Сумма RMB")),
                        "_usd_rmb_rate": _as_float(line.get("Курс USD/RMB")),
                        "_cif_percent": _as_float(line.get("CIF, %")),
                        "_purchase_usd": _as_float(line.get("Закупка USD/ед.")),
                        "_invoice_sale_usd": _as_float(line.get("Продажа USD при импорте")),
                        "_invoice_usd_vnd": as_int(line.get("Курс USD/VND при импорте")),
                        "_invoice_coefficient": _as_float(line.get("Коэффициент при импорте")),
                        "_invoice_sale_vnd": as_int(line.get("Продажа VND при импорте")),
                        "_sellable": bool(line.get("Продаётся отдельно")),
                    }
                )
            return sorted(result, key=lambda row: str(row.get("Артикул") or ""))

        result = []
        for row in souvenir_catalog.values():
            if supply_row_id in link_ids(row.get("Поставки")):
                result.append(
                    {
                        **row,
                        "_section": "Сувенирка",
                        "_line_id": 0,
                        "_supply_row_id": supply_row_id,
                        "_document": as_int(row.get("По документу, шт.")),
                        "_received": as_int(row.get("Получено по поставке, шт.")),
                        "_transferred": 0,
                        "_boxes": str(row.get("Номера коробок") or ""),
                        "_line_status": "Старая схема",
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
        section: str,
        supply_row_id: int,
        existing: dict[str, Any] | None,
        photo: dict[str, Any] | None,
    ) -> dict[str, Any]:
        supply_links = list(
            dict.fromkeys(
                (link_ids(existing.get("Поставки")) if existing else []) + [supply_row_id]
            )
        )
        active_field = "Активный SKU" if section == "Сувенирка" else "Активно"
        payload: dict[str, Any] = {
            "Артикул": product.sku,
            "Материал": split_multi_values(product.material),
            "Камень": split_multi_values(normalize_stone_names(product.stone)),
            "Цвет": split_multi_values(product.color),
            "Вес 1 шт. (кг)": product.unit_weight_kg,
            "Поставки": supply_links,
            "Минимальный остаток": 10,
            active_field: True,
            "Комментарий": product.comment,
        }
        # Silver 925 uses its own free-text classification field. The legacy
        # Baserow field «Категория» is a select whose options differ between
        # installations; sending the importer-only value «Аксессуары» makes
        # Baserow reject the whole supply. Keep the legacy category untouched
        # for silver rows and store the real grouping in «Серебряная категория».
        if not (section == "Комплектующие" and product.silver_925):
            payload["Категория"] = product.category or None

        if section == "Комплектующие" and product.silver_925:
            payload.update(
                {
                    "Название": product.name or product.description,
                    "Серебряная категория": product.silver_category,
                    "Серебро 925": True,
                    "Покрытие": product.plating,
                    "Размер": product.size,
                    "Единица учёта": product.unit_label or "шт.",
                    "Продаётся отдельно": bool(product.sellable),
                    "Закупка USD/ед.": product.purchase_usd_per_unit,
                }
            )
        # Boxes and supply quantities belong to the supply line, not to the
        # permanent product card. Existing fields are intentionally left
        # untouched once the safe schema is active.
        if not self.has_supply_lines:
            payload["Номера коробок"] = product.boxes
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
        section: str = "Сувенирка",
        command_id: str = "",
    ) -> dict[str, Any]:
        self.require_supply_lines()
        products = [product for product in products if product.sku and product.qty_document > 0]
        if not products:
            raise WarehouseServiceError("В поставке нет корректных строк с SKU и количеством.")
        normalized = [product.sku.casefold() for product in products]
        if len(set(normalized)) != len(products):
            duplicates = sorted({sku for sku in normalized if normalized.count(sku) > 1})
            raise WarehouseServiceError(
                "В файле есть повторяющиеся SKU после нормализации: " + ", ".join(duplicates[:20])
            )
        if section not in {"Сувенирка", "Комплектующие"}:
            raise WarehouseServiceError("Неизвестный тип поставки.")

        supply_id = str(supply_id or "").strip()
        command_id = str(command_id or self.client.batch_id("IMPORT")).strip()
        existing_supply = self.find_supply(supply_id)
        if existing_supply:
            existing_products = self.supply_products(int(existing_supply["id"]))
            existing_operations = [
                row
                for row in self.client.list_rows(self.config.operations_table_id)
                if str(row.get("ID поставки") or "").strip() == supply_id
            ]
            if existing_products or existing_operations:
                raise WarehouseServiceError(
                    f"Поставка {supply_id} уже содержит позиции или операции. "
                    "Повторное создание заблокировано."
                )

        supply = self.find_or_create_supply(
            supply_id,
            supplier=supplier,
            invoice=invoice,
            comment=comment,
        )
        supply_row_id = int(supply["id"])
        self.client.batch_update(
            self.config.supplies_table_id,
            [{"id": supply_row_id, "Import ID": command_id, "Статус импорта": "В процессе"}],
        )

        table_id = self.table_id(section)
        rows = self.client.list_rows(table_id)
        by_sku = {
            str(row.get("Артикул") or "").strip().casefold(): row
            for row in rows
            if str(row.get("Артикул") or "").strip()
        }

        batch_id = self.client.batch_id("SUP")
        operations: list[dict[str, Any]] = []
        line_payloads: list[dict[str, Any]] = []
        operation_product_indexes: list[int] = []
        created = 0
        updated = 0
        photos = 0
        failed_photos: list[str] = []
        received = 0

        try:
            for product_index, product in enumerate(products):
                existing = by_sku.get(product.sku.casefold())
                photo = None
                if product.image_path and Path(product.image_path).exists():
                    try:
                        photo = self.client.upload_file(Path(product.image_path))
                        photos += 1
                    except WarehouseClientError:
                        failed_photos.append(product.sku)

                payload = self._product_payload(
                    product,
                    section=section,
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
                    by_sku[product.sku.casefold()] = {**row, **payload}
                    created += 1

                actual = int(product.actual_qty or 0)
                product_link_field = "Товар сувенирки" if section == "Сувенирка" else "Комплектующее"
                line_payloads.append(
                    {
                        "Строка поставки": f"{supply_id} — {product.sku}",
                        "Поставка": [supply_row_id],
                        product_link_field: [row_id],
                        "По документу, шт.": product.qty_document,
                        "Принято, шт.": actual,
                        "Передано в бухгалтерию, шт.": 0,
                        "Номера коробок": product.boxes,
                        "Статус": (
                            "Получена полностью"
                            if actual >= product.qty_document
                            else "Частично получена"
                            if actual
                            else "Ожидается"
                        ),
                        "Комментарий": product.comment,
                        "Версия": 1,
                        "Активна": True,
                        "Command ID": command_id,
                        "Создано из импорта": command_id,
                        **(
                            {
                                "Оригинальное название": product.original_name,
                                "Название": product.name or product.description,
                                "Серебряная категория": product.silver_category,
                                "Покрытие": product.plating,
                                "Размер": product.size,
                                "Единица учёта": product.unit_label or "шт.",
                                "Вес партии, г": product.total_weight_g,
                                "Вес единицы, г": (
                                    product.total_weight_g / product.qty_document
                                    if product.total_weight_g is not None and product.qty_document
                                    else None
                                ),
                                "Серебро RMB/г": product.silver_rmb_per_g,
                                "Работа RMB/г": product.labour_rmb_per_g,
                                "Цена RMB/г": product.price_rmb_per_g,
                                "Сумма RMB": product.amount_rmb,
                                "Курс USD/RMB": product.usd_rmb_rate,
                                "CIF, %": product.cif_percent,
                                "Закупка USD/ед.": product.purchase_usd_per_unit,
                                "Продажа USD при импорте": product.invoice_sale_usd,
                                "Курс USD/VND при импорте": product.invoice_usd_vnd_rate,
                                "Коэффициент при импорте": product.invoice_coefficient,
                                "Продажа VND при импорте": product.invoice_sale_vnd,
                                "Серебро 925": True,
                                "Продаётся отдельно": bool(product.sellable),
                            }
                            if product.silver_925
                            else {}
                        ),
                    }
                )
                if actual > 0:
                    operations.append(
                        {
                            "Операция": f"{batch_id} — {product.sku}",
                            "Тип операции": "Приход",
                            "Раздел": section,
                            product_link_field: [row_id],
                            "Количество": actual,
                            "Поставка": [supply_row_id],
                            "ID поставки": supply_id,
                            "Batch ID": batch_id,
                            "Command ID": command_id,
                            "Комментарий": "Импорт поставки из Analitika Web 2.5.8",
                        }
                    )
                    operation_product_indexes.append(product_index)
                    received += actual

            created_lines = self.client.batch_create(
                int(self.config.supply_lines_table_id),
                line_payloads,
            )
            line_ids_by_index = {
                index: int(row.get("id") or 0)
                for index, row in enumerate(created_lines)
            }
            for operation_index, product_index in enumerate(operation_product_indexes):
                line_id = line_ids_by_index.get(product_index, 0)
                if line_id:
                    operations[operation_index]["Позиция поставки"] = [line_id]

            created_operations = []
            if operations:
                created_operations = self.client.create_operations(
                    operations,
                    batch_id=batch_id,
                    command_id=command_id,
                )

            complete = all((product.actual_qty or 0) >= product.qty_document for product in products)
            supply_status = (
                "Получена полностью"
                if complete
                else "Частично получена"
                if received > 0
                else "Ожидается"
            )
            self.client.batch_update(
                self.config.supplies_table_id,
                [
                    {
                        "id": supply_row_id,
                        "Статус": supply_status,
                        "Статус импорта": "Завершён",
                    }
                ],
            )
            if created_operations:
                self.client.mark_operations_status(created_operations, "Проведена")
        except Exception:
            try:
                self.client.batch_update(
                    self.config.supplies_table_id,
                    [{"id": supply_row_id, "Статус импорта": "Ошибка"}],
                )
            except Exception:
                pass
            raise

        return {
            "supply_id": supply_id,
            "batch_id": batch_id if operations else "",
            "command_id": command_id,
            "sku": len(products),
            "created": created,
            "updated": updated,
            "photos": photos,
            "failed_photos": failed_photos,
            "received": received,
            "waiting": sum(product.waiting_qty for product in products),
            "section": section,
        }

    def _finalize_document(
        self,
        created_operations: list[dict[str, Any]],
        *,
        line_updates: list[dict[str, Any]] | None = None,
        supply_updates: list[dict[str, Any]] | None = None,
    ) -> None:
        try:
            if line_updates:
                self.client.batch_update(int(self.config.supply_lines_table_id), line_updates)
            if supply_updates:
                self.client.batch_update(self.config.supplies_table_id, supply_updates)
        except Exception:
            self.client.mark_operations_status(created_operations, "Требует восстановления")
            raise
        self.client.mark_operations_status(created_operations, "Проведена")

    def receive_supply(
        self,
        supply: SupplySummary,
        quantities: dict[int, int],
        *,
        command_id: str = "",
    ) -> dict[str, Any]:
        self.require_supply_lines()
        rows = self.supply_products(supply)

        def requested_quantity(row: dict[str, Any]) -> int:
            line_id = as_int(row.get("_line_id"))
            product_id = int(row["id"])
            return as_int(quantities.get(line_id, quantities.get(product_id, 0)))

        selected = [row for row in rows if requested_quantity(row) > 0]
        if not selected:
            raise WarehouseServiceError("Не указано количество для приёмки.")
        command_id = command_id or self.client.batch_id("CMD-REC")
        batch_id = self.client.batch_id("REC")
        operations: list[dict[str, Any]] = []
        line_updates: list[dict[str, Any]] = []
        total = 0
        for row in selected:
            row_id = int(row["id"])
            quantity = requested_quantity(row)
            waiting = max(as_int(row.get("_document")) - as_int(row.get("_received")), 0)
            if quantity > waiting:
                raise WarehouseServiceError(f"{row.get('Артикул')}: можно принять не более {waiting} шт.")
            section = str(row.get("_section") or "Сувенирка")
            link_field = "Комплектующее" if section == "Комплектующие" else "Товар сувенирки"
            line_id = as_int(row.get("_line_id"))
            operations.append({
                "Операция": f"{batch_id} — {row.get('Артикул')}",
                "Тип операции": "Приход",
                "Раздел": section,
                link_field: [row_id],
                "Количество": quantity,
                "Поставка": [supply.row_id],
                "Позиция поставки": [line_id] if line_id else [],
                "ID поставки": supply.supply_id,
                "Batch ID": batch_id,
                "Command ID": command_id,
                "Комментарий": "Доприёмка из Analitika Web 2.5.8",
            })
            new_received = as_int(row.get("_received")) + quantity
            line_updates.append({
                "id": line_id,
                "Принято, шт.": new_received,
                "Статус": "Получена полностью" if new_received >= as_int(row.get("_document")) else "Частично получена",
                "Способ приёмки": "По количеству",
            })
            total += quantity

        created_operations = self.client.create_operations(
            operations, batch_id=batch_id, command_id=command_id
        )
        projected = {as_int(row.get("_line_id")): as_int(row.get("_received")) for row in rows}
        for update in line_updates:
            projected[as_int(update["id"])] = as_int(update["Принято, шт."])
        complete = bool(rows) and all(
            projected.get(as_int(row.get("_line_id")), 0) >= as_int(row.get("_document"))
            for row in rows
        )
        self._finalize_document(
            created_operations,
            line_updates=line_updates,
            supply_updates=[{"id": supply.row_id, "Статус": "Получена полностью" if complete else "Частично получена"}],
        )
        return {"batch_id": batch_id, "command_id": command_id, "sku": len(operations), "quantity": total}

    def receive_existing_supply_by_weight(
        self,
        supply: SupplySummary,
        measurements: dict[int, dict[str, Any]],
        *,
        command_id: str = "",
    ) -> dict[str, Any]:
        """Register a fully delivered old supply and restore its current stock by weight.

        Each selected line is received for its full document quantity. The
        difference between the document quantity and the confirmed current
        quantity is posted as a separate expense named
        ``Использовано до постановки на учёт``. This preserves the true supplier
        delivery while making the live stock equal to the weighed remainder.
        """
        self.require_supply_lines()
        rows = self.supply_products(supply)
        by_line = {as_int(row.get("_line_id")): row for row in rows if as_int(row.get("_line_id"))}
        selected: list[tuple[dict[str, Any], float, int, int, float]] = []

        for raw_line_id, payload in measurements.items():
            line_id = as_int(raw_line_id)
            row = by_line.get(line_id)
            if row is None:
                continue
            if as_int(row.get("_received")) > 0:
                raise WarehouseServiceError(
                    f"{row.get('Артикул')}: приёмка по весу доступна только до первой приёмки этой строки."
                )
            document = as_int(row.get("_document"))
            unit_weight_g = _as_float(row.get("_unit_weight_g"))
            if document <= 0:
                raise WarehouseServiceError(f"{row.get('Артикул')}: количество по документу не указано.")
            if unit_weight_g <= 0:
                raise WarehouseServiceError(
                    f"{row.get('Артикул')}: нет среднего веса единицы, используйте приёмку по количеству."
                )
            weight_g = _as_float(payload.get("weight_g"), -1.0)
            if weight_g < 0:
                raise WarehouseServiceError(f"{row.get('Артикул')}: укажите неотрицательный чистый вес.")
            estimated = self.estimate_quantity_from_weight(weight_g, unit_weight_g, maximum=document)
            final_quantity = as_int(payload.get("quantity"), estimated)
            if final_quantity < 0 or final_quantity > document:
                raise WarehouseServiceError(
                    f"{row.get('Артикул')}: итоговый остаток должен быть от 0 до {document}."
                )
            selected.append((row, weight_g, final_quantity, estimated, unit_weight_g))

        if not selected:
            raise WarehouseServiceError("Не введён вес ни для одной позиции.")

        command_id = command_id or self.client.batch_id("CMD-RECW")
        batch_id = self.client.batch_id("RECW")
        operations: list[dict[str, Any]] = []
        line_updates: list[dict[str, Any]] = []
        received_total = 0
        current_total = 0
        written_off_total = 0

        for row, weight_g, final_quantity, estimated, unit_weight_g in selected:
            row_id = int(row["id"])
            line_id = as_int(row.get("_line_id"))
            document = as_int(row.get("_document"))
            section = str(row.get("_section") or "Сувенирка")
            link_field = "Комплектующее" if section == "Комплектующие" else "Товар сувенирки"
            unit_label = str(row.get("_unit_label") or "шт.")
            common = {
                "Раздел": section,
                link_field: [row_id],
                "Поставка": [supply.row_id],
                "Позиция поставки": [line_id],
                "ID поставки": supply.supply_id,
                "Batch ID": batch_id,
                "Command ID": command_id,
            }
            operations.append(
                {
                    **common,
                    "Операция": f"{batch_id} — {row.get('Артикул')} — полный приход",
                    "Тип операции": "Приход",
                    "Количество": document,
                    "Комментарий": (
                        "Поставка получена полностью; текущий остаток восстановлен по весу "
                        f"{weight_g:.4f} г при среднем весе {unit_weight_g:.6f} г/{unit_label}."
                    ),
                }
            )
            used = max(document - final_quantity, 0)
            if used:
                operations.append(
                    {
                        **common,
                        "Операция": f"{batch_id} — {row.get('Артикул')} — использовано до учёта",
                        "Тип операции": "Расход",
                        "Количество": used,
                        "Комментарий": (
                            "Использовано до постановки поставки на учёт. "
                            f"Чистый остаток: {weight_g:.4f} г; расчёт: {estimated} {unit_label}; "
                            f"подтверждено оператором: {final_quantity} {unit_label}."
                        ),
                    }
                )
            expected_weight = final_quantity * unit_weight_g
            line_updates.append(
                {
                    "id": line_id,
                    "Принято, шт.": document,
                    "Статус": "Получена полностью",
                    "Способ приёмки": "По весу — товар уже в работе",
                    "Вес при приёмке, г": weight_g,
                    "Расчётное количество по весу": estimated,
                    "Погрешность веса, г": abs(weight_g - expected_weight),
                }
            )
            received_total += document
            current_total += final_quantity
            written_off_total += used

        projected = {as_int(row.get("_line_id")): as_int(row.get("_received")) for row in rows}
        for update in line_updates:
            projected[as_int(update["id"])] = as_int(update["Принято, шт."])
        complete = bool(rows) and all(
            projected.get(as_int(row.get("_line_id")), 0) >= as_int(row.get("_document"))
            for row in rows
        )
        created_operations = self.client.create_operations(
            operations, batch_id=batch_id, command_id=command_id
        )
        self._finalize_document(
            created_operations,
            line_updates=line_updates,
            supply_updates=[
                {
                    "id": supply.row_id,
                    "Статус": "Получена полностью" if complete else "Частично получена",
                }
            ],
        )
        return {
            "batch_id": batch_id,
            "command_id": command_id,
            "sku": len(selected),
            "received": received_total,
            "current": current_total,
            "written_off": written_off_total,
        }

    def transfer_supply(
        self,
        supply: SupplySummary,
        quantities: dict[int, int],
        *,
        comment: str = "",
        command_id: str = "",
    ) -> dict[str, Any]:
        self.require_supply_lines()
        rows = self.supply_products(supply)
        selected: list[tuple[dict[str, Any], int]] = []
        for row in rows:
            row_id = int(row["id"])
            line_id = as_int(row.get("_line_id"))
            quantity = as_int(quantities.get(line_id, quantities.get(row_id, 0)))
            if quantity <= 0:
                continue
            received = as_int(row.get("_received"))
            transferred = as_int(row.get("_transferred"))
            stock = as_int(row.get("Остаток"))
            maximum = min(max(received - transferred, 0), max(stock, 0))
            if quantity > maximum:
                raise WarehouseServiceError(f"{row.get('Артикул')}: можно передать не более {maximum} шт.")
            selected.append((row, quantity))
        if not selected:
            raise WarehouseServiceError("Не выбраны товары для передачи.")

        command_id = command_id or self.client.batch_id("CMD-ACC")
        batch_id = self.client.batch_id("ACC")
        operations: list[dict[str, Any]] = []
        line_updates: list[dict[str, Any]] = []
        for row, quantity in selected:
            row_id = int(row["id"])
            line_id = as_int(row.get("_line_id"))
            section = str(row.get("_section") or "Сувенирка")
            link_field = "Комплектующее" if section == "Комплектующие" else "Товар сувенирки"
            operations.append({
                "Операция": f"{batch_id} — {row.get('Артикул')}",
                "Тип операции": "Передача в бухгалтерию",
                "Раздел": section,
                link_field: [row_id],
                "Количество": quantity,
                "Поставка": [supply.row_id],
                "Позиция поставки": [line_id] if line_id else [],
                "ID поставки": supply.supply_id,
                "Batch ID": batch_id,
                "Command ID": command_id,
                "Комментарий": comment or f"Поставка {supply.supply_id}",
            })
            line_updates.append({
                "id": line_id,
                "Передано в бухгалтерию, шт.": as_int(row.get("_transferred")) + quantity,
                "Статус": (
                    "Передана полностью"
                    if as_int(row.get("_transferred")) + quantity >= as_int(row.get("_received"))
                    else "Частично передана"
                ),
            })
        created_operations = self.client.create_operations(
            operations, batch_id=batch_id, command_id=command_id
        )
        self._finalize_document(created_operations, line_updates=line_updates)
        return {
            "batch_id": batch_id,
            "command_id": command_id,
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
        command_id: str = "",
    ) -> dict[str, Any]:
        catalog = {item.row_id: item for item in self.catalog(section)}
        selected = [
            (catalog[row_id], as_int(quantity))
            for row_id, quantity in quantities.items()
            if row_id in catalog and as_int(quantity) > 0
        ]
        if not selected:
            raise WarehouseServiceError("Не выбрано ни одной позиции.")
        if operation_type == "Приход" and not str(comment or "").strip():
            raise WarehouseServiceError("Для ручного прихода обязательно укажите основание и комментарий.")
        if operation_type == "Передача в бухгалтерию":
            for item, quantity in selected:
                if quantity > item.balance:
                    raise WarehouseServiceError(f"{item.sku}: требуется {quantity}, доступно {item.balance}.")
        prefix = "ACC" if operation_type == "Передача в бухгалтерию" else "REC"
        command_id = command_id or self.client.batch_id(f"CMD-{prefix}")
        batch_id = self.client.batch_id(prefix)
        link_field = "Товар сувенирки" if section == "Сувенирка" else "Комплектующее"
        operations = [{
            "Операция": f"{batch_id} — {item.sku}",
            "Тип операции": operation_type,
            "Раздел": section,
            link_field: [item.row_id],
            "Количество": quantity,
            "ID поставки": supply_id,
            "Batch ID": batch_id,
            "Command ID": command_id,
            "Комментарий": comment,
        } for item, quantity in selected]
        created = self.client.create_operations(operations, batch_id=batch_id, command_id=command_id)
        self.client.mark_operations_status(created, "Проведена")
        return {
            "batch_id": batch_id,
            "command_id": command_id,
            "sku": len(selected),
            "quantity": sum(quantity for _, quantity in selected),
        }

    def correction_available(self, operation: dict[str, Any]) -> int:
        original_id = as_int(operation.get("id"))
        original_quantity = as_int(operation.get("Количество"))
        corrected = 0
        for row in self.client.list_rows(self.config.operations_table_id):
            if original_id in link_ids(row.get("Исходная операция")):
                corrected += as_int(row.get("Количество"))
        return max(original_quantity - corrected, 0)

    def correct_operation(
        self,
        operation: dict[str, Any],
        *,
        quantity: int,
        comment: str,
        command_id: str = "",
    ) -> dict[str, Any]:
        self.require_supply_lines()
        quantity = as_int(quantity)
        available = self.correction_available(operation)
        if quantity <= 0 or quantity > available:
            raise WarehouseServiceError(f"Доступно к корректировке: {available} шт.")
        original_type = select_text(operation.get("Тип операции"))
        original_is_incoming = "приход" in original_type.casefold()
        reverse_type = "Расход" if original_is_incoming else "Возврат"
        command_id = command_id or self.client.batch_id("CMD-COR")
        batch_id = self.client.batch_id("COR")
        payload = {
            "Операция": f"{batch_id} — отмена {operation.get('Операция') or operation.get('id')}",
            "Тип операции": reverse_type,
            "Раздел": select_text(operation.get("Раздел")),
            "Товар сувенирки": link_ids(operation.get("Товар сувенирки")),
            "Комплектующее": link_ids(operation.get("Комплектующее")),
            "Количество": quantity,
            "Поставка": link_ids(operation.get("Поставка")),
            "Позиция поставки": link_ids(operation.get("Позиция поставки")),
            "ID поставки": str(operation.get("ID поставки") or ""),
            "Batch ID": batch_id,
            "Command ID": command_id,
            "Исходная операция": [as_int(operation.get("id"))],
            "Причина корректировки": comment,
            "Комментарий": comment or f"Корректировка операции {operation.get('Batch ID') or operation.get('id')}",
        }
        created = self.client.create_operations([payload], batch_id=batch_id, command_id=command_id)

        line_updates: list[dict[str, Any]] = []
        for line_id in link_ids(operation.get("Позиция поставки")):
            line = next((row for row in self._supply_line_rows() if int(row["id"]) == line_id), None)
            if not line:
                continue
            if original_is_incoming:
                updated = max(as_int(line.get("Принято, шт.")) - quantity, 0)
                document = as_int(line.get("По документу, шт."))
                status = "Получена полностью" if updated >= document and document else "Частично получена" if updated else "Ожидается"
                line_updates.append({"id": line_id, "Принято, шт.": updated, "Статус": status})
            else:
                updated = max(as_int(line.get("Передано в бухгалтерию, шт.")) - quantity, 0)
                line_updates.append({"id": line_id, "Передано в бухгалтерию, шт.": updated})
        self._finalize_document(created, line_updates=line_updates)
        return {"batch_id": batch_id, "command_id": command_id, "quantity": quantity, "remaining": available - quantity}

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
        has_supply_lines = self.has_supply_lines and any(
            int(row_id) in link_ids(row.get(operation_link))
            for row in self._supply_line_rows()
        )
        if has_operations or has_supply_lines:
            field = "Активный SKU" if section == "Сувенирка" else "Активно"
            self.client.batch_update(self.table_id(section), [{"id": int(row_id), field: False}])
            return "deactivated"
        self.client.delete_row(self.table_id(section), int(row_id))
        return "deleted"

    @staticmethod
    def _operation_is_posted(operation: dict[str, Any]) -> bool:
        status = select_text(operation.get("Статус документа")).strip().casefold()
        return status not in {"создаётся", "требует восстановления", "ошибка", "отменена"}

    @staticmethod
    def _operation_effect(operation: dict[str, Any]) -> tuple[int, int]:
        operation_type = select_text(operation.get("Тип операции")).strip().casefold()
        if operation_type == "приход":
            return 1, 0
        if operation_type == "расход":
            return -1, 0
        if "передач" in operation_type:
            return 0, 1
        if operation_type == "возврат":
            return 0, -1
        return 0, 0

    @staticmethod
    def _operation_product_refs(operation: dict[str, Any]) -> list[tuple[str, int]]:
        refs: list[tuple[str, int]] = []
        refs.extend(("Сувенирка", row_id) for row_id in link_ids(operation.get("Товар сувенирки")))
        refs.extend(("Комплектующие", row_id) for row_id in link_ids(operation.get("Комплектующее")))
        return refs

    def _catalog_operation_refs(self, operations: Iterable[dict[str, Any]]) -> set[tuple[str, int]]:
        result: set[tuple[str, int]] = set()
        for operation in operations:
            if not self._operation_is_posted(operation):
                continue
            result.update(self._operation_product_refs(operation))
        return result

    def synchronize_baserow_from_documents(self) -> dict[str, int]:
        """Reconcile Baserow counters and links from current supplies and posted operations.

        The method does not create receipt operations or alter document quantities. It only
        repairs derived Baserow fields, supply statuses and stale catalog links. Catalog rows
        left behind by a deleted, never-received supply line are deleted only when they have
        zero stock and no posted operation history.
        """
        self.require_supply_lines()
        supply_lines = self.client.list_rows(int(self.config.supply_lines_table_id), refresh=True)
        operations = self.client.list_rows(self.config.operations_table_id, refresh=True)
        supplies = self.client.list_rows(self.config.supplies_table_id, refresh=True)
        souvenir_rows = self.client.list_rows(self.config.souvenirs_table_id, refresh=True)
        component_rows = self.client.list_rows(self.config.components_table_id, refresh=True)

        active_lines = [row for row in supply_lines if row.get("Активна") is not False]
        line_by_id = {int(row["id"]): row for row in active_lines}
        supply_id_by_number = {
            str(row.get("№ поставки") or "").strip(): int(row["id"])
            for row in supplies
            if str(row.get("№ поставки") or "").strip()
        }

        direct_totals: dict[int, list[int]] = defaultdict(lambda: [0, 0])
        fallback_totals: dict[tuple[int, str, int], list[int]] = defaultdict(lambda: [0, 0])
        operation_refs: set[tuple[str, int]] = set()

        for operation in operations:
            if not self._operation_is_posted(operation):
                continue
            incoming_direction, transfer_direction = self._operation_effect(operation)
            if not incoming_direction and not transfer_direction:
                continue
            quantity = max(as_int(operation.get("Количество")), 0)
            if quantity <= 0:
                continue
            product_refs = self._operation_product_refs(operation)
            operation_refs.update(product_refs)
            line_ids = [line_id for line_id in link_ids(operation.get("Позиция поставки")) if line_id in line_by_id]
            if line_ids:
                for line_id in line_ids:
                    direct_totals[line_id][0] += incoming_direction * quantity
                    direct_totals[line_id][1] += transfer_direction * quantity
                continue

            supply_ids = link_ids(operation.get("Поставка"))
            if not supply_ids:
                supply_number = str(operation.get("ID поставки") or "").strip()
                supply_row_id = supply_id_by_number.get(supply_number, 0)
                supply_ids = [supply_row_id] if supply_row_id else []
            for supply_row_id in supply_ids:
                for section, product_id in product_refs:
                    fallback_totals[(supply_row_id, section, product_id)][0] += incoming_direction * quantity
                    fallback_totals[(supply_row_id, section, product_id)][1] += transfer_direction * quantity

        line_keys: dict[tuple[int, str, int], list[int]] = defaultdict(list)
        for line in active_lines:
            supply_row_id = self._linked_row_id(line, ("Поставка",))
            souvenir_id = self._linked_row_id(line, ("Товар сувенирки", "Товар"))
            component_id = self._linked_row_id(line, ("Комплектующее",))
            section = "Комплектующие" if component_id else "Сувенирка"
            product_id = component_id or souvenir_id
            if supply_row_id and product_id:
                line_keys[(supply_row_id, section, product_id)].append(int(line["id"]))

        projected_lines: dict[int, dict[str, int | str]] = {}
        line_updates: list[dict[str, Any]] = []
        for line in active_lines:
            line_id = int(line["id"])
            document = max(as_int(line.get("По документу, шт.")), 0)
            received = max(as_int(line.get("Принято, шт.")), 0)
            transferred = max(as_int(line.get("Передано в бухгалтерию, шт.")), 0)
            evidence: list[int] | None = direct_totals.get(line_id)

            if evidence is None:
                supply_row_id = self._linked_row_id(line, ("Поставка",))
                souvenir_id = self._linked_row_id(line, ("Товар сувенирки", "Товар"))
                component_id = self._linked_row_id(line, ("Комплектующее",))
                section = "Комплектующие" if component_id else "Сувенирка"
                product_id = component_id or souvenir_id
                key = (supply_row_id, section, product_id)
                if len(line_keys.get(key, [])) == 1 and key in fallback_totals:
                    evidence = fallback_totals[key]

            if evidence is not None:
                received = max(int(evidence[0]), 0)
                transferred = max(int(evidence[1]), 0)
            transferred = min(transferred, received)

            if received > 0 and transferred >= received:
                status = "Передана полностью"
            elif transferred > 0:
                status = "Частично передана"
            elif document > 0 and received >= document:
                status = "Получена полностью"
            elif received > 0:
                status = "Частично получена"
            else:
                status = "Ожидается"

            projected_lines[line_id] = {
                "document": document,
                "received": received,
                "transferred": transferred,
                "status": status,
            }
            if (
                received != as_int(line.get("Принято, шт."))
                or transferred != as_int(line.get("Передано в бухгалтерию, шт."))
                or status != select_text(line.get("Статус"))
            ):
                line_updates.append(
                    {
                        "id": line_id,
                        "Принято, шт.": received,
                        "Передано в бухгалтерию, шт.": transferred,
                        "Статус": status,
                    }
                )

        if line_updates:
            self.client.batch_update(int(self.config.supply_lines_table_id), line_updates)

        lines_by_supply: dict[int, list[dict[str, int | str]]] = defaultdict(list)
        for line in active_lines:
            supply_row_id = self._linked_row_id(line, ("Поставка",))
            if supply_row_id:
                lines_by_supply[supply_row_id].append(projected_lines[int(line["id"])])

        supply_updates: list[dict[str, Any]] = []
        for supply in supplies:
            rows = lines_by_supply.get(int(supply["id"]), [])
            if not rows:
                continue
            if all(int(row["document"]) > 0 and int(row["received"]) >= int(row["document"]) for row in rows):
                status = "Получена полностью"
            elif any(int(row["received"]) > 0 for row in rows):
                status = "Частично получена"
            else:
                status = "Ожидается"
            if status != select_text(supply.get("Статус")):
                supply_updates.append({"id": int(supply["id"]), "Статус": status})
        if supply_updates:
            self.client.batch_update(self.config.supplies_table_id, supply_updates)

        valid_links: dict[tuple[str, int], set[int]] = defaultdict(set)
        for line in active_lines:
            supply_row_id = self._linked_row_id(line, ("Поставка",))
            souvenir_id = self._linked_row_id(line, ("Товар сувенирки", "Товар"))
            component_id = self._linked_row_id(line, ("Комплектующее",))
            if supply_row_id and souvenir_id:
                valid_links[("Сувенирка", souvenir_id)].add(supply_row_id)
            if supply_row_id and component_id:
                valid_links[("Комплектующие", component_id)].add(supply_row_id)

        catalog_updates: dict[int, list[dict[str, Any]]] = defaultdict(list)
        deleted = 0
        deactivated = 0
        relinked = 0
        for section, table_id, active_field, rows in (
            ("Сувенирка", self.config.souvenirs_table_id, "Активный SKU", souvenir_rows),
            ("Комплектующие", self.config.components_table_id, "Активно", component_rows),
        ):
            for row in rows:
                row_id = int(row["id"])
                current = set(link_ids(row.get("Поставки")))
                expected = valid_links.get((section, row_id), set())
                stale = current - expected
                if not stale and current == expected:
                    continue
                balance = as_int(row.get("Остаток"))
                has_history = (section, row_id) in operation_refs
                if stale and not expected and balance <= 0 and not has_history:
                    self.client.delete_row(int(table_id), row_id)
                    deleted += 1
                    continue
                payload: dict[str, Any] = {"id": row_id, "Поставки": sorted(expected)}
                if stale and not expected and balance <= 0 and has_history:
                    payload[active_field] = False
                    deactivated += 1
                catalog_updates[int(table_id)].append(payload)
                relinked += 1

        for table_id, updates in catalog_updates.items():
            if updates:
                self.client.batch_update(table_id, updates)

        return {
            "lines_updated": len(line_updates),
            "supplies_updated": len(supply_updates),
            "catalog_relinked": relinked,
            "catalog_deleted": deleted,
            "catalog_deactivated": deactivated,
        }

    def remove_waiting_from_supply(self, supply: SupplySummary, product_ids: list[int]) -> int:
        selected = [row for row in self.supply_products(supply) if int(row["id"]) in product_ids]
        blocked = [str(row.get("Артикул") or "") for row in selected if as_int(row.get("_received")) > 0]
        if blocked:
            raise WarehouseServiceError(
                "Нельзя убрать уже принятые позиции: " + ", ".join(blocked[:12])
            )
        if self.has_supply_lines:
            deleted_line_ids = {
                as_int(row.get("_line_id"))
                for row in selected
                if as_int(row.get("_line_id"))
            }
            for line_id in deleted_line_ids:
                self.client.delete_row(int(self.config.supply_lines_table_id), line_id)

            remaining_lines = [
                line
                for line in self.client.list_rows(int(self.config.supply_lines_table_id), refresh=True)
                if int(line.get("id") or 0) not in deleted_line_ids and line.get("Активна") is not False
            ]
            operations = self.client.list_rows(self.config.operations_table_id, refresh=True)
            operation_refs = self._catalog_operation_refs(operations)

            affected: dict[tuple[str, int], dict[str, Any]] = {}
            for row in selected:
                affected[(str(row.get("_section") or "Сувенирка"), int(row["id"]))] = row

            for (section, product_id), row in affected.items():
                table_id = self.table_id(section)
                active_field = "Активный SKU" if section == "Сувенирка" else "Активно"
                valid_supply_links: set[int] = set()
                for line in remaining_lines:
                    linked_product_id = self._linked_row_id(
                        line,
                        ("Комплектующее",) if section == "Комплектующие" else ("Товар сувенирки", "Товар"),
                    )
                    if linked_product_id != product_id:
                        continue
                    supply_row_id = self._linked_row_id(line, ("Поставка",))
                    if supply_row_id:
                        valid_supply_links.add(supply_row_id)

                has_history = (section, product_id) in operation_refs
                balance = as_int(row.get("Остаток"))
                if not valid_supply_links and balance <= 0 and not has_history:
                    self.client.delete_row(table_id, product_id)
                    continue

                payload: dict[str, Any] = {
                    "id": product_id,
                    "Поставки": sorted(valid_supply_links),
                }
                if not valid_supply_links and balance <= 0 and has_history:
                    payload[active_field] = False
                self.client.batch_update(table_id, [payload])
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
        souvenirs = self.catalog("Сувенирка", include_inactive=True)
        components = self.catalog("Комплектующие", include_inactive=True)
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
            "inactive": [item.sku for item in [*souvenirs, *components] if not item.active],
            "unfinished_operations": [
                str(row.get("Batch ID") or row.get("id"))
                for row in operations
                if select_text(row.get("Статус документа")) in {"Создаётся", "Требует восстановления"}
            ],
            "ambiguous_supply_lines": [
                str(row.get("Строка поставки") or row.get("id"))
                for row in self._supply_line_rows()
                if select_text(row.get("Статус")) == "Требует проверки"
            ] if self.has_supply_lines else [],
        }
