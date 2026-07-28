from __future__ import annotations

from dataclasses import asdict
from io import BytesIO
from pathlib import Path
import shutil
import tempfile
from typing import Any, Iterable
import uuid

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

from .client import WarehouseClient, WarehouseClientError, as_int, link_ids, select_text
from .models import Product, SupplySummary
from .packing import CATEGORIES, export_master, load_products
from .service import WarehouseService, WarehouseServiceError


WORKSPACES = (
    "Обзор",
    "Каталог",
    "Поставки",
    "Новая поставка",
    "Приёмка",
    "Передача в бухгалтерию",
    "Операции",
    "Обслуживание",
)

WAREHOUSE_MANAGEMENT_CSS = """
<style>
.wm-context {border:1px solid rgba(183,137,63,.25);border-radius:14px;padding:12px 15px;
 background:linear-gradient(90deg,#fffaf1,#fff);margin:.35rem 0 1rem;color:#5e5140}
.wm-good {border-left:4px solid #3a7d51;background:#f4fbf6;padding:11px 13px;border-radius:9px}
.wm-warning {border-left:4px solid #b7893f;background:#fffaf1;padding:11px 13px;border-radius:9px}
.wm-danger {border-left:4px solid #aa3939;background:#fff6f6;padding:11px 13px;border-radius:9px}
.wm-title {font-family:Georgia,serif;font-size:30px;color:#171411;margin:.3rem 0 .2rem}
@media (max-width:640px){.wm-title{font-size:25px}.wm-context{padding:10px 12px}}
</style>
"""


def _clear_cache() -> None:
    st.cache_data.clear()


def _service(config: Any) -> WarehouseService:
    return WarehouseService(WarehouseClient(config))


def _safe_action(action, *, success: str | None = None) -> Any:
    try:
        result = action()
    except (WarehouseClientError, WarehouseServiceError, ValueError, OSError) as exc:
        st.error(str(exc))
        return None
    _clear_cache()
    if success:
        st.success(success)
    # Preserve a truthy success signal for service methods that intentionally
    # return None after completing a write operation.
    return True if result is None else result


def _summary_options(summaries: list[SupplySummary]) -> tuple[list[str], dict[str, SupplySummary]]:
    mapping = {
        f"{item.supply_id} · {item.status or 'Без статуса'} · принято {item.qty_received:,} · ожидается {item.qty_waiting:,}": item
        for item in summaries
    }
    return list(mapping), mapping


def _item_photo_name(item: Any) -> str:
    photo = getattr(item, "photo", None)
    if isinstance(photo, list) and photo and isinstance(photo[0], dict):
        return str(photo[0].get("name") or photo[0].get("url") or "")
    return ""


def _catalog_dataframe(items: list[Any]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Артикул": item.sku,
                "Раздел": item.section,
                "Остаток": item.balance,
                "Минимум": item.min_balance,
                "Категория": item.category,
                "Материал": item.material,
                "Камень": item.stone,
                "Цвет": item.color,
                "Коробки": item.boxes,
                "Фото": _item_photo_name(item),
                "row_id": item.row_id,
            }
            for item in items
        ]
    )


def render_overview(config: Any, selected_metal_groups: Iterable[str]) -> None:
    from src.warehouse import filter_warehouse_bundle, load_bundle, render_attention, render_overview

    with st.spinner("Загружаем актуальный склад..."):
        bundle = load_bundle(config)
    selected = tuple(str(value) for value in selected_metal_groups)
    if selected:
        bundle = filter_warehouse_bundle(bundle, selected)
    render_overview(bundle)
    with st.expander("Позиции, требующие внимания", expanded=False):
        render_attention(bundle)


def render_catalog(config: Any) -> None:
    st.markdown('<div class="wm-title">Каталог</div>', unsafe_allow_html=True)
    service = _service(config)
    section = st.segmented_control(
        "Раздел",
        ["Сувенирка", "Комплектующие"],
        default="Сувенирка",
        key="warehouse_catalog_manage_section",
    ) or "Сувенирка"
    with st.spinner("Читаем каталог Baserow..."):
        items = service.catalog(section)

    view_tab, add_tab, edit_tab, delete_tab = st.tabs(
        ["Просмотр", "Добавить", "Редактировать", "Удалить / деактивировать"]
    )
    with view_tab:
        query = st.text_input(
            "Поиск",
            placeholder="Артикул, категория, материал, камень, коробка",
            key=f"warehouse_catalog_search_{section}",
        ).strip().casefold()
        status = st.segmented_control(
            "Остаток",
            ["Все", "Есть", "Мало", "Нет"],
            default="Все",
            key=f"warehouse_catalog_status_{section}",
        ) or "Все"
        filtered = []
        for item in items:
            if query and query not in item.search_text.casefold():
                continue
            if status == "Есть" and item.balance <= 0:
                continue
            if status == "Мало" and not (0 < item.balance <= 15):
                continue
            if status == "Нет" and item.balance > 0:
                continue
            filtered.append(item)
        st.caption(f"Найдено: {len(filtered)} SKU")
        st.dataframe(
            _catalog_dataframe(filtered).drop(columns=["row_id"], errors="ignore"),
            width="stretch",
            hide_index=True,
            height=620,
        )

    with add_tab:
        with st.form(f"warehouse_add_catalog_{section}", clear_on_submit=True):
            c1, c2 = st.columns(2)
            sku = c1.text_input("Артикул *")
            category = c2.text_input("Категория")
            material = c1.text_input("Материал", placeholder="Steel; Brass")
            stone = c2.text_input("Камни", placeholder="Agate; Pearl")
            color = c1.text_input("Цвет")
            boxes = c2.text_input("Коробки")
            minimum = c1.number_input("Минимальный остаток", min_value=1, value=10, step=1)
            photo = c2.file_uploader(
                "Фото",
                type=["jpg", "jpeg", "png", "webp"],
                key=f"warehouse_add_photo_{section}",
            )
            comment = st.text_area("Комментарий")
            submitted = st.form_submit_button("Создать карточку", type="primary", width="stretch")
        if submitted:
            temp_path = None
            if photo is not None:
                suffix = Path(photo.name).suffix or ".jpg"
                temp_path = Path(tempfile.gettempdir()) / f"warehouse-photo-{uuid.uuid4().hex}{suffix}"
                temp_path.write_bytes(photo.getvalue())
            result = _safe_action(
                lambda: service.add_catalog_item(
                    section=section,
                    sku=sku,
                    category=category,
                    material=material,
                    stone=stone,
                    color=color,
                    boxes=boxes,
                    minimum=int(minimum),
                    comment=comment,
                    photo_path=temp_path,
                )
            )
            if temp_path:
                temp_path.unlink(missing_ok=True)
            if result is not None:
                st.success(f"Карточка {sku} создана.")
                st.rerun()

    with edit_tab:
        if not items:
            st.info("Каталог пуст.")
        else:
            labels = {f"{item.sku} · остаток {item.balance}": item for item in items}
            label = st.selectbox("Карточка", list(labels), key=f"warehouse_edit_item_{section}")
            item = labels[label]
            with st.form(f"warehouse_edit_form_{section}_{item.row_id}"):
                c1, c2 = st.columns(2)
                category = c1.text_input("Категория", value=item.category)
                material = c2.text_input("Материал", value=item.material)
                stone = c1.text_input("Камни", value=item.stone)
                color = c2.text_input("Цвет", value=item.color)
                boxes = c1.text_input("Коробки", value=item.boxes)
                minimum = c2.number_input(
                    "Минимальный остаток",
                    min_value=1,
                    value=max(int(item.min_balance), 1),
                )
                comment = st.text_area(
                    "Комментарий",
                    value=str((item.raw or {}).get("Комментарий") or ""),
                )
                saved = st.form_submit_button("Сохранить изменения", type="primary", width="stretch")
            if saved:
                result = _safe_action(
                    lambda: service.update_catalog_item(
                        section,
                        item.row_id,
                        {
                            "Категория": category or None,
                            "Материал": material,
                            "Камень": stone,
                            "Цвет": color,
                            "Номера коробок": boxes,
                            "Минимальный остаток": int(minimum),
                            "Комментарий": comment,
                        },
                    )
                )
                if result:
                    st.success(f"Карточка {item.sku} обновлена.")
                    st.rerun()

    with delete_tab:
        st.markdown(
            '<div class="wm-warning">Карточка с операциями не удаляется физически — '
            'она деактивируется, чтобы сохранить историю.</div>',
            unsafe_allow_html=True,
        )
        if items:
            labels = {f"{item.sku} · остаток {item.balance}": item for item in items}
            label = st.selectbox("Карточка", list(labels), key=f"warehouse_delete_item_{section}")
            item = labels[label]
            confirmation = st.text_input(
                f"Для подтверждения введите артикул {item.sku}",
                key=f"warehouse_delete_confirmation_{section}_{item.row_id}",
            )
            if st.button(
                "Удалить или деактивировать",
                type="primary",
                disabled=confirmation.strip() != item.sku,
                key=f"warehouse_delete_button_{section}_{item.row_id}",
            ):
                action = _safe_action(
                    lambda: service.deactivate_or_delete_catalog_item(section, item.row_id)
                )
                if action == "deleted":
                    st.success("Карточка удалена: по ней не было операций.")
                elif action == "deactivated":
                    st.success("Карточка деактивирована, история сохранена.")
                st.rerun()


def _supply_table(summaries: list[SupplySummary]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Поставка": item.supply_id,
                "Дата": item.date,
                "Поставщик": item.supplier,
                "Статус": item.status,
                "SKU": item.sku_total,
                "По документу": item.qty_document,
                "Принято": item.qty_received,
                "Ожидается": item.qty_waiting,
            }
            for item in summaries
        ]
    )


def render_supplies(config: Any) -> None:
    st.markdown('<div class="wm-title">Поставки</div>', unsafe_allow_html=True)
    service = _service(config)
    with st.spinner("Загружаем реестр поставок..."):
        summaries = service.supply_summaries()
    if not summaries:
        st.info("В Baserow пока нет связанных поставок.")
        return

    st.dataframe(_supply_table(summaries), width="stretch", hide_index=True, height=330)
    options, mapping = _summary_options(summaries)
    selected_label = st.selectbox("Открыть поставку", options, key="warehouse_supply_detail")
    supply = mapping[selected_label]
    rows = service.supply_products(supply)
    detail = pd.DataFrame(
        [
            {
                "Артикул": row.get("Артикул"),
                "Коробки": row.get("_boxes"),
                "По документу": as_int(row.get("_document")),
                "Принято": as_int(row.get("_received")),
                "Ожидается": max(as_int(row.get("_document")) - as_int(row.get("_received")), 0),
                "Остаток": as_int(row.get("Остаток")),
                "row_id": int(row["id"]),
            }
            for row in rows
        ]
    )
    st.dataframe(detail.drop(columns=["row_id"], errors="ignore"), width="stretch", hide_index=True, height=430)

    with st.expander("Исправление поставки", expanded=False):
        waiting = detail.loc[detail["Принято"] <= 0] if not detail.empty else detail
        remove_skus = st.multiselect(
            "Убрать непринятые позиции",
            waiting["Артикул"].tolist() if not waiting.empty else [],
            key=f"warehouse_remove_waiting_{supply.row_id}",
        )
        if st.button(
            "Убрать выбранные из поставки",
            disabled=not remove_skus,
            key=f"warehouse_remove_waiting_button_{supply.row_id}",
        ):
            ids = detail.loc[detail["Артикул"].isin(remove_skus), "row_id"].astype(int).tolist()
            removed = _safe_action(lambda: service.remove_waiting_from_supply(supply, ids))
            if removed is not None:
                st.success(f"Убрано позиций: {removed}")
                st.rerun()

        st.divider()
        st.caption("Пустую поставку без приёмки можно удалить полностью.")
        confirm = st.text_input(
            f"Для удаления введите {supply.supply_id}",
            key=f"warehouse_delete_supply_confirm_{supply.row_id}",
        )
        if st.button(
            "Удалить пустую поставку",
            disabled=confirm.strip() != supply.supply_id or supply.qty_received > 0,
            key=f"warehouse_delete_supply_{supply.row_id}",
        ):
            result = _safe_action(lambda: service.delete_empty_supply(supply))
            if result:
                st.success("Пустая поставка удалена.")
                st.rerun()


def _runtime_dir() -> Path:
    root = Path(".runtime") / "warehouse_uploads"
    root.mkdir(parents=True, exist_ok=True)
    return root


def _parse_uploaded_supply(uploaded: Any) -> tuple[list[Product], Path]:
    session_dir = _runtime_dir() / uuid.uuid4().hex
    image_dir = session_dir / "images"
    image_dir.mkdir(parents=True, exist_ok=True)
    path = session_dir / Path(uploaded.name).name
    path.write_bytes(uploaded.getvalue())
    products = load_products(path, image_dir)
    return products, session_dir


def _products_editor_frame(products: list[Product]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "№": product.number,
                "Артикул": product.sku,
                "Коробки": product.boxes,
                "По документу": product.qty_document,
                "Категория": product.category,
                "Материал": product.material,
                "Камень": product.stone,
                "Цвет": product.color,
                "Получено сейчас": bool(product.received),
                "Факт": product.actual_manual if product.actual_manual is not None else product.qty_document if product.received else 0,
                "Комментарий": product.comment,
                "image_path": product.image_path,
                "description": product.description,
                "unit_weight_kg": product.unit_weight_kg,
            }
            for product in products
        ]
    )


def _products_from_editor(frame: pd.DataFrame) -> list[Product]:
    products: list[Product] = []
    for _, row in frame.iterrows():
        received = bool(row.get("Получено сейчас", False))
        actual = as_int(row.get("Факт")) if received else None
        products.append(
            Product(
                number=as_int(row.get("№")),
                boxes=str(row.get("Коробки") or ""),
                sku=str(row.get("Артикул") or "").strip(),
                qty_document=as_int(row.get("По документу")),
                description=str(row.get("description") or ""),
                category=str(row.get("Категория") or ""),
                material=str(row.get("Материал") or ""),
                stone=str(row.get("Камень") or ""),
                color=str(row.get("Цвет") or ""),
                unit_weight_kg=(float(row.get("unit_weight_kg")) if pd.notna(row.get("unit_weight_kg")) else None),
                image_path=str(row.get("image_path") or ""),
                received=received,
                actual_manual=actual,
                comment=str(row.get("Комментарий") or ""),
            )
        )
    return products


def render_new_supply(config: Any) -> None:
    st.markdown('<div class="wm-title">Новая поставка</div>', unsafe_allow_html=True)
    service = _service(config)
    st.markdown(
        '<div class="wm-context">Загрузите Packing List или готовый Master. '
        'Файл обрабатывается только после нажатия кнопки и не хранится в памяти сайта постоянно.</div>',
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader(
        "Packing List / Master",
        type=["xlsx", "xlsm"],
        key="warehouse_supply_file",
    )
    if uploaded is not None and st.button("Разобрать файл", type="primary", key="warehouse_parse_supply"):
        try:
            with st.spinner("Извлекаем строки и фотографии..."):
                products, session_dir = _parse_uploaded_supply(uploaded)
            old = st.session_state.get("warehouse_supply_runtime_dir")
            if old:
                shutil.rmtree(str(old), ignore_errors=True)
            st.session_state["warehouse_supply_runtime_dir"] = str(session_dir)
            st.session_state["warehouse_supply_products"] = [product.to_dict() for product in products]
            st.success(f"Распознано: {len(products)} SKU")
            st.rerun()
        except Exception as exc:
            st.error(f"Не удалось разобрать файл: {exc}")

    raw_products = st.session_state.get("warehouse_supply_products", [])
    if not raw_products:
        st.info("После разбора здесь появится редактируемая таблица поставки.")
        return

    products = [Product.from_dict(item) for item in raw_products]
    frame = _products_editor_frame(products)
    visible_columns = [
        "№", "Артикул", "Коробки", "По документу", "Категория", "Материал",
        "Камень", "Цвет", "Получено сейчас", "Факт", "Комментарий",
    ]
    edited = st.data_editor(
        frame,
        column_order=visible_columns,
        hide_index=True,
        width="stretch",
        height=560,
        num_rows="fixed",
        key="warehouse_supply_editor",
        column_config={
            "Категория": st.column_config.SelectboxColumn("Категория", options=CATEGORIES),
            "По документу": st.column_config.NumberColumn("По документу", min_value=0, step=1),
            "Факт": st.column_config.NumberColumn("Факт", min_value=0, step=1),
            "Получено сейчас": st.column_config.CheckboxColumn("Получено сейчас"),
        },
    )
    updated_products = _products_from_editor(edited)
    st.session_state["warehouse_supply_products"] = [product.to_dict() for product in updated_products]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("SKU", len(updated_products))
    c2.metric("По документу", sum(product.qty_document for product in updated_products))
    c3.metric("Принимается сейчас", sum(product.actual_qty or 0 for product in updated_products))
    c4.metric("Ожидается", sum(product.waiting_qty for product in updated_products))

    info1, info2 = st.columns(2)
    supply_id = info1.text_input("Номер поставки *", value="SUP-2026-")
    supplier = info2.text_input("Поставщик")
    invoice = info1.text_input("Invoice")
    comment = info2.text_input("Комментарий")
    allow_compatibility = False
    if not int(getattr(config, "supply_lines_table_id", 0) or 0):
        st.warning(
            "Таблица «Позиции поставок» ещё не подключена. Повторяющиеся SKU из старых "
            "поставок будут заблокированы, чтобы не испортить историю."
        )
        allow_compatibility = st.checkbox(
            "Разрешить совместимость со старой схемой для повторяющихся SKU",
            value=False,
        )

    master_buffer = BytesIO()
    with tempfile.TemporaryDirectory(prefix="analitika-master-") as temp_dir:
        master_path = Path(temp_dir) / "Master.xlsx"
        export_master(master_path, updated_products)
        master_buffer.write(master_path.read_bytes())
    st.download_button(
        "Скачать проверенный Master",
        data=master_buffer.getvalue(),
        file_name=f"{supply_id or 'Master'}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if st.button("Создать поставку в Baserow", type="primary", width="stretch"):
        with st.spinner("Создаём карточки, поставку и приходные операции..."):
            result = _safe_action(
                lambda: service.create_supply_from_products(
                    supply_id=supply_id,
                    supplier=supplier,
                    invoice=invoice,
                    comment=comment,
                    products=updated_products,
                    allow_reused_sku_compatibility=allow_compatibility,
                )
            )
        if result:
            st.success(
                f"Поставка {result['supply_id']} создана. SKU: {result['sku']}; "
                f"принято: {result['received']} шт.; ожидается: {result['waiting']} шт."
            )
            runtime = st.session_state.pop("warehouse_supply_runtime_dir", None)
            st.session_state.pop("warehouse_supply_products", None)
            if runtime:
                shutil.rmtree(str(runtime), ignore_errors=True)
            st.rerun()


def render_receiving(config: Any) -> None:
    st.markdown('<div class="wm-title">Приёмка</div>', unsafe_allow_html=True)
    service = _service(config)
    mode = st.segmented_control(
        "Способ",
        ["По поставке", "Ручной приход"],
        default="По поставке",
        key="warehouse_receiving_mode",
    ) or "По поставке"

    if mode == "Ручной приход":
        section = st.segmented_control(
            "Раздел",
            ["Сувенирка", "Комплектующие"],
            default="Сувенирка",
            key="warehouse_manual_receipt_section",
        ) or "Сувенирка"
        items = service.catalog(section)
        labels = {f"{item.sku} · остаток {item.balance}": item for item in items}
        selected_labels = st.multiselect("Товары", list(labels), key="warehouse_manual_receipt_items")
        frame = pd.DataFrame(
            [{"Артикул": labels[label].sku, "Количество": 1, "row_id": labels[label].row_id} for label in selected_labels]
        )
        if not frame.empty:
            edited = st.data_editor(
                frame,
                column_order=["Артикул", "Количество"],
                hide_index=True,
                width="stretch",
                column_config={"Количество": st.column_config.NumberColumn(min_value=1, step=1)},
            )
            comment = st.text_input("Комментарий", key="warehouse_manual_receipt_comment")
            if st.button("Провести ручной приход", type="primary"):
                quantities = {as_int(row["row_id"]): as_int(row["Количество"]) for _, row in edited.iterrows()}
                result = _safe_action(
                    lambda: service.manual_operation(
                        operation_type="Приход",
                        section=section,
                        quantities=quantities,
                        comment=comment,
                    )
                )
                if result:
                    st.success(f"Проведено: {result['batch_id']} · {result['quantity']} шт.")
                    st.rerun()
        return

    summaries = [item for item in service.supply_summaries() if item.qty_waiting > 0]
    if not summaries:
        st.info("Нет поставок с ожидаемым количеством.")
        return
    options, mapping = _summary_options(summaries)
    supply = mapping[st.selectbox("Поставка", options, key="warehouse_receiving_supply")]
    rows = service.supply_products(supply)
    fill_max = st.checkbox("Заполнить максимальным ожидаемым количеством", value=False)
    frame = pd.DataFrame(
        [
            {
                "Артикул": row.get("Артикул"),
                "По документу": as_int(row.get("_document")),
                "Принято": as_int(row.get("_received")),
                "Ожидается": max(as_int(row.get("_document")) - as_int(row.get("_received")), 0),
                "Принять сейчас": max(as_int(row.get("_document")) - as_int(row.get("_received")), 0) if fill_max else 0,
                "row_id": int(row["id"]),
            }
            for row in rows
            if as_int(row.get("_document")) > as_int(row.get("_received"))
        ]
    )
    edited = st.data_editor(
        frame,
        column_order=["Артикул", "По документу", "Принято", "Ожидается", "Принять сейчас"],
        hide_index=True,
        width="stretch",
        height=520,
        column_config={
            "Принять сейчас": st.column_config.NumberColumn(min_value=0, step=1),
        },
    )
    total = int(pd.to_numeric(edited.get("Принять сейчас", pd.Series(dtype=int)), errors="coerce").fillna(0).sum())
    st.metric("К приёмке", f"{total:,} шт.")
    if st.button("Провести приёмку", type="primary", width="stretch", disabled=total <= 0):
        quantities = {as_int(row["row_id"]): as_int(row["Принять сейчас"]) for _, row in edited.iterrows()}
        result = _safe_action(lambda: service.receive_supply(supply, quantities))
        if result:
            st.success(f"Приёмка {result['batch_id']}: {result['sku']} SKU, {result['quantity']} шт.")
            st.rerun()


def _read_transfer_excel(data: bytes) -> list[tuple[str, int]]:
    workbook = load_workbook(BytesIO(data), read_only=True, data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(value or "").strip().casefold() for value in rows[0]]
    sku_index = next((i for i, value in enumerate(headers) if value in {"артикул", "sku", "item", "код"}), 0)
    qty_index = next((i for i, value in enumerate(headers) if value in {"количество", "qty", "quantity", "шт"}), 1 if len(headers) > 1 else 0)
    result: list[tuple[str, int]] = []
    for row in rows[1:]:
        sku = str(row[sku_index] or "").strip() if sku_index < len(row) else ""
        quantity = as_int(row[qty_index]) if qty_index < len(row) else 0
        if sku and quantity > 0:
            result.append((sku, quantity))
    return result


def render_transfer(config: Any) -> None:
    st.markdown('<div class="wm-title">Передача в бухгалтерию</div>', unsafe_allow_html=True)
    service = _service(config)
    tab_supply, tab_manual, tab_excel = st.tabs(["По поставке", "По SKU", "Из Excel"])

    with tab_supply:
        summaries = [item for item in service.supply_summaries() if item.qty_received > 0]
        if not summaries:
            st.info("Нет поставок с принятым товаром.")
        else:
            options, mapping = _summary_options(summaries)
            supply = mapping[st.selectbox("Поставка", options, key="warehouse_transfer_supply")]
            rows = service.supply_products(supply)
            already = service.transferred_by_supply(supply.supply_id)
            records = []
            for row in rows:
                row_id = int(row["id"])
                transferred = as_int(row.get("_transferred")) if service.has_supply_lines else already.get(row_id, 0)
                received = as_int(row.get("_received"))
                stock = as_int(row.get("Остаток"))
                maximum = min(max(received - transferred, 0), max(stock, 0))
                if maximum <= 0:
                    continue
                records.append(
                    {
                        "Артикул": row.get("Артикул"),
                        "Принято из поставки": received,
                        "Уже передано": transferred,
                        "Остаток": stock,
                        "Максимум": maximum,
                        "Передать": maximum,
                        "row_id": row_id,
                    }
                )
            frame = pd.DataFrame(records)
            if frame.empty:
                st.info("По этой поставке больше нет доступного количества для передачи.")
            else:
                edited = st.data_editor(
                    frame,
                    column_order=["Артикул", "Принято из поставки", "Уже передано", "Остаток", "Максимум", "Передать"],
                    hide_index=True,
                    width="stretch",
                    height=540,
                    column_config={"Передать": st.column_config.NumberColumn(min_value=0, step=1)},
                )
                total = int(pd.to_numeric(edited["Передать"], errors="coerce").fillna(0).sum())
                attention = edited.loc[(edited["Остаток"] - edited["Передать"]) <= 15]
                if not attention.empty:
                    with st.expander(f"После передачи требуют внимания: {len(attention)} позиций"):
                        show = attention[["Артикул", "Остаток", "Передать"]].copy()
                        show["Останется"] = show["Остаток"] - show["Передать"]
                        st.dataframe(show, hide_index=True, width="stretch", height=280)
                comment = st.text_input("Комментарий", value=f"Поставка {supply.supply_id}", key=f"warehouse_transfer_comment_{supply.row_id}")
                st.metric("К передаче", f"{total:,} шт.")
                if st.button("Передать выбранное", type="primary", width="stretch", disabled=total <= 0, key=f"warehouse_transfer_submit_{supply.row_id}"):
                    quantities = {as_int(row["row_id"]): as_int(row["Передать"]) for _, row in edited.iterrows()}
                    result = _safe_action(lambda: service.transfer_supply(supply, quantities, comment=comment))
                    if result:
                        st.success(f"Передача {result['batch_id']}: {result['sku']} SKU, {result['quantity']} шт.")
                        st.rerun()

    with tab_manual:
        section = st.segmented_control(
            "Раздел",
            ["Сувенирка", "Комплектующие"],
            default="Сувенирка",
            key="warehouse_manual_transfer_section",
        ) or "Сувенирка"
        items = [item for item in service.catalog(section) if item.balance > 0]
        labels = {f"{item.sku} · остаток {item.balance}": item for item in items}
        selected_labels = st.multiselect("Товары", list(labels), key="warehouse_manual_transfer_items")
        frame = pd.DataFrame(
            [{"Артикул": labels[label].sku, "Остаток": labels[label].balance, "Количество": 1, "row_id": labels[label].row_id} for label in selected_labels]
        )
        if not frame.empty:
            edited = st.data_editor(
                frame,
                column_order=["Артикул", "Остаток", "Количество"],
                hide_index=True,
                width="stretch",
                column_config={"Количество": st.column_config.NumberColumn(min_value=1, step=1)},
            )
            comment = st.text_input("Комментарий", key="warehouse_manual_transfer_comment")
            if st.button("Провести передачу", type="primary", key="warehouse_manual_transfer_submit"):
                quantities = {as_int(row["row_id"]): as_int(row["Количество"]) for _, row in edited.iterrows()}
                result = _safe_action(
                    lambda: service.manual_operation(
                        operation_type="Передача в бухгалтерию",
                        section=section,
                        quantities=quantities,
                        comment=comment,
                    )
                )
                if result:
                    st.success(f"Передача {result['batch_id']}: {result['quantity']} шт.")
                    st.rerun()

    with tab_excel:
        uploaded = st.file_uploader("Excel со столбцами SKU и Количество", type=["xlsx", "xlsm"], key="warehouse_transfer_excel")
        section = st.segmented_control(
            "Раздел для Excel",
            ["Сувенирка", "Комплектующие"],
            default="Сувенирка",
            key="warehouse_transfer_excel_section",
        ) or "Сувенирка"
        if uploaded is not None:
            requested = _read_transfer_excel(uploaded.getvalue())
            catalog = {item.sku.casefold(): item for item in service.catalog(section)}
            records = []
            missing = []
            for sku, quantity in requested:
                item = catalog.get(sku.casefold())
                if item is None:
                    missing.append(sku)
                    continue
                records.append({"Артикул": item.sku, "Остаток": item.balance, "Количество": quantity, "row_id": item.row_id})
            edited = st.data_editor(
                pd.DataFrame(records),
                column_order=["Артикул", "Остаток", "Количество"],
                hide_index=True,
                width="stretch",
                column_config={"Количество": st.column_config.NumberColumn(min_value=1, step=1)},
            )
            if missing:
                st.warning("Не найдены: " + ", ".join(missing[:30]))
            if records and st.button("Провести Excel-передачу", type="primary"):
                quantities = {as_int(row["row_id"]): as_int(row["Количество"]) for _, row in edited.iterrows()}
                result = _safe_action(
                    lambda: service.manual_operation(
                        operation_type="Передача в бухгалтерию",
                        section=section,
                        quantities=quantities,
                        comment=f"Excel {uploaded.name}",
                    )
                )
                if result:
                    st.success(f"Передача {result['batch_id']}: {result['quantity']} шт.")
                    st.rerun()


def render_operations(config: Any) -> None:
    from src.warehouse import normalize_operations

    st.markdown('<div class="wm-title">Операции</div>', unsafe_allow_html=True)
    service = _service(config)
    raw = service.client.list_rows(config.operations_table_id)
    frame = normalize_operations(raw)
    operation_types = ["Все", *sorted(frame["Тип операции"].dropna().astype(str).unique().tolist())] if not frame.empty else ["Все"]
    selected_type = st.selectbox("Тип", operation_types, key="warehouse_operations_type")
    query = st.text_input("SKU, Batch ID или поставка", key="warehouse_operations_query").strip().casefold()
    current = frame.copy()
    if selected_type != "Все":
        current = current.loc[current["Тип операции"] == selected_type]
    if query:
        current = current.loc[
            current.astype(str).apply(lambda row: query in " ".join(row).casefold(), axis=1)
        ]
    if "Дата" in current.columns:
        current["Дата"] = current["Дата"].dt.strftime("%d.%m.%Y %H:%M")
    st.dataframe(current, width="stretch", hide_index=True, height=560)

    with st.expander("Создать корректировку операции", expanded=False):
        selectable = [row for row in raw if as_int(row.get("Количество")) > 0]
        labels = {
            f"{row.get('Batch ID') or row.get('id')} · {select_text(row.get('Тип операции'))} · {row.get('Операция') or ''}": row
            for row in selectable
        }
        if not labels:
            st.info("Нет операций для корректировки.")
        else:
            label = st.selectbox("Операция", list(labels), key="warehouse_correction_operation")
            operation = labels[label]
            maximum = max(as_int(operation.get("Количество")), 1)
            quantity = st.number_input("Количество для отмены", min_value=1, max_value=maximum, value=maximum)
            comment = st.text_input("Причина корректировки", key="warehouse_correction_comment")
            confirm = st.checkbox("Подтверждаю создание обратной операции", key="warehouse_correction_confirm")
            if st.button("Создать корректировку", type="primary", disabled=not confirm):
                result = _safe_action(lambda: service.correct_operation(operation, quantity=int(quantity), comment=comment))
                if result:
                    st.success(f"Корректировка создана: {result['batch_id']}")
                    st.rerun()


def render_maintenance(config: Any) -> None:
    st.markdown('<div class="wm-title">Обслуживание</div>', unsafe_allow_html=True)
    service = _service(config)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Подключение")
        st.write(
            {
                "Baserow": config.base_url,
                "Сувенирка": config.souvenirs_table_id,
                "Комплектующие": config.components_table_id,
                "Операции": config.operations_table_id,
                "Поставки": config.supplies_table_id,
                "Позиции поставок": int(getattr(config, "supply_lines_table_id", 0) or 0) or "не создана",
            }
        )
        if st.button("Проверить и обновить данные", type="primary"):
            _clear_cache()
            result = _safe_action(service.diagnostics)
            if result:
                st.session_state["warehouse_diagnostics"] = result
                st.rerun()
    with c2:
        st.markdown("### Режим данных")
        if service.has_supply_lines:
            st.markdown('<div class="wm-good">Используется безопасная таблица «Позиции поставок».</div>', unsafe_allow_html=True)
        else:
            st.markdown(
                '<div class="wm-warning">Используется совместимость со старой схемой. '
                'Повторные SKU блокируются при создании поставки.</div>',
                unsafe_allow_html=True,
            )
            st.caption("Спецификация миграции находится в WAREHOUSE_MIGRATION_SUPPLY_LINES.md.")

    report = st.session_state.get("warehouse_diagnostics")
    if report:
        metrics = st.columns(4)
        metrics[0].metric("Сувенирные SKU", report["souvenir_sku"])
        metrics[1].metric("Комплектующие", report["component_sku"])
        metrics[2].metric("Операции", report["operations"])
        metrics[3].metric("Поставки", report["supplies"])
        with st.expander(f"Дубликаты SKU: {len(report['duplicate_sku'])}"):
            st.write(report["duplicate_sku"] or "Не найдено")
        with st.expander(f"Без фото: {len(report['without_photo'])}"):
            st.write(report["without_photo"] or "Не найдено")
        with st.expander(f"Без категории: {len(report['without_category'])}"):
            st.write(report["without_category"] or "Не найдено")


def render_warehouse_workspace(config: Any, selected_metal_groups: Iterable[str]) -> None:
    """Render one lazy warehouse workspace inside the existing Analitika mode."""
    st.markdown(WAREHOUSE_MANAGEMENT_CSS, unsafe_allow_html=True)
    st.markdown(
        '<div class="wm-context"><strong>Princess Warehouse Online</strong> · '
        'загружается только выбранный подраздел, поэтому остальные режимы Analitika '
        'не пересчитываются и не занимают память.</div>',
        unsafe_allow_html=True,
    )
    st.session_state.setdefault("warehouse_workspace", "Обзор")
    current = st.segmented_control(
        "Раздел склада",
        list(WORKSPACES),
        key="warehouse_workspace",
    ) or "Обзор"

    refresh, status = st.columns([1, 5])
    with refresh:
        if st.button("Обновить", key="warehouse_workspace_refresh", width="stretch"):
            _clear_cache()
            st.rerun()
    with status:
        status.caption("Кэш чтения — 60 секунд. После записи кэш очищается автоматически.")

    if current == "Обзор":
        render_overview(config, selected_metal_groups)
    elif current == "Каталог":
        render_catalog(config)
    elif current == "Поставки":
        render_supplies(config)
    elif current == "Новая поставка":
        render_new_supply(config)
    elif current == "Приёмка":
        render_receiving(config)
    elif current == "Передача в бухгалтерию":
        render_transfer(config)
    elif current == "Операции":
        render_operations(config)
    else:
        render_maintenance(config)
