from __future__ import annotations

from dataclasses import asdict, replace
import base64
import json
from html import escape
from io import BytesIO
from pathlib import Path
import shutil
from datetime import datetime, timedelta
import tempfile
from typing import Any, Iterable
from urllib.parse import urljoin
import uuid

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from PIL import Image, ImageOps

from .client import WarehouseClient, WarehouseClientError, as_int, link_ids, select_text
from .models import Product, SupplySummary
from .packing import CATEGORIES, export_master, load_products
from .service import WarehouseService, WarehouseServiceError
from .schema import BaserowSchemaManager, WarehouseSchemaError, SUPPLY_LINES_TABLE_NAME


WORKSPACES = (
    "Главная",
    "Товары",
    "Поставки",
    "Передача",
    "История",
)

SUPPLY_WORKSPACES = ("Реестр", "Новая поставка", "Приёмка")
HISTORY_WORKSPACES = ("Операции", "Обслуживание")

WAREHOUSE_MANAGEMENT_CSS = """
<style>
.wm-shell {border:1px solid rgba(183,137,63,.22);border-radius:20px;padding:18px 20px;
 background:linear-gradient(135deg,#fffdf8 0%,#f8f1e4 58%,#fff 100%);margin:.25rem 0 1rem;box-shadow:0 10px 30px rgba(70,51,25,.05)}
.wm-kicker {text-transform:uppercase;letter-spacing:.13em;font-size:11px;font-weight:800;color:#a47732;margin-bottom:5px}
.wm-shell-title {font-family:Georgia,serif;font-size:31px;line-height:1.08;color:#19140e;margin:0 0 5px}
.wm-shell-copy {color:#6f6253;font-size:14px;max-width:850px}
.wm-page-head {margin:.2rem 0 1rem}
.wm-page-head h2 {font-family:Georgia,serif;font-size:28px;line-height:1.12;color:#19140e;margin:0 0 4px}
.wm-page-head p {margin:0;color:#786b5d;font-size:14px}
.wm-context {border:1px solid rgba(183,137,63,.24);border-radius:13px;padding:11px 14px;background:#fffaf1;margin:.35rem 0 1rem;color:#5e5140}
.wm-good {border-left:4px solid #3a7d51;background:#f4fbf6;padding:11px 13px;border-radius:9px}
.wm-warning {border-left:4px solid #b7893f;background:#fffaf1;padding:11px 13px;border-radius:9px}
.wm-danger {border-left:4px solid #aa3939;background:#fff6f6;padding:11px 13px;border-radius:9px}
.wm-photo-placeholder {min-height:178px;border:1px dashed rgba(183,137,63,.35);border-radius:14px;display:flex;align-items:center;justify-content:center;background:#fbfaf7;color:#8a8175}
.wm-product-card {border:1px solid rgba(183,137,63,.20);border-radius:14px;padding:11px 12px;background:#fff;margin:-2px 0 8px;min-height:116px}
.wm-product-card .sku {font-family:Georgia,serif;font-size:19px;font-weight:700;color:#171411;margin-bottom:5px}
.wm-product-card .meta {font-size:12px;color:#71685c;min-height:34px;line-height:1.35}
.wm-product-card .stock {font-size:13px;font-weight:800;margin-top:7px}
.wm-stock-ok {color:#315d43}.wm-stock-low {color:#9b651c}.wm-stock-zero {color:#9c3535}
.wm-stepper {display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px;margin:.4rem 0 1rem}
.wm-step {border:1px solid #e8dfd0;border-radius:12px;padding:9px 10px;background:#fff;color:#796d5f;font-size:12px}
.wm-step b {display:block;color:#2a241d;font-size:13px;margin-bottom:2px}
.wm-step.active {border-color:#b7893f;background:#fff9ed;box-shadow:inset 0 0 0 1px rgba(183,137,63,.16)}
.wm-supply-card {border:1px solid rgba(183,137,63,.21);border-radius:15px;padding:13px 14px;background:#fff;margin-bottom:8px}
.wm-supply-card .title {font-family:Georgia,serif;font-weight:700;font-size:19px;color:#211a12}
.wm-supply-card .sub {color:#786d60;font-size:12px;margin:3px 0 10px}
.wm-status {display:inline-block;border-radius:999px;padding:3px 8px;font-size:11px;font-weight:800;margin-top:5px}
.wm-status-good {background:#eaf6ee;color:#2f6744}.wm-status-warn {background:#fff2d9;color:#8a5a18}.wm-status-neutral {background:#f1efeb;color:#6c6258}
.wm-empty {border:1px dashed #d9cfbf;border-radius:14px;padding:24px;text-align:center;color:#756a5e;background:#fcfbf8}
.wm-toolbar-note {font-size:12px;color:#84786b;padding-top:8px}
@media (max-width:900px){.wm-stepper{grid-template-columns:repeat(2,minmax(0,1fr))}.wm-shell-title{font-size:27px}}
@media (max-width:640px){.wm-shell{padding:14px}.wm-shell-title{font-size:24px}.wm-page-head h2{font-size:24px}.wm-photo-placeholder{min-height:135px}.wm-stepper{grid-template-columns:1fr}}
</style>
"""


def _clear_cache(*, photos: bool = False) -> None:
    """Clear only warehouse caches; do not invalidate analysis/SONU/order caches."""
    try:
        from src.warehouse import fetch_image_bytes, fetch_table_rows

        fetch_table_rows.clear()
        if photos:
            fetch_image_bytes.clear()
    except Exception:
        pass
    try:
        _remote_thumbnail_data_uri.clear()
    except Exception:
        pass


def _resolved_config(config: Any) -> Any:
    current_id = int(getattr(config, "supply_lines_table_id", 0) or 0)
    runtime_id = int(st.session_state.get("warehouse_supply_lines_table_id", 0) or 0)
    if current_id or runtime_id:
        return replace(config, supply_lines_table_id=current_id or runtime_id)
    try:
        discovered = WarehouseClient(config).discover_table_id(SUPPLY_LINES_TABLE_NAME)
    except WarehouseClientError:
        discovered = 0
    if discovered:
        st.session_state["warehouse_supply_lines_table_id"] = discovered
        return replace(config, supply_lines_table_id=discovered)
    return config


def _service(config: Any) -> WarehouseService:
    resolved = _resolved_config(config)
    return WarehouseService(WarehouseClient(resolved))


def _require_safe_schema(config: Any) -> WarehouseService | None:
    service = _service(config)
    if service.has_supply_lines:
        return service
    st.error(
        "Операция заблокирована: безопасная таблица «Позиции поставок» ещё не создана. "
        "Откройте История → Обслуживание и нажмите «Создать и мигрировать»."
    )
    if st.button("Перейти в обслуживание", key="warehouse_open_maintenance_schema"):
        st.session_state["warehouse_workspace"] = "История"
        st.session_state["warehouse_history_workspace"] = "Обслуживание"
        st.rerun()
    return None


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


def _photo_url(value: Any, base_url: str, *, size: str = "small") -> str:
    """Return a Baserow image URL, preferring a lightweight thumbnail."""
    files = value if isinstance(value, list) else [value] if value else []
    if not files or not isinstance(files[0], dict):
        return ""
    item = files[0]
    thumbnails = item.get("thumbnails") or {}
    orders = {
        "small": ("small", "card_cover", "tiny", "large"),
        "large": ("large", "card_cover", "small", "tiny"),
    }
    for name in orders.get(size, orders["small"]):
        candidate = thumbnails.get(name)
        if isinstance(candidate, dict) and candidate.get("url"):
            return urljoin(str(base_url).rstrip("/") + "/", str(candidate["url"]).strip())
    url = str(item.get("url") or "").strip()
    return urljoin(str(base_url).rstrip("/") + "/", url) if url else ""


def _item_photo_url(item: Any, config: Any, *, size: str = "small") -> str:
    return _photo_url(getattr(item, "photo", None), str(config.base_url), size=size)


def _row_photo_url(row: dict[str, Any], config: Any, *, size: str = "small") -> str:
    return _photo_url(row.get("Фото"), str(config.base_url), size=size)


def _local_thumbnail_data_uri(path_value: str, modified_ns: int = 0) -> str:
    del modified_ns
    path = Path(str(path_value or ""))
    if not path.exists():
        return ""
    try:
        with Image.open(path) as source:
            image = ImageOps.exif_transpose(source).convert("RGB")
            image.thumbnail((180, 180), getattr(Image, "Resampling", Image).LANCZOS)
            canvas = Image.new("RGB", (190, 190), "white")
            canvas.paste(image, ((190 - image.width) // 2, (190 - image.height) // 2))
            output = BytesIO()
            canvas.save(output, format="JPEG", quality=76, optimize=True)
        encoded = base64.b64encode(output.getvalue()).decode("ascii")
        return f"data:image/jpeg;base64,{encoded}"
    except Exception:
        return ""


@st.cache_data(ttl=1800, show_spinner=False, max_entries=2048)
def _remote_thumbnail_data_uri(url: str, token: str) -> str:
    """Fetch private Baserow media server-side and expose a compact browser-safe thumbnail."""
    if not url:
        return ""
    from src.warehouse import fetch_image_bytes

    raw = fetch_image_bytes(url, token)
    if not raw:
        return ""
    try:
        with Image.open(BytesIO(raw)) as source:
            image = ImageOps.exif_transpose(source).convert("RGB")
            image.thumbnail((190, 190), getattr(Image, "Resampling", Image).LANCZOS)
            canvas = Image.new("RGB", (196, 196), "white")
            canvas.paste(image, ((196 - image.width) // 2, (196 - image.height) // 2))
            output = BytesIO()
            canvas.save(output, format="JPEG", quality=78, optimize=True)
        encoded = base64.b64encode(output.getvalue()).decode("ascii")
        return f"data:image/jpeg;base64,{encoded}"
    except Exception:
        return ""


def _item_photo_data_uri(item: Any, config: Any) -> str:
    return _remote_thumbnail_data_uri(_item_photo_url(item, config, size="small"), str(config.token))


def _row_photo_data_uri(row: dict[str, Any], config: Any) -> str:
    return _remote_thumbnail_data_uri(_row_photo_url(row, config, size="small"), str(config.token))


def _catalog_dataframe(items: list[Any], config: Any) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Фото": _item_photo_data_uri(item, config),
                "Артикул": item.sku,
                "Раздел": item.section,
                "Остаток": item.balance,
                "Минимум": item.min_balance,
                "Категория": item.category,
                "Материал": item.material,
                "Камень": item.stone,
                "Цвет": item.color,
                "Коробки": item.boxes,
                "row_id": item.row_id,
            }
            for item in items
        ]
    )


def _render_item_photo(item: Any, config: Any, *, width: int | str = "stretch") -> None:
    data_uri = _item_photo_data_uri(item, config)
    if data_uri:
        st.image(data_uri, width=width)
    else:
        st.markdown('<div class="wm-photo-placeholder">Нет фотографии</div>', unsafe_allow_html=True)


def _render_catalog_cards(items: list[Any], config: Any, key: str) -> None:
    if not items:
        st.markdown('<div class="wm-empty">По выбранным фильтрам позиций нет.</div>', unsafe_allow_html=True)
        return
    page_size = st.segmented_control(
        "Карточек на странице",
        [6, 12, 18],
        default=12,
        key=f"{key}_page_size",
    ) or 12
    page_count = max(1, (len(items) + int(page_size) - 1) // int(page_size))
    page = st.selectbox(
        "Страница",
        list(range(1, page_count + 1)),
        format_func=lambda value: f"{value} из {page_count}",
        key=f"{key}_page",
    )
    current = items[(int(page) - 1) * int(page_size): int(page) * int(page_size)]
    for start in range(0, len(current), 3):
        columns = st.columns(3)
        for column, item in zip(columns, current[start:start + 3]):
            with column:
                with st.container(border=True):
                    _render_item_photo(item, config)
                    details = " · ".join(
                        part for part in (item.category, item.material, item.stone, item.color) if part
                    ) or "Характеристики не указаны"
                    stock_class = "wm-stock-zero" if item.balance <= 0 else "wm-stock-low" if item.balance <= 15 else "wm-stock-ok"
                    stock_text = "Нет в наличии" if item.balance <= 0 else "Заканчивается" if item.balance <= 15 else "В наличии"
                    st.markdown(
                        '<div class="wm-product-card">'
                        f'<div class="sku">{escape(item.sku)}</div>'
                        f'<div class="meta">{escape(details)}</div>'
                        f'<div class="stock {stock_class}">{stock_text} · {int(item.balance):,} шт. · минимум {int(item.min_balance):,}</div>'
                        '</div>',
                        unsafe_allow_html=True,
                    )
                    if st.button("Открыть карточку", key=f"{key}_open_{item.row_id}", width="stretch"):
                        st.session_state["warehouse_catalog_mode"] = "Управление"
                        st.session_state["warehouse_catalog_action"] = "Редактировать"
                        st.session_state["warehouse_catalog_selected_id"] = int(item.row_id)
                        st.rerun()


def _page_header(title: str, copy: str) -> None:
    st.markdown(
        '<div class="wm-page-head">'
        f'<h2>{escape(title)}</h2>'
        f'<p>{escape(copy)}</p>'
        '</div>',
        unsafe_allow_html=True,
    )


def _navigate(workspace: str, subpage: str | None = None) -> None:
    st.session_state["warehouse_workspace"] = workspace
    if workspace == "Поставки" and subpage:
        st.session_state["warehouse_supply_workspace"] = subpage
    if workspace == "История" and subpage:
        st.session_state["warehouse_history_workspace"] = subpage


def _workflow(active: int = 0) -> None:
    steps = (
        ("1", "Создать поставку", "Загрузить Master и проверить фото"),
        ("2", "Принять товар", "Указать фактическое количество"),
        ("3", "Передать", "Отправить выбранное в бухгалтерию"),
        ("4", "Проверить историю", "Контролировать остатки и операции"),
    )
    cards = []
    for index, (number, title, copy) in enumerate(steps, start=1):
        css = "wm-step active" if active == index else "wm-step"
        cards.append(f'<div class="{css}"><b>{number}. {escape(title)}</b>{escape(copy)}</div>')
    st.markdown('<div class="wm-stepper">' + ''.join(cards) + '</div>', unsafe_allow_html=True)


def _supply_status_class(status: str, waiting: int) -> str:
    value = str(status or "").casefold()
    if waiting <= 0 or "полностью" in value:
        return "wm-status-good"
    if "част" in value or waiting > 0:
        return "wm-status-warn"
    return "wm-status-neutral"


def _reset_supply_draft() -> None:
    runtime = st.session_state.pop("warehouse_supply_runtime_dir", None)
    st.session_state.pop("warehouse_supply_products", None)
    st.session_state.pop("warehouse_supply_editor", None)
    st.session_state.pop("warehouse_supply_file", None)
    if runtime:
        shutil.rmtree(str(runtime), ignore_errors=True)


def _selected_summary(summaries: list[SupplySummary], row_id: int | None) -> SupplySummary | None:
    return next((item for item in summaries if int(item.row_id) == int(row_id or 0)), None)


def render_overview(config: Any, selected_metal_groups: Iterable[str]) -> None:
    from src.warehouse import filter_warehouse_bundle, load_bundle, render_attention, render_overview

    _page_header("Главная склада", "Текущие остатки, проблемные позиции и быстрый переход к ежедневным операциям.")
    _workflow(0)
    quick = st.columns(3)
    with quick[0]:
        with st.container(border=True):
            st.markdown("### Новая поставка")
            st.caption("Загрузить Packing List или Master, проверить фото и создать поставку.")
            st.button("Добавить поставку", type="primary", width="stretch", on_click=_navigate, args=("Поставки", "Новая поставка"), key="warehouse_home_new")
    with quick[1]:
        with st.container(border=True):
            st.markdown("### Приёмка")
            st.caption("Принять всё ожидаемое или указать фактическое количество по строкам.")
            st.button("Перейти к приёмке", width="stretch", on_click=_navigate, args=("Поставки", "Приёмка"), key="warehouse_home_receive")
    with quick[2]:
        with st.container(border=True):
            st.markdown("### Бухгалтерия")
            st.caption("Выбрать поставку и передать максимально доступное количество.")
            st.button("Передать товар", width="stretch", on_click=_navigate, args=("Передача",), key="warehouse_home_transfer")

    st.divider()
    with st.spinner("Загружаем актуальный склад..."):
        bundle = load_bundle(config)
    selected = tuple(str(value) for value in selected_metal_groups)
    if selected:
        bundle = filter_warehouse_bundle(bundle, selected)
    render_overview(bundle)
    with st.expander("Позиции, требующие внимания", expanded=False):
        render_attention(bundle)


def render_catalog(config: Any) -> None:
    _page_header("Товары", "Просматривайте каталог с фотографиями и управляйте карточками без перехода в Baserow.")
    service = _service(config)
    section_col, mode_col = st.columns([1, 1])
    with section_col:
        section = st.segmented_control(
            "Раздел",
            ["Сувенирка", "Комплектующие"],
            default="Сувенирка",
            key="warehouse_catalog_manage_section",
        ) or "Сувенирка"
    with mode_col:
        mode = st.segmented_control(
            "Режим",
            ["Каталог", "Управление"],
            default="Каталог",
            key="warehouse_catalog_mode",
        ) or "Каталог"
    show_archive = st.checkbox(
        "Показать архивные карточки",
        value=False,
        key=f"warehouse_catalog_archive_{section}",
    )
    with st.spinner("Читаем каталог Baserow..."):
        items = service.catalog(section, include_inactive=show_archive)

    if mode == "Каталог":
        filter_cols = st.columns([2.6, 1, 1.2, 1])
        query = filter_cols[0].text_input(
            "Поиск",
            placeholder="Артикул, категория, материал, камень, коробка",
            key=f"warehouse_catalog_search_{section}",
        ).strip().casefold()
        status = filter_cols[1].selectbox(
            "Остаток", ["Все", "Есть", "Мало", "Нет"], key=f"warehouse_catalog_status_{section}"
        )
        categories = sorted({item.category for item in items if item.category})
        category = filter_cols[2].selectbox(
            "Категория", ["Все", *categories], key=f"warehouse_catalog_category_{section}"
        )
        view = filter_cols[3].selectbox(
            "Вид", ["Карточки", "Таблица"], key=f"warehouse_catalog_view_{section}"
        )
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
            if category != "Все" and item.category != category:
                continue
            filtered.append(item)
        metrics = st.columns(4)
        metrics[0].metric("Найдено SKU", len(filtered))
        metrics[1].metric("С фотографией", sum(bool(_item_photo_url(item, config)) for item in filtered))
        metrics[2].metric("Заканчиваются", sum(0 < item.balance <= 15 for item in filtered))
        metrics[3].metric("Нет в наличии", sum(item.balance <= 0 for item in filtered))
        if view == "Карточки":
            _render_catalog_cards(filtered, config, f"warehouse_catalog_cards_{section}")
        else:
            st.dataframe(
                _catalog_dataframe(filtered, config).drop(columns=["row_id"], errors="ignore"),
                width="stretch",
                hide_index=True,
                height=650,
                row_height=92,
                column_config={
                    "Фото": st.column_config.ImageColumn("Фото", width="medium"),
                    "Остаток": st.column_config.NumberColumn("Остаток", format="localized"),
                    "Минимум": st.column_config.NumberColumn("Минимум", format="localized"),
                },
            )
        return

    action = st.segmented_control(
        "Действие",
        ["Добавить", "Редактировать", "Удалить / деактивировать"],
        default="Редактировать" if items else "Добавить",
        key="warehouse_catalog_action",
    )
    if action == "Добавить":
        st.markdown('<div class="wm-context">Создайте постоянную карточку товара. Остаток появится только после операции прихода.</div>', unsafe_allow_html=True)
        with st.form(f"warehouse_add_catalog_{section}", clear_on_submit=True):
            c1, c2 = st.columns(2)
            sku = c1.text_input("Артикул *")
            category = c2.text_input("Категория")
            material = c1.text_input("Материал", placeholder="Steel; Brass")
            stone = c2.text_input("Камни", placeholder="Agate; Pearl")
            color = c1.text_input("Цвет")
            c2.caption("Коробки указываются в конкретной поставке, а не в карточке товара.")
            boxes = ""
            minimum = c1.number_input("Минимальный остаток", min_value=1, value=10, step=1)
            photo = c2.file_uploader("Фото", type=["jpg", "jpeg", "png", "webp"], key=f"warehouse_add_photo_{section}")
            comment = st.text_area("Комментарий")
            submitted = st.form_submit_button("Создать карточку", type="primary", width="stretch")
        if submitted:
            temp_path = None
            if photo is not None:
                suffix = Path(photo.name).suffix or ".jpg"
                temp_path = Path(tempfile.gettempdir()) / f"warehouse-photo-{uuid.uuid4().hex}{suffix}"
                temp_path.write_bytes(photo.getvalue())
            result = _safe_action(lambda: service.add_catalog_item(
                section=section, sku=sku, category=category, material=material, stone=stone,
                color=color, boxes=boxes, minimum=int(minimum), comment=comment, photo_path=temp_path,
            ))
            if temp_path: temp_path.unlink(missing_ok=True)
            if result is not None:
                st.success(f"Карточка {sku} создана.")
                st.rerun()
        return

    if not items:
        st.info("Каталог пуст.")
        return
    labels = {f"{item.sku} · остаток {item.balance}": item for item in items}
    selected_id = int(st.session_state.get("warehouse_catalog_selected_id", 0) or 0)
    label_values = list(labels)
    default_index = next((i for i, label in enumerate(label_values) if labels[label].row_id == selected_id), 0)
    label = st.selectbox("Карточка", label_values, index=default_index, key=f"warehouse_manage_item_{section}_{action}")
    item = labels[label]
    st.session_state["warehouse_catalog_selected_id"] = int(item.row_id)
    preview, editor = st.columns([1, 2])
    with preview:
        _render_item_photo(item, config)
        st.markdown(f"**{item.sku}**")
        st.caption(f"Остаток {item.balance} шт. · минимум {item.min_balance}")
    with editor:
        if action == "Редактировать":
            with st.form(f"warehouse_edit_form_{section}_{item.row_id}"):
                c1, c2 = st.columns(2)
                category = c1.text_input("Категория", value=item.category)
                material = c2.text_input("Материал", value=item.material)
                stone = c1.text_input("Камни", value=item.stone)
                color = c2.text_input("Цвет", value=item.color)
                boxes = item.boxes
                c1.caption("Коробки редактируются внутри позиции поставки.")
                minimum = c2.number_input("Минимальный остаток", min_value=1, value=max(int(item.min_balance), 1))
                replacement_photo = st.file_uploader("Заменить фотографию", type=["jpg", "jpeg", "png", "webp"], key=f"warehouse_edit_photo_{section}_{item.row_id}")
                comment = st.text_area("Комментарий", value=str((item.raw or {}).get("Комментарий") or ""))
                saved = st.form_submit_button("Сохранить изменения", type="primary", width="stretch")
            if saved:
                photo_path = None
                if replacement_photo is not None:
                    suffix = Path(replacement_photo.name).suffix or ".jpg"
                    photo_path = Path(tempfile.gettempdir()) / f"warehouse-photo-{uuid.uuid4().hex}{suffix}"
                    photo_path.write_bytes(replacement_photo.getvalue())
                result = _safe_action(lambda: service.update_catalog_item(section, item.row_id, {
                    "Категория": category or None, "Материал": material, "Камень": stone,
                    "Цвет": color, "Минимальный остаток": int(minimum),
                    "Комментарий": comment,
                }, photo_path=photo_path))
                if photo_path: photo_path.unlink(missing_ok=True)
                if result:
                    st.success(f"Карточка {item.sku} обновлена.")
                    st.rerun()
        else:
            st.markdown('<div class="wm-warning">Карточка с операциями не удаляется физически — она деактивируется, чтобы сохранить историю.</div>', unsafe_allow_html=True)
            confirmation = st.text_input(f"Для подтверждения введите артикул {item.sku}", key=f"warehouse_delete_confirmation_{section}_{item.row_id}")
            if st.button("Удалить или деактивировать", type="primary", disabled=confirmation.strip() != item.sku, key=f"warehouse_delete_button_{section}_{item.row_id}"):
                result = _safe_action(lambda: service.deactivate_or_delete_catalog_item(section, item.row_id))
                if result == "deleted": st.success("Карточка удалена: по ней не было операций.")
                elif result == "deactivated": st.success("Карточка деактивирована, история сохранена.")
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
    _page_header("Реестр поставок", "Быстро найдите поставку, посмотрите прогресс и откройте товары с фотографиями.")
    service = _service(config)
    with st.spinner("Загружаем реестр поставок..."):
        summaries = service.supply_summaries()
    if not summaries:
        st.markdown('<div class="wm-empty">В Baserow пока нет связанных поставок.</div>', unsafe_allow_html=True)
        return

    selected = _selected_summary(summaries, st.session_state.get("warehouse_selected_supply_id"))
    if selected is None:
        filters = st.columns([2, 1, 1])
        query = filters[0].text_input("Поиск", placeholder="Номер или поставщик", key="warehouse_supply_search").strip().casefold()
        statuses = sorted({item.status for item in summaries if item.status})
        status = filters[1].selectbox("Статус", ["Все", *statuses], key="warehouse_supply_status_filter")
        only_open = filters[2].toggle("Только незавершённые", value=False, key="warehouse_supply_only_open")
        current = [item for item in summaries if (not query or query in f"{item.supply_id} {item.supplier}".casefold()) and (status == "Все" or item.status == status) and (not only_open or item.qty_waiting > 0)]
        metrics = st.columns(4)
        metrics[0].metric("Поставок", len(current))
        metrics[1].metric("По документу", f"{sum(x.qty_document for x in current):,}")
        metrics[2].metric("Принято", f"{sum(x.qty_received for x in current):,}")
        metrics[3].metric("Ожидается", f"{sum(x.qty_waiting for x in current):,}")
        if not current:
            st.markdown('<div class="wm-empty">Поставок по выбранным фильтрам нет.</div>', unsafe_allow_html=True)
            return
        for start in range(0, len(current), 2):
            columns = st.columns(2)
            for column, supply in zip(columns, current[start:start + 2]):
                with column:
                    with st.container(border=True):
                        status_class = _supply_status_class(supply.status, supply.qty_waiting)
                        progress = 1.0 if supply.qty_document <= 0 else min(max(supply.qty_received / supply.qty_document, 0.0), 1.0)
                        st.markdown(
                            '<div class="wm-supply-card">'
                            f'<div class="title">{escape(supply.supply_id)}</div>'
                            f'<div class="sub">{escape(supply.date or "Без даты")} · {escape(supply.supplier or "Поставщик не указан")}</div>'
                            f'<span class="wm-status {status_class}">{escape(supply.status or "Без статуса")}</span>'
                            '</div>', unsafe_allow_html=True,
                        )
                        st.progress(progress)
                        small = st.columns(3)
                        small[0].metric("SKU", supply.sku_total)
                        small[1].metric("Принято", supply.qty_received)
                        small[2].metric("Осталось", supply.qty_waiting)
                        if st.button("Открыть поставку", key=f"warehouse_open_supply_{supply.row_id}", width="stretch"):
                            st.session_state["warehouse_selected_supply_id"] = int(supply.row_id)
                            st.rerun()
        with st.expander("Показать все поставки таблицей", expanded=False):
            st.dataframe(_supply_table(current), width="stretch", hide_index=True, height=360)
        return

    top = st.columns([1, 4, 1])
    if top[0].button("← К реестру", key="warehouse_supply_back"):
        st.session_state.pop("warehouse_selected_supply_id", None)
        st.rerun()
    top[1].markdown(f"### {selected.supply_id}")
    if selected.qty_waiting > 0:
        top[2].button("Принять", type="primary", width="stretch", on_click=_navigate, args=("Поставки", "Приёмка"), key=f"warehouse_supply_to_receiving_{selected.row_id}")
        st.session_state["warehouse_receiving_supply_id"] = int(selected.row_id)
    summary_cols = st.columns(4)
    summary_cols[0].metric("SKU", selected.sku_total)
    summary_cols[1].metric("По документу", selected.qty_document)
    summary_cols[2].metric("Принято", selected.qty_received)
    summary_cols[3].metric("Ожидается", selected.qty_waiting)
    rows = service.supply_products(selected)
    detail = pd.DataFrame([{
        "Фото": _row_photo_data_uri(row, config), "Артикул": row.get("Артикул"), "Коробки": row.get("_boxes"),
        "По документу": as_int(row.get("_document")), "Принято": as_int(row.get("_received")),
        "Ожидается": max(as_int(row.get("_document")) - as_int(row.get("_received")), 0),
        "Остаток": as_int(row.get("Остаток")), "row_id": int(row["id"]),
    } for row in rows])
    st.dataframe(detail.drop(columns=["row_id"], errors="ignore"), width="stretch", hide_index=True, height=510, row_height=92, column_config={"Фото": st.column_config.ImageColumn("Фото", width="medium")})
    with st.expander("Исправить поставку", expanded=False):
        waiting = detail.loc[detail["Принято"] <= 0] if not detail.empty else detail
        remove_skus = st.multiselect("Убрать непринятые позиции", waiting["Артикул"].tolist() if not waiting.empty else [], key=f"warehouse_remove_waiting_{selected.row_id}")
        if st.button("Убрать выбранные из поставки", disabled=not remove_skus, key=f"warehouse_remove_waiting_button_{selected.row_id}"):
            ids = detail.loc[detail["Артикул"].isin(remove_skus), "row_id"].astype(int).tolist()
            removed = _safe_action(lambda: service.remove_waiting_from_supply(selected, ids))
            if removed is not None:
                st.success(f"Убрано позиций: {removed}")
                st.rerun()
        st.divider()
        st.caption("Полностью удалить можно только пустую поставку без приёмки.")
        confirm = st.text_input(f"Для удаления введите {selected.supply_id}", key=f"warehouse_delete_supply_confirm_{selected.row_id}")
        if st.button("Удалить пустую поставку", disabled=confirm.strip() != selected.supply_id or selected.qty_received > 0, key=f"warehouse_delete_supply_{selected.row_id}"):
            result = _safe_action(lambda: service.delete_empty_supply(selected))
            if result:
                st.session_state.pop("warehouse_selected_supply_id", None)
                st.success("Пустая поставка удалена.")
                st.rerun()


def _runtime_dir() -> Path:
    root = Path(".runtime") / "warehouse_uploads"
    root.mkdir(parents=True, exist_ok=True)
    cutoff = datetime.now() - timedelta(hours=24)
    for child in root.iterdir():
        try:
            if child.is_dir() and datetime.fromtimestamp(child.stat().st_mtime) < cutoff:
                shutil.rmtree(child, ignore_errors=True)
        except OSError:
            continue
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
                "Фото": _local_thumbnail_data_uri(
                    product.image_path,
                    Path(product.image_path).stat().st_mtime_ns if product.image_path and Path(product.image_path).exists() else 0,
                ),
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
    _page_header("Новая поставка", "Три шага: загрузите Excel, проверьте товары и подтвердите создание в Baserow.")
    service = _require_safe_schema(config)
    if service is None:
        return
    raw_products = st.session_state.get("warehouse_supply_products", [])
    _workflow(1 if not raw_products else 2)
    if not raw_products:
        with st.container(border=True):
            st.markdown("### 1. Загрузите Packing List или Master")
            st.caption("Поддерживаются XLSX и XLSM до 150 МБ. Встроенные изображения будут извлечены автоматически.")
            uploaded = st.file_uploader("Файл поставки", type=["xlsx", "xlsm"], key="warehouse_supply_file", label_visibility="collapsed")
            action_cols = st.columns([1, 2])
            if action_cols[0].button("Разобрать файл", type="primary", width="stretch", disabled=uploaded is None, key="warehouse_parse_supply"):
                try:
                    with st.spinner("Извлекаем строки и фотографии..."):
                        products, session_dir = _parse_uploaded_supply(uploaded)
                    old = st.session_state.get("warehouse_supply_runtime_dir")
                    if old: shutil.rmtree(str(old), ignore_errors=True)
                    st.session_state["warehouse_supply_runtime_dir"] = str(session_dir)
                    st.session_state["warehouse_supply_products"] = [product.to_dict() for product in products]
                    st.success(f"Распознано: {len(products)} SKU")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Не удалось разобрать файл: {exc}")
            action_cols[1].caption("Файл не отправляется в браузер после разбора; рабочий черновик хранится во временной папке текущей сессии.")
        return

    products = [Product.from_dict(item) for item in raw_products]
    summary = st.columns(5)
    summary[0].metric("SKU", len(products))
    summary[1].metric("С фото", sum(bool(product.image_path and Path(product.image_path).exists()) for product in products))
    summary[2].metric("По документу", sum(product.qty_document for product in products))
    summary[3].metric("Принимается", sum(product.actual_qty or 0 for product in products))
    summary[4].metric("Ожидается", sum(product.waiting_qty for product in products))
    control = st.columns([1, 1, 3])
    if control[0].button("Загрузить другой файл", width="stretch", key="warehouse_supply_reset"):
        _reset_supply_draft(); st.rerun()
    control[2].caption("Фотографии показаны прямо в таблице. Категорию, камни и фактическое количество можно исправить до сохранения.")

    frame = _products_editor_frame(products)
    visible_columns = ["Фото", "№", "Артикул", "Коробки", "По документу", "Категория", "Материал", "Камень", "Цвет", "Получено сейчас", "Факт", "Комментарий"]
    edited = st.data_editor(
        frame, column_order=visible_columns, hide_index=True, width="stretch", height=590, row_height=92,
        num_rows="fixed", disabled=["Фото"], key="warehouse_supply_editor",
        column_config={
            "Фото": st.column_config.ImageColumn("Фото", width="medium"),
            "Категория": st.column_config.SelectboxColumn("Категория", options=CATEGORIES),
            "По документу": st.column_config.NumberColumn("По документу", min_value=0, step=1),
            "Факт": st.column_config.NumberColumn("Факт", min_value=0, step=1),
            "Получено сейчас": st.column_config.CheckboxColumn("Получено сейчас"),
        },
    )
    updated_products = _products_from_editor(edited)
    st.session_state["warehouse_supply_products"] = [product.to_dict() for product in updated_products]
    issues = []
    issues += [f"Строка {p.number}: нет артикула" for p in updated_products if not p.sku]
    issues += [f"{p.sku or p.number}: количество по документу равно 0" for p in updated_products if p.qty_document <= 0]
    sku_rows: dict[str, list[int]] = {}
    for product in updated_products:
        if product.sku:
            sku_rows.setdefault(product.sku.casefold(), []).append(product.number)
    for sku_key, row_numbers in sku_rows.items():
        if len(row_numbers) > 1:
            shown = next(product.sku for product in updated_products if product.sku.casefold() == sku_key)
            issues.append(f"SKU {shown} повторяется в строках {', '.join(map(str, row_numbers))}")
    no_photo = [p.sku for p in updated_products if not (p.image_path and Path(p.image_path).exists())]
    with st.expander(f"Проверка данных · ошибок {len(issues)} · без фото {len(no_photo)}", expanded=bool(issues)):
        if issues: st.error("\n".join(f"• {item}" for item in issues[:50]))
        else: st.success("Обязательные поля заполнены.")
        if no_photo: st.warning("Без фотографии: " + ", ".join(no_photo[:40]))

    st.markdown("### 3. Подтвердите поставку")
    info1, info2 = st.columns(2)
    default_supply_id = st.session_state.setdefault("warehouse_new_supply_id", service.next_supply_id())
    supply_id = info1.text_input("Номер поставки *", value=default_supply_id, key="warehouse_new_supply_id_input")
    section = info2.segmented_control(
        "Тип поставки",
        ["Сувенирка", "Комплектующие"],
        default="Сувенирка",
        key="warehouse_new_supply_section",
    ) or "Сувенирка"
    supplier = info2.text_input("Поставщик")
    invoice = info1.text_input("Invoice")
    comment = info2.text_input("Комментарий")
    master_buffer = BytesIO()
    with tempfile.TemporaryDirectory(prefix="analitika-master-") as temp_dir:
        master_path = Path(temp_dir) / "Master.xlsx"
        export_master(master_path, updated_products)
        master_buffer.write(master_path.read_bytes())
    download_col, create_col = st.columns([1, 2])
    download_col.download_button("Скачать проверенный Master", data=master_buffer.getvalue(), file_name=f"{supply_id or 'Master'}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    can_create = bool(supply_id.strip()) and bool(updated_products) and not issues
    if create_col.button("Создать поставку в Baserow", type="primary", width="stretch", disabled=not can_create):
        with st.spinner("Создаём карточки, поставку и приходные операции..."):
            command_id = st.session_state.setdefault(
                "warehouse_new_supply_command_id", f"IMPORT-{uuid.uuid4().hex}"
            )
            result = _safe_action(lambda: service.create_supply_from_products(
                supply_id=supply_id, supplier=supplier, invoice=invoice, comment=comment,
                products=updated_products, section=section, command_id=command_id,
            ))
        if result:
            _reset_supply_draft()
            st.success(f"Поставка {result['supply_id']} создана. SKU: {result['sku']}; принято: {result['received']} шт.; ожидается: {result['waiting']} шт.")
            if result.get("failed_photos"):
                st.warning("Не удалось загрузить фото: " + ", ".join(result["failed_photos"][:30]))
            st.session_state.pop("warehouse_new_supply_command_id", None)
            st.session_state.pop("warehouse_new_supply_id", None)
            st.session_state["warehouse_supply_workspace"] = "Реестр"
            st.rerun()


def render_receiving(config: Any) -> None:
    _page_header("Приёмка", "Выберите поставку, проверьте фотографии и укажите фактически полученное количество.")
    service = _require_safe_schema(config)
    if service is None:
        return
    mode = st.segmented_control("Способ", ["По поставке", "Ручной приход"], default="По поставке", key="warehouse_receiving_mode") or "По поставке"
    if mode == "Ручной приход":
        section = st.segmented_control("Раздел", ["Сувенирка", "Комплектующие"], default="Сувенирка", key="warehouse_manual_receipt_section") or "Сувенирка"
        items = service.catalog(section)
        labels = {f"{item.sku} · остаток {item.balance}": item for item in items}
        selected_labels = st.multiselect("Найдите и выберите товары", list(labels), key="warehouse_manual_receipt_items")
        frame = pd.DataFrame([{"Фото": _item_photo_data_uri(labels[label], config), "Артикул": labels[label].sku, "Количество": 1, "row_id": labels[label].row_id} for label in selected_labels])
        if frame.empty:
            st.markdown('<div class="wm-empty">Выберите хотя бы один товар.</div>', unsafe_allow_html=True)
            return
        edited = st.data_editor(frame, column_order=["Фото", "Артикул", "Количество"], hide_index=True, width="stretch", row_height=92, disabled=["Фото", "Артикул"], column_config={"Фото": st.column_config.ImageColumn("Фото", width="medium"), "Количество": st.column_config.NumberColumn(min_value=1, step=1)})
        comment = st.text_input(
            "Основание и комментарий *",
            placeholder="Например: товар без Packing List / возврат / инвентаризация",
            key="warehouse_manual_receipt_comment",
        )
        total = int(pd.to_numeric(edited["Количество"], errors="coerce").fillna(0).sum())
        action = st.columns([1, 2])
        action[0].metric("К приходу", f"{total:,} шт.")
        if action[1].button(
            "Провести ручной приход",
            type="primary",
            width="stretch",
            disabled=total <= 0 or not comment.strip(),
        ):
            quantities = {as_int(row["row_id"]): as_int(row["Количество"]) for _, row in edited.iterrows()}
            command_id = st.session_state.setdefault("warehouse_manual_receipt_command", f"CMD-REC-{uuid.uuid4().hex}")
            result = _safe_action(lambda: service.manual_operation(
                operation_type="Приход", section=section, quantities=quantities,
                comment=comment, command_id=command_id,
            ))
            if result:
                st.success(f"Проведено: {result['batch_id']} · {result['quantity']} шт.")
                st.session_state.pop("warehouse_manual_receipt_command", None)
                st.rerun()
        return

    summaries = [item for item in service.supply_summaries() if item.qty_waiting > 0]
    if not summaries:
        st.markdown('<div class="wm-empty">Нет поставок с ожидаемым количеством.</div>', unsafe_allow_html=True)
        return
    options, mapping = _summary_options(summaries)
    preferred_id = int(st.session_state.get("warehouse_receiving_supply_id", 0) or 0)
    option_values = list(options)
    default_index = next((i for i, label in enumerate(option_values) if mapping[label].row_id == preferred_id), 0)
    selected_label = st.selectbox("Поставка", option_values, index=default_index, key="warehouse_receiving_supply")
    supply = mapping[selected_label]
    st.session_state["warehouse_receiving_supply_id"] = int(supply.row_id)
    supply_metrics = st.columns(4)
    supply_metrics[0].metric("SKU", supply.sku_total)
    supply_metrics[1].metric("По документу", supply.qty_document)
    supply_metrics[2].metric("Уже принято", supply.qty_received)
    supply_metrics[3].metric("Ожидается", supply.qty_waiting)
    rows = service.supply_products(supply)
    revision_key = f"warehouse_receiving_revision_{supply.row_id}"
    revision = int(st.session_state.get(revision_key, 0))
    mode_key = f"warehouse_receiving_fill_{supply.row_id}"
    fill_mode = st.session_state.get(mode_key, "empty")
    buttons = st.columns([1, 1, 3])
    if buttons[0].button("Принять всё", type="primary", width="stretch", key=f"warehouse_receiving_all_{supply.row_id}"):
        st.session_state[mode_key] = "max"; st.session_state[revision_key] = revision + 1; st.rerun()
    if buttons[1].button("Очистить", width="stretch", key=f"warehouse_receiving_clear_{supply.row_id}"):
        st.session_state[mode_key] = "empty"; st.session_state[revision_key] = revision + 1; st.rerun()
    buttons[2].caption("После массового заполнения любое количество можно изменить вручную.")
    draft_key = f"warehouse_receiving_draft_{supply.row_id}"
    saved_draft = st.session_state.get(draft_key, {})
    frame = pd.DataFrame([{
        "Фото": _row_photo_data_uri(row, config), "Артикул": row.get("Артикул"),
        "По документу": as_int(row.get("_document")), "Принято": as_int(row.get("_received")),
        "Ожидается": max(as_int(row.get("_document")) - as_int(row.get("_received")), 0),
        "Принять сейчас": (
            max(as_int(row.get("_document")) - as_int(row.get("_received")), 0)
            if fill_mode == "max"
            else as_int(saved_draft.get(str(as_int(row.get("_line_id")) or int(row["id"])), 0))
        ),
        "line_id": as_int(row.get("_line_id")) or int(row["id"]),
        "product_row_id": int(row["id"]),
    } for row in rows if as_int(row.get("_document")) > as_int(row.get("_received"))])
    edited = st.data_editor(frame, column_order=["Фото", "Артикул", "По документу", "Принято", "Ожидается", "Принять сейчас"], hide_index=True, width="stretch", height=540, row_height=92, disabled=["Фото", "Артикул", "По документу", "Принято", "Ожидается"], key=f"warehouse_receiving_editor_{supply.row_id}_{revision}", column_config={"Фото": st.column_config.ImageColumn("Фото", width="medium"), "Принять сейчас": st.column_config.NumberColumn(min_value=0, step=1)})
    st.session_state[draft_key] = {
        str(as_int(row["line_id"])): as_int(row["Принять сейчас"])
        for _, row in edited.iterrows()
    }
    total = int(pd.to_numeric(edited.get("Принять сейчас", pd.Series(dtype=int)), errors="coerce").fillna(0).sum())
    footer = st.columns([1, 2])
    footer[0].metric("К приёмке", f"{total:,} шт.")
    if footer[1].button("Провести приёмку", type="primary", width="stretch", disabled=total <= 0, key=f"warehouse_receive_submit_{supply.row_id}"):
        quantities = {as_int(row["line_id"]): as_int(row["Принять сейчас"]) for _, row in edited.iterrows()}
        command_key = f"warehouse_receiving_command_{supply.row_id}"
        command_id = st.session_state.setdefault(command_key, f"CMD-REC-{uuid.uuid4().hex}")
        result = _safe_action(lambda: service.receive_supply(supply, quantities, command_id=command_id))
        if result:
            st.success(f"Приёмка {result['batch_id']}: {result['sku']} SKU, {result['quantity']} шт.")
            st.session_state[mode_key] = "empty"
            st.session_state.pop(draft_key, None)
            st.session_state.pop(command_key, None)
            st.rerun()


def _read_transfer_excel(data: bytes) -> list[tuple[str, int]]:
    workbook = load_workbook(BytesIO(data), read_only=True, data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(value or "").strip().casefold() for value in rows[0]]
    sku_index = next((i for i, value in enumerate(headers) if value in {"артикул", "sku", "item", "код"}), None)
    qty_index = next((i for i, value in enumerate(headers) if value in {"количество", "qty", "quantity", "шт"}), None)
    missing = []
    if sku_index is None:
        missing.append("Артикул / SKU")
    if qty_index is None:
        missing.append("Количество / Qty")
    if missing:
        raise ValueError("В Excel отсутствуют обязательные колонки: " + ", ".join(missing))
    result: list[tuple[str, int]] = []
    for row in rows[1:]:
        sku = str(row[sku_index] or "").strip() if sku_index < len(row) else ""
        quantity = as_int(row[qty_index]) if qty_index < len(row) else 0
        if sku and quantity > 0:
            result.append((sku, quantity))
    return result


def render_transfer(config: Any) -> None:
    _page_header("Передача в бухгалтерию", "Основной способ — выбрать поставку. Количество сразу ставится максимальным и остаётся редактируемым.")
    _workflow(3)
    service = _require_safe_schema(config)
    if service is None:
        return
    tab_supply, tab_manual, tab_excel = st.tabs(["По поставке — рекомендуемый", "По отдельным SKU", "Из Excel"])
    with tab_supply:
        summaries = [item for item in service.supply_summaries() if item.qty_received > 0]
        if not summaries:
            st.markdown('<div class="wm-empty">Нет поставок с принятым товаром.</div>', unsafe_allow_html=True)
        else:
            options, mapping = _summary_options(summaries)
            supply = mapping[st.selectbox("Выберите поставку", options, key="warehouse_transfer_supply")]
            info = st.columns(4)
            info[0].metric("SKU", supply.sku_total)
            info[1].metric("Принято", supply.qty_received)
            info[2].metric("Уже ожидается", supply.qty_waiting)
            info[3].metric("Статус", supply.status or "—")
            rows = service.supply_products(supply)
            already = service.transferred_by_supply(supply.supply_id)
            records = []
            for row in rows:
                row_id = int(row["id"])
                transferred = as_int(row.get("_transferred")) if service.has_supply_lines else already.get(row_id, 0)
                received = as_int(row.get("_received")); stock = as_int(row.get("Остаток"))
                maximum = min(max(received - transferred, 0), max(stock, 0))
                if maximum <= 0: continue
                records.append({"Фото": _row_photo_data_uri(row, config), "Артикул": row.get("Артикул"), "Принято из поставки": received, "Уже передано": transferred, "Остаток": stock, "Максимум": maximum, "Передать": maximum, "line_id": as_int(row.get("_line_id")) or row_id, "product_row_id": row_id})
            frame = pd.DataFrame(records)
            if frame.empty:
                st.info("По этой поставке больше нет доступного количества для передачи.")
            else:
                st.markdown('<div class="wm-good">Все доступные позиции уже выбраны в максимальном количестве. Уменьшите нужные строки перед проведением.</div>', unsafe_allow_html=True)
                edited = st.data_editor(frame, column_order=["Фото", "Артикул", "Принято из поставки", "Уже передано", "Остаток", "Максимум", "Передать"], hide_index=True, width="stretch", height=550, row_height=92, disabled=["Фото", "Артикул", "Принято из поставки", "Уже передано", "Остаток", "Максимум"], column_config={"Фото": st.column_config.ImageColumn("Фото", width="medium"), "Передать": st.column_config.NumberColumn(min_value=0, step=1)})
                total = int(pd.to_numeric(edited["Передать"], errors="coerce").fillna(0).sum())
                attention = edited.loc[(edited["Остаток"] - edited["Передать"]) <= 15]
                if not attention.empty:
                    with st.expander(f"После передачи требуют внимания: {len(attention)} позиций", expanded=False):
                        show = attention[["Фото", "Артикул", "Остаток", "Передать"]].copy(); show["Останется"] = show["Остаток"] - show["Передать"]
                        st.dataframe(show, hide_index=True, width="stretch", height=300, row_height=82, column_config={"Фото": st.column_config.ImageColumn("Фото", width="small")})
                comment = st.text_input("Комментарий", value=f"Поставка {supply.supply_id}", key=f"warehouse_transfer_comment_{supply.row_id}")
                footer = st.columns([1, 2])
                footer[0].metric("К передаче", f"{total:,} шт.")
                if footer[1].button("Передать выбранное", type="primary", width="stretch", disabled=total <= 0, key=f"warehouse_transfer_submit_{supply.row_id}"):
                    quantities = {as_int(row["line_id"]): as_int(row["Передать"]) for _, row in edited.iterrows()}
                    command_key = f"warehouse_transfer_command_{supply.row_id}"
                    command_id = st.session_state.setdefault(command_key, f"CMD-ACC-{uuid.uuid4().hex}")
                    result = _safe_action(lambda: service.transfer_supply(
                        supply, quantities, comment=comment, command_id=command_id
                    ))
                    if result:
                        st.success(f"Передача {result['batch_id']}: {result['sku']} SKU, {result['quantity']} шт.")
                        st.session_state.pop(command_key, None)
                        st.rerun()
    with tab_manual:
        section = st.segmented_control("Раздел", ["Сувенирка", "Комплектующие"], default="Сувенирка", key="warehouse_manual_transfer_section") or "Сувенирка"
        items = [item for item in service.catalog(section) if item.balance > 0]
        labels = {f"{item.sku} · остаток {item.balance}": item for item in items}
        selected_labels = st.multiselect("Найдите и выберите товары", list(labels), key="warehouse_manual_transfer_items")
        frame = pd.DataFrame([{"Фото": _item_photo_data_uri(labels[label], config), "Артикул": labels[label].sku, "Остаток": labels[label].balance, "Количество": 1, "row_id": labels[label].row_id} for label in selected_labels])
        if frame.empty:
            st.markdown('<div class="wm-empty">Выберите товары для передачи.</div>', unsafe_allow_html=True)
        else:
            edited = st.data_editor(frame, column_order=["Фото", "Артикул", "Остаток", "Количество"], hide_index=True, width="stretch", row_height=92, disabled=["Фото", "Артикул", "Остаток"], column_config={"Фото": st.column_config.ImageColumn("Фото", width="medium"), "Количество": st.column_config.NumberColumn(min_value=1, step=1)})
            comment = st.text_input("Комментарий", key="warehouse_manual_transfer_comment")
            total = int(pd.to_numeric(edited["Количество"], errors="coerce").fillna(0).sum())
            if st.button("Провести передачу", type="primary", width="stretch", disabled=total <= 0, key="warehouse_manual_transfer_submit"):
                quantities = {as_int(row["row_id"]): as_int(row["Количество"]) for _, row in edited.iterrows()}
                command_id = st.session_state.setdefault("warehouse_manual_transfer_command", f"CMD-ACC-{uuid.uuid4().hex}")
                result = _safe_action(lambda: service.manual_operation(
                    operation_type="Передача в бухгалтерию", section=section, quantities=quantities,
                    comment=comment, command_id=command_id,
                ))
                if result:
                    st.success(f"Передача {result['batch_id']}: {result['quantity']} шт.")
                    st.session_state.pop("warehouse_manual_transfer_command", None)
                    st.rerun()
    with tab_excel:
        st.caption("Excel должен содержать столбцы SKU/Артикул и Количество.")
        uploaded = st.file_uploader("Файл передачи", type=["xlsx", "xlsm"], key="warehouse_transfer_excel")
        section = st.segmented_control("Раздел для Excel", ["Сувенирка", "Комплектующие"], default="Сувенирка", key="warehouse_transfer_excel_section") or "Сувенирка"
        if uploaded is not None:
            try:
                requested = _read_transfer_excel(uploaded.getvalue())
            except ValueError as exc:
                st.error(str(exc))
                requested = []
            catalog = {item.sku.casefold(): item for item in service.catalog(section)}
            records, missing = [], []
            for sku, quantity in requested:
                item = catalog.get(sku.casefold())
                if item is None:
                    missing.append(sku)
                    continue
                records.append({
                    "Фото": _item_photo_data_uri(item, config),
                    "Артикул": item.sku,
                    "Остаток": item.balance,
                    "Количество": quantity,
                    "Ошибка": "Превышение остатка" if quantity > item.balance else "",
                    "row_id": item.row_id,
                })
            edited = st.data_editor(
                pd.DataFrame(records),
                column_order=["Фото", "Артикул", "Остаток", "Количество", "Ошибка"],
                hide_index=True, width="stretch", row_height=92,
                disabled=["Фото", "Артикул", "Остаток", "Ошибка"],
                column_config={
                    "Фото": st.column_config.ImageColumn("Фото", width="medium"),
                    "Количество": st.column_config.NumberColumn(min_value=1, step=1),
                },
            )
            if missing:
                st.warning("Не найдены: " + ", ".join(missing[:30]))
            has_errors = bool(records) and any(as_int(row["Количество"]) > as_int(row["Остаток"]) for _, row in edited.iterrows())
            if has_errors:
                st.error("Исправьте строки, где количество превышает доступный остаток.")
            if records and st.button("Провести Excel-передачу", type="primary", width="stretch", disabled=has_errors):
                quantities = {as_int(row["row_id"]): as_int(row["Количество"]) for _, row in edited.iterrows()}
                command_id = st.session_state.setdefault("warehouse_excel_transfer_command", f"CMD-ACC-{uuid.uuid4().hex}")
                result = _safe_action(lambda: service.manual_operation(
                    operation_type="Передача в бухгалтерию", section=section, quantities=quantities,
                    comment=f"Excel {uploaded.name}", command_id=command_id,
                ))
                if result:
                    st.success(f"Передача {result['batch_id']}: {result['quantity']} шт.")
                    st.session_state.pop("warehouse_excel_transfer_command", None)
                    st.rerun()


def render_operations(config: Any) -> None:
    from src.warehouse import normalize_operations

    _page_header("История операций", "Фильтруйте движения склада и создавайте обратную корректировку без удаления истории.")
    service = _service(config)
    raw = service.client.list_rows(config.operations_table_id)
    frame = normalize_operations(raw)
    filters = st.columns([1, 1, 2])
    operation_types = ["Все", *sorted(frame["Тип операции"].dropna().astype(str).unique().tolist())] if not frame.empty else ["Все"]
    selected_type = filters[0].selectbox("Тип", operation_types, key="warehouse_operations_type")
    period = filters[1].selectbox("Показать", ["Все", "Последние 100", "Последние 500"], key="warehouse_operations_limit")
    query = filters[2].text_input("Поиск", placeholder="SKU, Batch ID или поставка", key="warehouse_operations_query").strip().casefold()
    current = frame.copy()
    if selected_type != "Все": current = current.loc[current["Тип операции"] == selected_type]
    if query: current = current.loc[current.astype(str).apply(lambda row: query in " ".join(row).casefold(), axis=1)]
    if period != "Все": current = current.head(int(period.split()[1]))
    metrics = st.columns(4)
    metrics[0].metric("Операций", len(current))
    metrics[1].metric("Приход", int(current.loc[current["Тип операции"].eq("Приход"), "Количество"].sum()) if not current.empty and "Количество" in current else 0)
    metrics[2].metric("Передано", int(current.loc[current["Тип операции"].eq("Передача в бухгалтерию"), "Количество"].sum()) if not current.empty and "Количество" in current else 0)
    metrics[3].metric("Корректировки", int(current["Тип операции"].isin(["Возврат", "Расход", "Корректировка"]).sum()) if not current.empty else 0)
    if "Дата" in current.columns: current["Дата"] = current["Дата"].dt.strftime("%d.%m.%Y %H:%M")
    st.dataframe(current, width="stretch", hide_index=True, height=590)
    with st.expander("Создать корректировку операции", expanded=False):
        selectable = [row for row in raw if as_int(row.get("Количество")) > 0]
        labels = {f"{row.get('Batch ID') or row.get('id')} · {select_text(row.get('Тип операции'))} · {row.get('Операция') or ''}": row for row in selectable}
        if not labels:
            st.info("Нет операций для корректировки.")
        else:
            label = st.selectbox("Операция", list(labels), key="warehouse_correction_operation")
            operation = labels[label]
            available = service.correction_available(operation)
            st.caption(f"Доступно к корректировке: {available} шт.")
            if available <= 0:
                st.info("Операция уже скорректирована полностью.")
            else:
                quantity = st.number_input("Количество для отмены", min_value=1, max_value=available, value=available)
                comment = st.text_input("Причина корректировки *", key="warehouse_correction_comment")
                confirm = st.checkbox("Подтверждаю создание обратной операции", key="warehouse_correction_confirm")
                if st.button("Создать корректировку", type="primary", disabled=not confirm or not comment.strip()):
                    command_id = st.session_state.setdefault("warehouse_correction_command", f"CMD-COR-{uuid.uuid4().hex}")
                    result = _safe_action(lambda: service.correct_operation(
                        operation, quantity=int(quantity), comment=comment, command_id=command_id
                    ))
                    if result:
                        st.success(f"Корректировка создана: {result['batch_id']}; осталось {result['remaining']} шт.")
                        st.session_state.pop("warehouse_correction_command", None)
                        st.rerun()


def render_maintenance(config: Any) -> None:
    _page_header("Обслуживание", "Схема Baserow, миграция позиций поставок и контроль качества складских данных.")
    service = _service(config)
    resolved = service.config
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Подключение")
        st.write(
            {
                "Baserow": resolved.base_url,
                "Сувенирка": resolved.souvenirs_table_id,
                "Комплектующие": resolved.components_table_id,
                "Операции": resolved.operations_table_id,
                "Поставки": resolved.supplies_table_id,
                "Позиции поставок": int(getattr(resolved, "supply_lines_table_id", 0) or 0) or "не создана",
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
                '<div class="wm-danger">Рабочие операции заблокированы до создания таблицы «Позиции поставок».</div>',
                unsafe_allow_html=True,
            )

    with st.expander("Создать / проверить таблицу «Позиции поставок»", expanded=not service.has_supply_lines):
        st.caption(
            "Пароль используется только для получения короткого JWT и не сохраняется. "
            "Миграция повторяемая: уже созданные строки не дублируются."
        )
        default_email = str(getattr(resolved, "email", "") or "")
        email = st.text_input("Email Baserow", value=default_email, key="warehouse_schema_email")
        password = st.text_input("Пароль Baserow", type="password", key="warehouse_schema_password")
        confirm = st.checkbox(
            "Подтверждаю создание таблицы, служебных полей и миграцию существующих поставок",
            key="warehouse_schema_confirm",
        )
        if st.button(
            "Создать и мигрировать",
            type="primary",
            disabled=not confirm or not email.strip() or not password,
            key="warehouse_schema_run",
        ):
            try:
                with st.spinner("Создаём схему и переносим существующие поставки..."):
                    manager = BaserowSchemaManager(resolved.base_url, email, password)
                    report = manager.ensure_and_migrate(
                        database_id=int(resolved.database_id),
                        souvenirs_table_id=int(resolved.souvenirs_table_id),
                        components_table_id=int(resolved.components_table_id),
                        operations_table_id=int(resolved.operations_table_id),
                        supplies_table_id=int(resolved.supplies_table_id),
                    )
                st.session_state["warehouse_supply_lines_table_id"] = report.table_id
                st.session_state["warehouse_schema_report"] = report.to_dict()
                _clear_cache(photos=True)
                st.success(
                    f"Таблица готова: ID {report.table_id}. Перенесено строк: {report.migrated_lines}; "
                    f"пропущено существующих: {report.skipped_lines}."
                )
                st.rerun()
            except WarehouseSchemaError as exc:
                st.error(str(exc))

        schema_report = st.session_state.get("warehouse_schema_report")
        if schema_report:
            st.json(schema_report)
            st.download_button(
                "Скачать отчёт миграции",
                data=json.dumps(schema_report, ensure_ascii=False, indent=2),
                file_name="warehouse_supply_lines_migration_2.4.0.json",
                mime="application/json",
            )
            if schema_report.get("ambiguous_skus"):
                st.warning(
                    "Требуют ручной проверки: "
                    + ", ".join(schema_report["ambiguous_skus"][:40])
                )

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
        with st.expander(f"Архивные карточки: {len(report.get('inactive', []))}"):
            st.write(report.get("inactive") or "Не найдено")
        with st.expander(f"Незавершённые документы: {len(report.get('unfinished_operations', []))}"):
            st.write(report.get("unfinished_operations") or "Не найдено")
        with st.expander(f"Позиции, требующие проверки: {len(report.get('ambiguous_supply_lines', []))}"):
            st.write(report.get("ambiguous_supply_lines") or "Не найдено")


def render_supply_hub(config: Any) -> None:
    current = st.segmented_control(
        "Раздел поставок",
        list(SUPPLY_WORKSPACES),
        default="Реестр",
        key="warehouse_supply_workspace",
        label_visibility="collapsed",
    ) or "Реестр"
    if current == "Реестр":
        render_supplies(config)
    elif current == "Новая поставка":
        render_new_supply(config)
    else:
        render_receiving(config)


def render_history_hub(config: Any) -> None:
    current = st.segmented_control(
        "История и сервис",
        list(HISTORY_WORKSPACES),
        default="Операции",
        key="warehouse_history_workspace",
        label_visibility="collapsed",
    ) or "Операции"
    if current == "Операции":
        render_operations(config)
    else:
        render_maintenance(config)


def render_warehouse_workspace(config: Any, selected_metal_groups: Iterable[str]) -> None:
    # загружается только выбранный раздел; остальные рабочие пространства не выполняются.
    """Render one lazy, task-oriented warehouse workspace inside Analitika."""
    st.markdown(WAREHOUSE_MANAGEMENT_CSS, unsafe_allow_html=True)
    st.markdown(
        '<div class="wm-shell">'
        '<div class="wm-kicker">Princess Warehouse Online</div>'
        '<div class="wm-shell-title">Сувениры и касты на складе</div>'
        '<div class="wm-shell-copy">Фото товаров, поставки, приёмка, передача в бухгалтерию и история операций — в одном рабочем блоке.</div>'
        '</div>',
        unsafe_allow_html=True,
    )
    st.session_state.setdefault("warehouse_workspace", "Главная")
    current = st.segmented_control(
        "Раздел склада",
        list(WORKSPACES),
        default="Главная",
        key="warehouse_workspace",
        label_visibility="collapsed",
    ) or "Главная"
    toolbar = st.columns([1, 5])
    if toolbar[0].button("Обновить данные", key="warehouse_workspace_refresh", width="stretch"):
        _clear_cache(); st.rerun()
    toolbar[1].markdown('<div class="wm-toolbar-note">Загружается только выбранный раздел. Фото и данные Baserow кэшируются и не перегружают остальные модули Analitika.</div>', unsafe_allow_html=True)
    if current == "Главная":
        render_overview(config, selected_metal_groups)
    elif current == "Товары":
        render_catalog(config)
    elif current == "Поставки":
        render_supply_hub(config)
    elif current == "Передача":
        render_transfer(config)
    else:
        render_history_hub(config)


