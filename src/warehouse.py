from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from html import escape
import os
import re
from typing import Any, Iterable
from urllib.parse import urljoin

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

DEFAULT_MINIMUM_STOCK = 10
EARLY_WARNING_STOCK = 15
DEFAULT_TABLE_IDS = {
    "souvenirs": 642,
    "components": 643,
    "operations": 644,
    "supplies": 645,
}
WAREHOUSE_SECTIONS: tuple[tuple[str, str, str], ...] = (
    ("Обзор", "warehouse-overview", "📊 Обзор"),
    ("Сувениры", "warehouse-souvenirs", "🎁 Сувениры"),
    ("Касты", "warehouse-components", "🧩 Касты"),
    ("Требует внимания", "warehouse-attention", "⚠️ Требует внимания"),
    ("Движение", "warehouse-movement", "↔️ Движение"),
    ("Поставки", "warehouse-supplies", "📦 Поставки"),
)

LOCKED_CHART_CONFIG = {
    "displayModeBar": False,
    "displaylogo": False,
    "scrollZoom": False,
    "doubleClick": False,
    "showTips": False,
    "editable": False,
    "staticPlot": False,
    "responsive": True,
    "showAxisDragHandles": False,
    "showAxisRangeEntryBoxes": False,
}

WAREHOUSE_CSS = """
<style>
.warehouse-header {
  border:1px solid #e9e4dc; border-radius:20px; padding:22px 24px; margin:8px 0 18px;
  background:linear-gradient(135deg,rgba(255,255,255,.98),rgba(247,241,231,.94));
  box-shadow:0 12px 34px rgba(34,24,9,.06);
}
.warehouse-header-kicker { color:#b7893f; font-size:12px; font-weight:800; letter-spacing:.12em; text-transform:uppercase; }
.warehouse-header h2 { font-family:Georgia,serif; margin:4px 0 6px; font-size:36px; color:#171411; }
.warehouse-header p { color:#6c6c6c; margin:0; }
.wh-metric { border:1px solid #e9e4dc; border-radius:16px; padding:17px 18px; min-height:112px; background:#fff; box-shadow:0 8px 25px rgba(34,24,9,.045); }
.wh-metric-label { color:#7c7469; font-size:12px; letter-spacing:.055em; text-transform:uppercase; font-weight:750; }
.wh-metric-value { margin-top:8px; font-size:30px; line-height:1.05; font-weight:800; color:#171411; }
.wh-metric-note { color:#777067; font-size:12px; margin-top:7px; }
.wh-alert { border-left:4px solid #b7893f; background:#fffaf1; padding:13px 15px; border-radius:10px; margin:8px 0 14px; }
.wh-stock-card { border:1px solid #e9e4dc; border-radius:16px; padding:13px; background:#fff; box-shadow:0 8px 22px rgba(34,24,9,.04); min-height:178px; margin-bottom:12px; }
.wh-stock-card .sku { font-size:17px; font-weight:800; color:#171411; }
.wh-stock-card .meta { color:#6c6c6c; font-size:13px; line-height:1.45; }
.wh-stock-card .balance { margin-top:8px; font-weight:800; }
.wh-status-ok { color:#2f6f45; } .wh-status-warning { color:#9b6500; } .wh-status-critical { color:#a82f2f; }
.wh-photo-placeholder { min-height:180px; border-radius:14px; border:1px dashed #d8c8ad; background:#faf7f2; display:flex; align-items:center; justify-content:center; color:#8b8174; margin-bottom:10px; }
@media (max-width:900px) {
  .warehouse-header { padding:20px; border-radius:16px; }
  .warehouse-header h2 { font-size:31px; }
  .wh-stock-card { min-height:156px; }
  .wh-photo-placeholder { min-height:145px; }
}
@media (max-width:640px) {
  .warehouse-header { padding:17px 16px; }
  .warehouse-header h2 { font-size:27px; line-height:1.12; }
  .warehouse-header p { font-size:13px; line-height:1.45; }
  .wh-metric { min-height:96px; padding:14px; }
  .wh-metric-value { font-size:25px; }
  .wh-stock-card { min-height:auto; }
  .wh-photo-placeholder { min-height:120px; }
}
</style>
"""


class WarehouseConfigError(RuntimeError):
    pass


class WarehouseApiError(RuntimeError):
    pass


@dataclass(frozen=True)
class WarehouseConfig:
    base_url: str
    token: str
    souvenirs_table_id: int
    components_table_id: int
    operations_table_id: int
    supplies_table_id: int

    @classmethod
    def load(cls) -> "WarehouseConfig":
        section: dict[str, Any] = {}
        try:
            if "baserow" in st.secrets:
                section = dict(st.secrets["baserow"])
        except Exception:
            section = {}
        base_url = str(section.get("url") or os.getenv("BASEROW_URL") or "").strip().rstrip("/")
        token = str(section.get("token") or os.getenv("BASEROW_TOKEN") or "").strip()

        def table_id(name: str) -> int:
            raw = section.get(f"{name}_table_id") or os.getenv(f"BASEROW_{name.upper()}_TABLE_ID")
            return int(raw) if raw not in (None, "") else DEFAULT_TABLE_IDS[name]

        if not base_url or not token:
            raise WarehouseConfigError("Не настроено read-only подключение к Baserow.")
        return cls(base_url, token, table_id("souvenirs"), table_id("components"), table_id("operations"), table_id("supplies"))


@dataclass
class WarehouseBundle:
    souvenirs: pd.DataFrame
    components: pd.DataFrame
    operations: pd.DataFrame
    supplies: pd.DataFrame
    loaded_at: datetime


def _headers(token: str) -> dict[str, str]:
    return {"Authorization": f"Token {token}", "Accept": "application/json", "User-Agent": "Princess-Analitika-Warehouse/1.2"}


@st.cache_data(ttl=60, max_entries=16, show_spinner=False)
def fetch_table_rows(base_url: str, token: str, table_id: int) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    page = 1
    session = requests.Session()
    session.headers.update(_headers(token))
    while True:
        try:
            response = session.get(
                f"{base_url}/api/database/rows/table/{int(table_id)}/",
                params={"user_field_names": "true", "size": 200, "page": page},
                timeout=35,
            )
        except requests.RequestException as exc:
            raise WarehouseApiError("Не удалось подключиться к Baserow.") from exc
        if response.status_code in (401, 403):
            raise WarehouseApiError("Baserow отклонил read-only токен или право Read для таблицы.")
        if not response.ok:
            raise WarehouseApiError(f"Baserow вернул HTTP {response.status_code} для таблицы {table_id}.")
        payload = response.json()
        rows.extend(payload.get("results", []))
        if not payload.get("next"):
            break
        page += 1
    return rows


@st.cache_data(ttl=1800, max_entries=256, show_spinner=False)
def fetch_image_bytes(url: str, token: str) -> bytes | None:
    if not url:
        return None
    for headers in (_headers(token), {"User-Agent": "Princess-Analitika-Warehouse/1.2"}):
        try:
            response = requests.get(url, headers=headers, timeout=20)
            if response.ok and response.content:
                return response.content
        except requests.RequestException:
            continue
    return None


def _scalar(value: Any) -> Any:
    if isinstance(value, dict):
        for key in ("value", "name", "text"):
            if key in value:
                return _scalar(value[key])
        return ""
    return value


def text_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        values = [text_value(item).strip() for item in value]
        return "; ".join(dict.fromkeys(item for item in values if item))
    return str(_scalar(value) or "").strip()


def relation_values(value: Any) -> list[str]:
    values = value if isinstance(value, list) else [value] if value else []
    result: list[str] = []
    for item in values:
        text = text_value(item)
        if text and text not in result:
            result.append(text)
    return result


def number_value(value: Any, default: float = 0.0) -> float:
    value = _scalar(value)
    if value in (None, ""):
        return default
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    clean = str(value).strip().replace("\u00a0", "").replace(" ", "")
    if clean.count(",") == 1 and "." not in clean:
        clean = clean.replace(",", ".")
    clean = re.sub(r"[^0-9.\-]", "", clean)
    try:
        return float(clean)
    except ValueError:
        return default


def datetime_value(value: Any) -> pd.Timestamp | pd.NaT:
    parsed = pd.to_datetime(_scalar(value), errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        return pd.NaT
    if getattr(parsed, "tzinfo", None) is not None:
        parsed = parsed.tz_convert(None)
    return parsed


def first_existing(row: dict[str, Any], names: Iterable[str]) -> Any:
    for name in names:
        if name in row and row.get(name) not in (None, "", []):
            return row.get(name)
    return None


def file_url(value: Any, base_url: str) -> str:
    files = value if isinstance(value, list) else [value] if value else []
    if not files or not isinstance(files[0], dict):
        return ""
    item = files[0]
    url = str(item.get("url") or "").strip()
    if not url:
        for name in ("large", "card_cover", "small", "tiny"):
            candidate = (item.get("thumbnails") or {}).get(name)
            if isinstance(candidate, dict) and candidate.get("url"):
                url = str(candidate["url"])
                break
    return urljoin(base_url + "/", url) if url else ""


def _normalize_stone(value: str) -> str:
    return re.sub(r"\bLAPIS\s+LAZULI\b", "Lapis Lazurite", value, flags=re.I)


def stock_status(balance: float, minimum: float) -> str:
    minimum = minimum if minimum > 0 else DEFAULT_MINIMUM_STOCK
    if balance <= 0:
        return "Нет в наличии"
    if balance <= minimum:
        return "Ниже минимума"
    if balance <= EARLY_WARNING_STOCK:
        return "Заканчивается"
    return "В наличии"


def normalize_inventory_rows(rows: list[dict[str, Any]], section: str, base_url: str) -> pd.DataFrame:
    records: list[dict[str, Any]] = []
    for row in rows:
        sku = text_value(first_existing(row, ("Артикул", "SKU", "Наименование", "Название")))
        if not sku:
            continue
        balance = number_value(first_existing(row, ("Остаток", "Текущий остаток")))
        minimum = number_value(first_existing(row, ("Минимальный остаток", "Минимум")), DEFAULT_MINIMUM_STOCK)
        if minimum <= 0:
            minimum = DEFAULT_MINIMUM_STOCK
        stone = _normalize_stone(text_value(first_existing(row, ("Камень", "Вставка"))))
        records.append({
            "Раздел": section,
            "Фото": file_url(first_existing(row, ("Фото", "Изображение")), base_url),
            "Артикул": sku,
            "Категория": text_value(first_existing(row, ("Категория", "Тип", "Вид"))),
            "Материал": text_value(first_existing(row, ("Материал", "Металл"))),
            "Камень": stone,
            "Цвет": text_value(row.get("Цвет")),
            "Коробки": text_value(first_existing(row, ("Номера коробок", "Коробка", "Коробки"))),
            "Описание": text_value(first_existing(row, ("Описание", "Комментарий", "Notes"))),
            "Поставки": "; ".join(relation_values(first_existing(row, ("Поставки", "Поставка")))),
            "Остаток": int(round(balance)),
            "Минимальный остаток": int(round(minimum)),
            "По документу": int(round(number_value(row.get("По документу, шт.")))),
            "Получено": int(round(number_value(row.get("Получено по поставке, шт.")))),
            "Ожидается": int(round(number_value(row.get("Ожидается, шт.")))),
            "Статус": stock_status(balance, minimum),
        })
    frame = pd.DataFrame.from_records(records)
    if frame.empty:
        return pd.DataFrame(columns=["Раздел", "Фото", "Артикул", "Категория", "Материал", "Камень", "Цвет", "Коробки", "Описание", "Поставки", "Остаток", "Минимальный остаток", "По документу", "Получено", "Ожидается", "Статус"])
    order = {"Нет в наличии": 0, "Ниже минимума": 1, "Заканчивается": 2, "В наличии": 3}
    frame["_status_order"] = frame["Статус"].map(order).fillna(9)
    return frame.sort_values(["_status_order", "Остаток", "Артикул"]).drop(columns="_status_order").reset_index(drop=True)


def normalize_operations(rows: list[dict[str, Any]]) -> pd.DataFrame:
    records: list[dict[str, Any]] = []
    for row in rows:
        operation = text_value(first_existing(row, ("Тип операции", "Операция")))
        quantity = number_value(row.get("Количество"))
        change = number_value(row.get("Изменение"), float("nan"))
        if pd.isna(change):
            low = operation.casefold()
            change = abs(quantity) if "приход" in low else -abs(quantity) if ("передач" in low or "расход" in low) else quantity
        souvenirs = relation_values(row.get("Товар сувенирки"))
        components = relation_values(row.get("Комплектующее"))
        records.append({
            "Дата": datetime_value(first_existing(row, ("Дата и время", "Дата", "Created on", "Дата создания"))),
            "Тип операции": operation,
            "Раздел": text_value(row.get("Раздел")) or ("Сувенирка" if souvenirs else "Комплектующие" if components else ""),
            "SKU": "; ".join(souvenirs + [item for item in components if item not in souvenirs]),
            "Количество": int(round(abs(quantity))),
            "Изменение": int(round(change)),
            "Batch ID": text_value(row.get("Batch ID")),
            "Ответственный": text_value(first_existing(row, ("Ответственный", "Сотрудник", "Пользователь"))),
            "Комментарий": text_value(row.get("Комментарий")),
        })
    frame = pd.DataFrame.from_records(records)
    return frame.sort_values("Дата", ascending=False, na_position="last").reset_index(drop=True) if not frame.empty else pd.DataFrame(columns=["Дата", "Тип операции", "Раздел", "SKU", "Количество", "Изменение", "Batch ID", "Ответственный", "Комментарий"])


def normalize_supplies(rows: list[dict[str, Any]]) -> pd.DataFrame:
    records: list[dict[str, Any]] = []
    for row in rows:
        number = text_value(first_existing(row, ("№ поставки", "Поставка", "Название")))
        if not number:
            continue
        records.append({
            "Поставка": number,
            "Дата": datetime_value(first_existing(row, ("Дата", "Дата создания"))),
            "Поставщик": text_value(row.get("Поставщик")),
            "Invoice": text_value(row.get("Invoice")),
            "Статус": text_value(row.get("Статус")),
            "SKU": int(round(number_value(first_existing(row, ("SKU", "Количество SKU"))))),
            "По документу": int(round(number_value(first_existing(row, ("По документу", "Количество по документу"))))),
            "Получено": int(round(number_value(row.get("Получено")))),
            "Ожидается": int(round(number_value(row.get("Ожидается")))),
            "Комментарий": text_value(row.get("Комментарий")),
        })
    frame = pd.DataFrame.from_records(records)
    return frame.sort_values("Дата", ascending=False, na_position="last").reset_index(drop=True) if not frame.empty else pd.DataFrame(columns=["Поставка", "Дата", "Поставщик", "Invoice", "Статус", "SKU", "По документу", "Получено", "Ожидается", "Комментарий"])


def load_bundle(config: WarehouseConfig) -> WarehouseBundle:
    souvenirs = normalize_inventory_rows(fetch_table_rows(config.base_url, config.token, config.souvenirs_table_id), "Сувенирка", config.base_url)
    components = normalize_inventory_rows(fetch_table_rows(config.base_url, config.token, config.components_table_id), "Комплектующие", config.base_url)
    operations = normalize_operations(fetch_table_rows(config.base_url, config.token, config.operations_table_id))
    supplies = normalize_supplies(fetch_table_rows(config.base_url, config.token, config.supplies_table_id))
    return WarehouseBundle(souvenirs, components, operations, supplies, datetime.now())


def metric_card(label: str, value: str, note: str = "") -> None:
    st.markdown(f'<div class="wh-metric"><div class="wh-metric-label">{escape(label)}</div><div class="wh-metric-value">{escape(value)}</div><div class="wh-metric-note">{escape(note)}</div></div>', unsafe_allow_html=True)


def locked_chart(fig: go.Figure, key: str) -> None:
    fig.update_layout(dragmode=False, clickmode="event", hovermode="closest", legend_itemclick=False, legend_itemdoubleclick=False, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(family="Inter, Arial", color="#171411"))
    fig.update_xaxes(fixedrange=True, automargin=True)
    fig.update_yaxes(fixedrange=True, automargin=True)
    st.plotly_chart(fig, width="stretch", key=key, config=LOCKED_CHART_CONFIG)


def category_chart(frame: pd.DataFrame, title: str) -> go.Figure:
    data = frame.loc[frame["Остаток"] > 0].copy()
    if data.empty:
        return go.Figure().update_layout(title=title)
    data["Категория"] = data["Категория"].replace("", "Без категории")
    grouped = data.groupby("Категория", as_index=False)["Остаток"].sum().sort_values("Остаток").tail(12)
    fig = go.Figure(go.Bar(x=grouped["Остаток"], y=grouped["Категория"], orientation="h", marker_color="#b7893f", text=grouped["Остаток"], textposition="outside", cliponaxis=False))
    fig.update_layout(title=title, height=max(350, len(grouped) * 38 + 120), margin=dict(l=20, r=90, t=55, b=38), showlegend=False)
    return fig


def movement_window(operations: pd.DataFrame, days: int) -> pd.DataFrame:
    if operations.empty:
        return operations.copy()
    threshold = pd.Timestamp.now().normalize() - pd.Timedelta(days=max(days - 1, 0))
    return operations.loc[operations["Дата"].notna() & (operations["Дата"] >= threshold)].copy()


def operation_totals(operations: pd.DataFrame, days: int = 30) -> tuple[int, int]:
    current = movement_window(operations, days)
    if current.empty:
        return 0, 0
    return int(current.loc[current["Изменение"] > 0, "Изменение"].sum()), int(abs(current.loc[current["Изменение"] < 0, "Изменение"].sum()))


def movement_chart(operations: pd.DataFrame, days: int) -> go.Figure:
    data = movement_window(operations, days)
    if data.empty:
        return go.Figure().update_layout(title=f"Движение за {days} дней")
    data["День"] = data["Дата"].dt.normalize()
    data["Приход"] = data["Изменение"].clip(lower=0)
    data["Передано"] = -data["Изменение"].clip(upper=0)
    grouped = data.groupby("День", as_index=False)[["Приход", "Передано"]].sum()
    fig = go.Figure()
    fig.add_bar(x=grouped["День"], y=grouped["Приход"], name="Приход", marker_color="#5f876d", text=grouped["Приход"], textposition="outside", cliponaxis=False)
    fig.add_bar(x=grouped["День"], y=grouped["Передано"], name="Передано", marker_color="#b7893f", text=grouped["Передано"], textposition="outside", cliponaxis=False)
    fig.update_layout(title=f"Приход и передача за {days} дней", barmode="group", height=330, margin=dict(l=20, r=55, t=55, b=48), legend=dict(orientation="h", y=1.12))
    return fig


def apply_filters(frame: pd.DataFrame, prefix: str) -> pd.DataFrame:
    if frame.empty:
        return frame.copy()
    first, second, third = st.columns(3)
    with first:
        categories = st.multiselect("Категория", sorted(x for x in frame["Категория"].unique() if x), key=f"{prefix}_category")
    with second:
        materials = st.multiselect("Материал", sorted(x for x in frame["Материал"].unique() if x), key=f"{prefix}_material")
    with third:
        statuses = st.multiselect("Статус", ["В наличии", "Заканчивается", "Ниже минимума", "Нет в наличии"], default=["В наличии", "Заканчивается", "Ниже минимума"], key=f"{prefix}_status")
    search = st.text_input("Артикул, коробка, камень или текст", key=f"{prefix}_search")
    result = frame.copy()
    if categories:
        result = result[result["Категория"].isin(categories)]
    if materials:
        result = result[result["Материал"].isin(materials)]
    result = result[result["Статус"].isin(statuses)] if statuses else result.iloc[0:0]
    if search:
        haystack = result[["Артикул", "Категория", "Материал", "Камень", "Цвет", "Коробки", "Описание"]].fillna("").astype(str).agg(" ".join, axis=1).str.casefold()
        result = result[haystack.str.contains(re.escape(search.casefold()), regex=True)]
    return result.reset_index(drop=True)


def render_inventory_table(frame: pd.DataFrame, key: str) -> None:
    columns = ["Фото", "Артикул", "Категория", "Материал", "Камень", "Цвет", "Коробки", "Остаток", "Минимальный остаток", "Статус", "Поставки"]
    st.dataframe(frame[[c for c in columns if c in frame]], width="stretch", hide_index=True, key=key, column_config={"Фото": st.column_config.ImageColumn("Фото", width="medium"), "Остаток": st.column_config.NumberColumn(format="%d шт."), "Минимальный остаток": st.column_config.NumberColumn("Минимум", format="%d шт.")})


def render_inventory_cards(frame: pd.DataFrame, config: WarehouseConfig, key: str) -> None:
    if frame.empty:
        st.info("По выбранным фильтрам позиций нет.")
        return
    page_size = st.segmented_control("Карточек на странице", [6, 12, 18], default=12, key=f"{key}_page_size") or 12
    page_count = max(1, (len(frame) + int(page_size) - 1) // int(page_size))
    page = st.number_input("Страница", 1, page_count, 1, key=f"{key}_page")
    current = frame.iloc[(int(page) - 1) * int(page_size): int(page) * int(page_size)]
    for start in range(0, len(current), 3):
        cols = st.columns(3)
        for col, (_, item) in zip(cols, current.iloc[start:start + 3].iterrows()):
            with col:
                image = fetch_image_bytes(str(item.get("Фото", "")), config.token)
                if image:
                    st.image(image, width="stretch")
                else:
                    st.markdown('<div class="wh-photo-placeholder">Нет фотографии</div>', unsafe_allow_html=True)
                meta = " · ".join(str(item.get(x, "")) for x in ("Категория", "Материал", "Камень", "Цвет") if str(item.get(x, "")).strip()) or "Характеристики не указаны"
                status = str(item.get("Статус", ""))
                css = "wh-status-ok" if status == "В наличии" else "wh-status-warning" if status in ("Заканчивается", "Ниже минимума") else "wh-status-critical"
                st.markdown(f'<div class="wh-stock-card"><div class="sku">{escape(str(item.get("Артикул", "")))}</div><div class="meta">{escape(meta)}</div><div class="balance {css}">{escape(status)} · {int(item.get("Остаток", 0)):,} шт.</div></div>', unsafe_allow_html=True)


def render_inventory_section(frame: pd.DataFrame, title: str, prefix: str, config: WarehouseConfig) -> None:
    st.markdown(f'<div class="wh-alert"><b>{escape(title)}</b><br>Показываются позиции, принятые в Baserow и находящиеся на складе.</div>', unsafe_allow_html=True)
    filtered = apply_filters(frame, prefix)
    c1, c2, c3 = st.columns(3)
    with c1: metric_card("Найдено SKU", f"{len(filtered):,}")
    with c2: metric_card("Единиц", f"{int(filtered['Остаток'].clip(lower=0).sum()) if not filtered.empty else 0:,}")
    with c3: metric_card("Остаток ≤ 15", f"{int(((filtered['Остаток'] > 0) & (filtered['Остаток'] <= EARLY_WARNING_STOCK)).sum()) if not filtered.empty else 0:,}")
    view = st.segmented_control("Вид", ["Карточки", "Таблица"], default="Карточки", key=f"{prefix}_view") or "Карточки"
    render_inventory_table(filtered, f"{prefix}_table") if view == "Таблица" else render_inventory_cards(filtered, config, prefix)


def render_overview(bundle: WarehouseBundle) -> None:
    incoming, outgoing = operation_totals(bundle.operations, 30)
    c1, c2, c3, c4 = st.columns(4)
    with c1: metric_card("Сувениры", f"{int(bundle.souvenirs['Остаток'].clip(lower=0).sum()):,} шт.")
    with c2: metric_card("Касты", f"{int(bundle.components['Остаток'].clip(lower=0).sum()):,} шт.")
    with c3: metric_card("SKU", f"{len(bundle.souvenirs) + len(bundle.components):,}")
    with c4: metric_card("Требует внимания", f"{int((bundle.souvenirs['Остаток'] <= EARLY_WARNING_STOCK).sum()) + int((bundle.components['Остаток'] <= EARLY_WARNING_STOCK).sum()):,}")
    second = st.columns(4)
    with second[0]: metric_card("Ниже минимума", f"{int(((bundle.souvenirs['Остаток'] > 0) & (bundle.souvenirs['Остаток'] <= bundle.souvenirs['Минимальный остаток'])).sum()) + int(((bundle.components['Остаток'] > 0) & (bundle.components['Остаток'] <= bundle.components['Минимальный остаток'])).sum()):,}")
    with second[1]: metric_card("Нет в наличии", f"{int((bundle.souvenirs['Остаток'] <= 0).sum()) + int((bundle.components['Остаток'] <= 0).sum()):,}")
    with second[2]: metric_card("Принято за 30 дней", f"{incoming:,} шт.")
    with second[3]: metric_card("Передано за 30 дней", f"{outgoing:,} шт.")
    charts = st.columns(2)
    with charts[0]: locked_chart(category_chart(bundle.souvenirs, "Сувениры по категориям"), "warehouse_souvenir_categories")
    with charts[1]: locked_chart(category_chart(bundle.components, "Касты по категориям"), "warehouse_component_categories")
    locked_chart(movement_chart(bundle.operations, 30), "warehouse_movement_overview")


def render_attention(bundle: WarehouseBundle) -> None:
    data = pd.concat([bundle.souvenirs, bundle.components], ignore_index=True)
    selected = st.segmented_control("Проблема", ["Заканчивается", "Ниже минимума", "Нет в наличии", "Без фото", "Без категории"], default="Заканчивается", key="warehouse_attention_type") or "Заканчивается"
    if selected == "Заканчивается":
        current = data[(data["Остаток"] > data["Минимальный остаток"]) & (data["Остаток"] <= EARLY_WARNING_STOCK)]
    elif selected == "Ниже минимума":
        current = data[(data["Остаток"] > 0) & (data["Остаток"] <= data["Минимальный остаток"])]
    elif selected == "Нет в наличии":
        current = data[data["Остаток"] <= 0]
    elif selected == "Без фото":
        current = data[data["Фото"].fillna("").eq("")]
    else:
        current = data[data["Категория"].fillna("").eq("")]
    metric_card(selected, f"{len(current):,} позиций")
    render_inventory_table(current, "warehouse_attention_table")


def render_movement(operations: pd.DataFrame) -> None:
    period = st.segmented_control("Период", [7, 30, 90], default=30, key="warehouse_movement_period") or 30
    current = movement_window(operations, int(period))
    incoming, outgoing = operation_totals(operations, int(period))
    c1, c2, c3 = st.columns(3)
    with c1: metric_card("Приход", f"{incoming:,} шт.")
    with c2: metric_card("Передано", f"{outgoing:,} шт.")
    with c3: metric_card("Операций", f"{len(current):,}")
    locked_chart(movement_chart(operations, int(period)), f"warehouse_movement_{period}")
    if not current.empty:
        table = current.copy()
        table["Дата"] = table["Дата"].dt.strftime("%d.%m.%Y %H:%M")
        st.dataframe(table, width="stretch", hide_index=True, key="warehouse_operations_table")


def render_supplies(supplies: pd.DataFrame) -> None:
    if supplies.empty:
        st.info("В Baserow пока нет данных для реестра поставок.")
        return
    c1, c2, c3 = st.columns(3)
    with c1: metric_card("Поставок", f"{len(supplies):,}")
    with c2: metric_card("Получено", f"{int(supplies['Получено'].sum()):,} шт.")
    with c3: metric_card("Ожидается", f"{int(supplies['Ожидается'].sum()):,} шт.")
    table = supplies.copy()
    table["Дата"] = table["Дата"].dt.strftime("%d.%m.%Y")
    st.dataframe(table, width="stretch", hide_index=True, key="warehouse_supplies_table")


def render_setup_help() -> None:
    st.error("Раздел склада пока не подключён к Baserow.")
    st.markdown("""
Добавьте в **Streamlit Cloud → App settings → Secrets** read-only подключение:

```toml
[baserow]
url = "https://storage.princess-jewelry.com"
token = "READ_ONLY_TOKEN"
souvenirs_table_id = 642
components_table_id = 643
operations_table_id = 644
supplies_table_id = 645
```
""")


def render_navigation(current_section: str) -> None:
    """Functional desktop navigation: each button switches the lazy warehouse block."""
    with st.sidebar:
        st.markdown("**Princess Warehouse Analytics**")
        st.markdown('<div class="nav-hint">Разделы склада</div>', unsafe_allow_html=True)
        for section, _anchor_name, label in WAREHOUSE_SECTIONS:
            if st.button(
                label,
                key=f"warehouse_nav_{section}",
                width="stretch",
                type="primary" if section == current_section else "secondary",
            ) and section != current_section:
                st.session_state["warehouse_section"] = section
                st.rerun()
        st.success("Данные склада подключены")
        st.caption("Источник: Baserow · только чтение")


def render_warehouse_dashboard() -> None:
    st.markdown(WAREHOUSE_CSS, unsafe_allow_html=True)
    st.markdown('<section class="warehouse-header"><div class="warehouse-header-kicker">BASEROW · WAREHOUSE ANALYTICS</div><h2>Сувениры и касты на складе</h2><p>Остатки, проблемные позиции, движение товара и поставки. Подключение работает только на чтение.</p></section>', unsafe_allow_html=True)
    try:
        config = WarehouseConfig.load()
    except WarehouseConfigError:
        render_setup_help()
        return
    try:
        with st.spinner("Загружаем актуальные данные Baserow..."):
            bundle = load_bundle(config)
    except (WarehouseApiError, ValueError) as exc:
        st.error(str(exc))
        return
    valid_sections = [section for section, _anchor_name, _label in WAREHOUSE_SECTIONS]
    if st.session_state.get("warehouse_section") not in valid_sections:
        st.session_state["warehouse_section"] = "Обзор"
    render_navigation(str(st.session_state["warehouse_section"]))
    section = st.segmented_control(
        "Раздел склада",
        valid_sections,
        default=str(st.session_state["warehouse_section"]),
        key="warehouse_section",
    ) or "Обзор"
    anchors = {section_name: anchor_name for section_name, anchor_name, _label in WAREHOUSE_SECTIONS}
    st.markdown(f'<div id="{anchors[section]}" class="report-anchor"></div>', unsafe_allow_html=True)
    if section == "Сувениры":
        render_inventory_section(bundle.souvenirs, "Сувениры", "warehouse_souvenirs", config)
    elif section == "Касты":
        render_inventory_section(bundle.components, "Касты и комплектующие", "warehouse_components", config)
    elif section == "Требует внимания":
        render_attention(bundle)
    elif section == "Движение":
        render_movement(bundle.operations)
    elif section == "Поставки":
        render_supplies(bundle.supplies)
    else:
        render_overview(bundle)
    st.caption(f"Обновлено: {bundle.loaded_at:%d.%m.%Y %H:%M} · Baserow · только чтение")
