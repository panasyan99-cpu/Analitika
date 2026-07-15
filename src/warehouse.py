from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
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


class WarehouseConfigError(RuntimeError):
    """Raised when the Streamlit/Baserow connection is not configured."""


class WarehouseApiError(RuntimeError):
    """Raised when Baserow returns an unexpected response."""


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

        base_url = str(
            section.get("url")
            or os.getenv("BASEROW_URL")
            or ""
        ).strip().rstrip("/")
        token = str(
            section.get("token")
            or os.getenv("BASEROW_TOKEN")
            or ""
        ).strip()

        def table_id(name: str) -> int:
            env_name = f"BASEROW_{name.upper()}_TABLE_ID"
            raw = section.get(f"{name}_table_id") or os.getenv(env_name)
            if raw in (None, ""):
                return int(DEFAULT_TABLE_IDS[name])
            try:
                return int(raw)
            except (TypeError, ValueError) as exc:
                raise WarehouseConfigError(
                    f"Некорректный ID таблицы Baserow: {name}."
                ) from exc

        if not base_url or not token:
            raise WarehouseConfigError(
                "Не настроено read-only подключение к Baserow. "
                "Добавьте секцию [baserow] в Secrets приложения."
            )

        return cls(
            base_url=base_url,
            token=token,
            souvenirs_table_id=table_id("souvenirs"),
            components_table_id=table_id("components"),
            operations_table_id=table_id("operations"),
            supplies_table_id=table_id("supplies"),
        )


WAREHOUSE_CSS = """
<style>
.warehouse-header {
  border: 1px solid #e9e4dc;
  border-radius: 20px;
  padding: 22px 24px;
  margin: 8px 0 18px;
  background: linear-gradient(135deg, rgba(255,255,255,.98), rgba(247,241,231,.94));
  box-shadow: 0 12px 34px rgba(34,24,9,.06);
}
.warehouse-header-kicker {
  color: #b7893f;
  font-size: 12px;
  font-weight: 800;
  letter-spacing: .12em;
  text-transform: uppercase;
}
.warehouse-header h2 {
  font-family: Georgia, serif;
  margin: 4px 0 6px;
  font-size: 36px;
  color: #171411;
}
.warehouse-header p { color: #6c6c6c; margin: 0; }
.wh-metric {
  border: 1px solid #e9e4dc;
  border-radius: 16px;
  padding: 17px 18px;
  min-height: 112px;
  background: rgba(255,255,255,.97);
  box-shadow: 0 8px 25px rgba(34,24,9,.045);
}
.wh-metric-label {
  color: #7c7469;
  font-size: 12px;
  letter-spacing: .055em;
  text-transform: uppercase;
  font-weight: 750;
}
.wh-metric-value {
  margin-top: 8px;
  font-size: 30px;
  line-height: 1.05;
  font-weight: 800;
  color: #171411;
}
.wh-metric-note { color: #777067; font-size: 12px; margin-top: 7px; }
.wh-alert {
  border-left: 4px solid #b7893f;
  background: #fffaf1;
  padding: 13px 15px;
  border-radius: 10px;
  margin: 8px 0 14px;
}
.wh-stock-card {
  border: 1px solid #e9e4dc;
  border-radius: 16px;
  padding: 13px;
  background: rgba(255,255,255,.97);
  box-shadow: 0 8px 22px rgba(34,24,9,.04);
  min-height: 178px;
  margin-bottom: 12px;
}
.wh-stock-card .sku { font-size: 17px; font-weight: 800; color: #171411; }
.wh-stock-card .meta { color: #6c6c6c; font-size: 13px; line-height: 1.45; }
.wh-stock-card .balance { margin-top: 8px; font-weight: 800; }
.wh-status-ok { color: #2f6f45; }
.wh-status-warning { color: #9b6500; }
.wh-status-critical { color: #a82f2f; }
.wh-photo-placeholder {
  min-height: 180px;
  border-radius: 14px;
  border: 1px dashed #d8c8ad;
  background: #faf7f2;
  display: flex;
  align-items: center;
  justify-content: center;
  color: #8b8174;
  margin-bottom: 10px;
}
@media (max-width: 980px) {
  .warehouse-header h2 { font-size: 30px; }
  .wh-metric-value { font-size: 26px; }
}
@media (max-width: 640px) {
  .warehouse-header { padding: 18px; }
  .warehouse-header h2 { font-size: 27px; }
}
</style>
"""


def _request_headers(token: str) -> dict[str, str]:
    return {
        "Authorization": f"Token {token}",
        "Accept": "application/json",
        "User-Agent": "Princess-Analitika-Warehouse/1.0",
    }


@st.cache_data(ttl=60, max_entries=16, show_spinner=False)
def fetch_table_rows(
    base_url: str,
    token: str,
    table_id: int,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    page = 1
    session = requests.Session()
    session.headers.update(_request_headers(token))

    while True:
        url = f"{base_url}/api/database/rows/table/{int(table_id)}/"
        try:
            response = session.get(
                url,
                params={
                    "user_field_names": "true",
                    "size": 200,
                    "page": page,
                },
                timeout=35,
            )
        except requests.RequestException as exc:
            raise WarehouseApiError(
                "Не удалось подключиться к Baserow."
            ) from exc

        if response.status_code == 401:
            raise WarehouseApiError(
                "Baserow отклонил read-only токен сайта."
            )
        if response.status_code == 403:
            raise WarehouseApiError(
                f"У токена сайта нет права Read для таблицы {table_id}."
            )
        if not response.ok:
            raise WarehouseApiError(
                f"Baserow вернул HTTP {response.status_code} для таблицы {table_id}."
            )

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
    # Self-hosted Baserow installations can expose user-files either publicly
    # or behind the same Token header as the Database API. Support both.
    for headers in (_request_headers(token), {"User-Agent": "Princess-Analitika-Warehouse/1.0"}):
        try:
            response = requests.get(url, headers=headers, timeout=20)
            if response.ok and response.content:
                return response.content
        except requests.RequestException:
            continue
    return None


def clear_warehouse_cache() -> None:
    fetch_table_rows.clear()
    fetch_image_bytes.clear()


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
    scalar = _scalar(value)
    if scalar is None:
        return ""
    return str(scalar).strip()


def relation_values(value: Any) -> list[str]:
    if value is None:
        return []
    if not isinstance(value, list):
        value = [value]
    result = []
    for item in value:
        text = text_value(item)
        if text and text not in result:
            result.append(text)
    return result


def number_value(value: Any, default: float = 0.0) -> float:
    value = _scalar(value)
    if value in (None, ""):
        return default
    if isinstance(value, bool):
        return float(int(value))
    if isinstance(value, (int, float)):
        return float(value)
    clean = str(value).strip().replace("\u00a0", "").replace(" ", "")
    if not clean:
        return default
    if clean.count(",") == 1 and clean.count(".") == 0:
        clean = clean.replace(",", ".")
    clean = re.sub(r"[^0-9.\-]", "", clean)
    try:
        return float(clean)
    except ValueError:
        return default


def datetime_value(value: Any) -> pd.Timestamp | pd.NaT:
    value = _scalar(value)
    if value in (None, ""):
        return pd.NaT
    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
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
    if not files:
        return ""
    item = files[0]
    if not isinstance(item, dict):
        return ""

    url = str(item.get("url") or "").strip()
    if not url:
        thumbnails = item.get("thumbnails") or {}
        for name in ("large", "card_cover", "small", "tiny"):
            candidate = thumbnails.get(name)
            if isinstance(candidate, dict) and candidate.get("url"):
                url = str(candidate["url"])
                break
    if not url:
        return ""
    return urljoin(base_url + "/", url)


def stock_status(balance: float, minimum: float) -> str:
    minimum = minimum if minimum > 0 else DEFAULT_MINIMUM_STOCK
    if balance <= 0:
        return "Нет в наличии"
    if balance <= minimum:
        return "Ниже минимума"
    if balance <= EARLY_WARNING_STOCK:
        return "Заканчивается"
    return "В наличии"


def status_order(status: str) -> int:
    return {
        "Нет в наличии": 0,
        "Ниже минимума": 1,
        "Заканчивается": 2,
        "В наличии": 3,
    }.get(status, 9)


def normalize_inventory_rows(
    rows: list[dict[str, Any]],
    *,
    section: str,
    base_url: str,
) -> pd.DataFrame:
    records: list[dict[str, Any]] = []
    for row in rows:
        sku = text_value(
            first_existing(
                row,
                ("Артикул", "SKU", "Наименование", "Название"),
            )
        )
        if not sku:
            continue

        balance = number_value(
            first_existing(row, ("Остаток", "Текущий остаток"))
        )
        minimum = number_value(
            first_existing(row, ("Минимальный остаток", "Минимум")),
            DEFAULT_MINIMUM_STOCK,
        )
        if minimum <= 0:
            minimum = DEFAULT_MINIMUM_STOCK

        category = text_value(
            first_existing(row, ("Категория", "Тип", "Вид"))
        )
        material = text_value(
            first_existing(row, ("Материал", "Металл"))
        )
        stone = text_value(
            first_existing(row, ("Камень", "Вставка"))
        )
        color = text_value(first_existing(row, ("Цвет",)))
        boxes = text_value(
            first_existing(row, ("Номера коробок", "Коробка", "Коробки"))
        )
        supplies = relation_values(
            first_existing(row, ("Поставки", "Поставка"))
        )
        photo = file_url(
            first_existing(row, ("Фото", "Изображение")),
            base_url,
        )
        description = text_value(
            first_existing(row, ("Описание", "Комментарий", "Notes"))
        )

        record = {
            "ID": int(row.get("id", 0) or 0),
            "Раздел": section,
            "Фото": photo,
            "Артикул": sku,
            "Категория": category,
            "Материал": material,
            "Камень": stone,
            "Цвет": color,
            "Коробки": boxes,
            "Описание": description,
            "Поставки": "; ".join(supplies),
            "Остаток": int(round(balance)),
            "Минимальный остаток": int(round(minimum)),
            "По документу": int(round(number_value(row.get("По документу, шт.")))),
            "Получено": int(round(number_value(row.get("Получено по поставке, шт.")))),
            "Ожидается": int(round(number_value(row.get("Ожидается, шт.")))),
        }
        record["Статус"] = stock_status(balance, minimum)
        record["СтатусПорядок"] = status_order(record["Статус"])
        records.append(record)

    columns = [
        "ID", "Раздел", "Фото", "Артикул", "Категория", "Материал",
        "Камень", "Цвет", "Коробки", "Описание", "Поставки", "Остаток",
        "Минимальный остаток", "По документу", "Получено", "Ожидается",
        "Статус", "СтатусПорядок",
    ]
    frame = pd.DataFrame.from_records(records, columns=columns)
    if not frame.empty:
        frame = frame.sort_values(
            ["СтатусПорядок", "Остаток", "Артикул"],
            ascending=[True, True, True],
        ).reset_index(drop=True)
    return frame


def normalize_operations(rows: list[dict[str, Any]]) -> pd.DataFrame:
    records = []
    for row in rows:
        operation_type = text_value(
            first_existing(row, ("Тип операции", "Операция"))
        )
        quantity = number_value(row.get("Количество"))
        change = number_value(row.get("Изменение"), float("nan"))
        if pd.isna(change):
            lowered = operation_type.casefold()
            if "приход" in lowered:
                change = abs(quantity)
            elif "передач" in lowered or "расход" in lowered:
                change = -abs(quantity)
            else:
                change = quantity

        souvenir_skus = relation_values(row.get("Товар сувенирки"))
        component_skus = relation_values(row.get("Комплектующее"))
        sku_values = souvenir_skus + [x for x in component_skus if x not in souvenir_skus]
        section = text_value(row.get("Раздел"))
        if not section:
            section = "Сувенирка" if souvenir_skus else "Комплектующие" if component_skus else ""

        records.append(
            {
                "Дата": datetime_value(
                    first_existing(
                        row,
                        ("Дата и время", "Дата", "Created on", "Дата создания"),
                    )
                ),
                "Тип операции": operation_type,
                "Раздел": section,
                "SKU": "; ".join(sku_values),
                "Количество": int(round(abs(quantity))),
                "Изменение": int(round(change)),
                "Batch ID": text_value(row.get("Batch ID")),
                "Поставка": text_value(row.get("Поставка")),
                "Ответственный": text_value(
                    first_existing(row, ("Ответственный", "Сотрудник", "Пользователь"))
                ),
                "Комментарий": text_value(row.get("Комментарий")),
            }
        )
    frame = pd.DataFrame.from_records(records)
    if not frame.empty:
        frame = frame.sort_values("Дата", ascending=False, na_position="last").reset_index(drop=True)
    return frame


def normalize_supplies(
    supply_rows: list[dict[str, Any]],
    souvenir_df: pd.DataFrame,
) -> pd.DataFrame:
    supply_records = []
    for row in supply_rows:
        number = text_value(
            first_existing(row, ("№ поставки", "Поставка", "Название"))
        )
        if not number:
            continue
        supply_records.append(
            {
                "Поставка": number,
                "Дата": datetime_value(first_existing(row, ("Дата", "Дата создания"))),
                "Поставщик": text_value(row.get("Поставщик")),
                "Invoice": text_value(row.get("Invoice")),
                "Статус": text_value(row.get("Статус")),
                "Комментарий": text_value(row.get("Комментарий")),
            }
        )
    supplies = pd.DataFrame.from_records(supply_records)

    linked_records: list[dict[str, Any]] = []
    if not souvenir_df.empty:
        for _, row in souvenir_df.iterrows():
            for supply in [x.strip() for x in str(row.get("Поставки", "")).split(";") if x.strip()]:
                linked_records.append(
                    {
                        "Поставка": supply,
                        "SKU": 1,
                        "По документу": int(row.get("По документу", 0)),
                        "Получено": int(row.get("Получено", 0)),
                        "Ожидается": int(row.get("Ожидается", 0)),
                    }
                )
    if linked_records:
        totals = (
            pd.DataFrame(linked_records)
            .groupby("Поставка", as_index=False)
            .agg(
                SKU=("SKU", "sum"),
                **{
                    "По документу": ("По документу", "sum"),
                    "Получено": ("Получено", "sum"),
                    "Ожидается": ("Ожидается", "sum"),
                },
            )
        )
        if supplies.empty:
            supplies = totals
        else:
            supplies = supplies.merge(totals, on="Поставка", how="outer")

    if supplies.empty:
        return pd.DataFrame(
            columns=[
                "Поставка", "Дата", "Поставщик", "Invoice", "Статус",
                "SKU", "По документу", "Получено", "Ожидается", "Комментарий",
            ]
        )
    required_defaults: dict[str, Any] = {
        "Дата": pd.NaT,
        "Поставщик": "",
        "Invoice": "",
        "Статус": "",
        "Комментарий": "",
        "SKU": 0,
        "По документу": 0,
        "Получено": 0,
        "Ожидается": 0,
    }
    for column, default in required_defaults.items():
        if column not in supplies:
            supplies[column] = default
    for column in ("SKU", "По документу", "Получено", "Ожидается"):
        supplies[column] = supplies[column].fillna(0).astype(int)
    supplies["Статус"] = supplies["Статус"].fillna("").astype(str)
    supplies["СтатусПорядок"] = supplies["Статус"].str.casefold().map(
        lambda value: 0 if "частич" in value else 1
    )
    return supplies.sort_values(
        ["СтатусПорядок", "Дата"],
        ascending=[True, False],
        na_position="last",
    ).reset_index(drop=True)


@dataclass
class WarehouseBundle:
    souvenirs: pd.DataFrame
    components: pd.DataFrame
    operations: pd.DataFrame
    supplies: pd.DataFrame
    loaded_at: datetime


def load_bundle(config: WarehouseConfig) -> WarehouseBundle:
    souvenir_rows = fetch_table_rows(
        config.base_url, config.token, config.souvenirs_table_id
    )
    component_rows = fetch_table_rows(
        config.base_url, config.token, config.components_table_id
    )
    operation_rows = fetch_table_rows(
        config.base_url, config.token, config.operations_table_id
    )
    supply_rows = fetch_table_rows(
        config.base_url, config.token, config.supplies_table_id
    )

    souvenirs = normalize_inventory_rows(
        souvenir_rows,
        section="Сувенирка",
        base_url=config.base_url,
    )
    components = normalize_inventory_rows(
        component_rows,
        section="Комплектующие",
        base_url=config.base_url,
    )
    operations = normalize_operations(operation_rows)
    supplies = normalize_supplies(supply_rows, souvenirs)
    return WarehouseBundle(
        souvenirs=souvenirs,
        components=components,
        operations=operations,
        supplies=supplies,
        loaded_at=datetime.now(),
    )


def locked_chart(fig: go.Figure, key: str) -> None:
    # Preserve larger chart-specific margins. Previously this helper always
    # reset the right margin to 20 px, so labels drawn outside horizontal bars
    # could be clipped even when the chart itself requested extra space.
    current_margin = fig.layout.margin

    def margin_value(name: str, minimum: int) -> int:
        value = getattr(current_margin, name, None)
        try:
            return max(int(value or 0), minimum)
        except (TypeError, ValueError):
            return minimum

    fig.update_layout(
        dragmode=False,
        clickmode="event",
        hovermode="closest",
        legend_itemclick=False,
        legend_itemdoubleclick=False,
        margin=dict(
            l=margin_value("l", 20),
            r=margin_value("r", 28),
            t=margin_value("t", 55),
            b=margin_value("b", 36),
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, Arial, sans-serif", color="#171411"),
        uniformtext_minsize=10,
        uniformtext_mode="show",
    )
    fig.update_xaxes(fixedrange=True, automargin=True)
    fig.update_yaxes(fixedrange=True, automargin=True)
    st.plotly_chart(
        fig,
        width="stretch",
        key=key,
        config=LOCKED_CHART_CONFIG,
    )


def metric_card(label: str, value: str, note: str = "") -> None:
    st.markdown(
        '<div class="wh-metric">'
        f'<div class="wh-metric-label">{escape(label)}</div>'
        f'<div class="wh-metric-value">{escape(value)}</div>'
        f'<div class="wh-metric-note">{escape(note)}</div>'
        '</div>',
        unsafe_allow_html=True,
    )


def render_navigation() -> None:
    items = [
        ("#warehouse-overview", "📊 Обзор"),
        ("#warehouse-souvenirs", "🎁 Сувениры"),
        ("#warehouse-components", "🧩 Касты"),
        ("#warehouse-attention", "⚠️ Требует внимания"),
        ("#warehouse-movement", "↔️ Движение"),
        ("#warehouse-supplies", "📦 Поставки"),
        ("#about", "ℹ️ О платформе"),
    ]
    links = "".join(f'<a href="{href}">{label}</a>' for href, label in items)
    with st.sidebar:
        st.markdown(
            '<div class="sidebar-brand"><b>PRINCESS</b><br><span>Warehouse Analytics</span></div>',
            unsafe_allow_html=True,
        )
        st.markdown(f'<nav class="side-nav">{links}</nav>', unsafe_allow_html=True)
        st.markdown("---")
        st.success("Данные склада подключены")
        st.caption("Источник: Baserow · только чтение")
        st.caption("Разработка: Vladimir Panasyan")

    mobile_links = "".join(
        f'<a href="{href}">{label.split(" ", 1)[0]}</a>' for href, label in items[:-1]
    )
    st.markdown(
        f'<div class="mobile-nav-shell"><nav class="mobile-nav">{mobile_links}</nav></div>',
        unsafe_allow_html=True,
    )


def inventory_available(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return frame.copy()
    return frame[frame["Остаток"] > 0].copy()


def low_stock_count(frame: pd.DataFrame) -> int:
    if frame.empty:
        return 0
    return int(((frame["Остаток"] > 0) & (frame["Остаток"] <= EARLY_WARNING_STOCK)).sum())


def below_minimum_count(frame: pd.DataFrame) -> int:
    if frame.empty:
        return 0
    return int(
        (
            (frame["Остаток"] > 0)
            & (frame["Остаток"] <= frame["Минимальный остаток"])
        ).sum()
    )


def movement_window(operations: pd.DataFrame, days: int) -> pd.DataFrame:
    if operations.empty or "Дата" not in operations:
        return operations.copy()
    threshold = pd.Timestamp.now().normalize() - pd.Timedelta(days=days - 1)
    return operations[operations["Дата"].notna() & (operations["Дата"] >= threshold)].copy()


def operation_totals(operations: pd.DataFrame, days: int = 30) -> tuple[int, int]:
    current = movement_window(operations, days)
    if current.empty:
        return 0, 0
    incoming = int(current.loc[current["Изменение"] > 0, "Изменение"].sum())
    outgoing = int(abs(current.loc[current["Изменение"] < 0, "Изменение"].sum()))
    return incoming, outgoing


def _formatted_integer(value: Any) -> str:
    try:
        return f"{int(round(float(value))):,}".replace(",", " ")
    except (TypeError, ValueError):
        return "0"


def _positive_axis_range(values: Iterable[Any], padding: float = 0.18) -> list[float]:
    numeric = pd.to_numeric(pd.Series(list(values)), errors="coerce").dropna()
    if numeric.empty:
        return [0.0, 1.0]
    maximum = max(float(numeric.max()), 0.0)
    if maximum <= 0:
        return [0.0, 1.0]
    # A proportional reserve plus a small absolute reserve guarantees that an
    # outside label is visible both for large totals and for small SKU counts.
    upper = maximum * (1.0 + padding)
    upper = max(upper, maximum + max(1.0, maximum * 0.08))
    return [0.0, upper]


def category_chart(frame: pd.DataFrame, title: str) -> go.Figure:
    data = inventory_available(frame)
    if data.empty:
        return go.Figure().update_layout(title=title)
    data["Категория"] = data["Категория"].replace("", "Без категории")
    grouped = (
        data.groupby("Категория", as_index=False)["Остаток"]
        .sum()
        .sort_values("Остаток", ascending=True)
        .tail(12)
    )
    values = grouped["Остаток"].astype(float)
    labels = [_formatted_integer(value) for value in values]
    fig = go.Figure(
        go.Bar(
            x=values,
            y=grouped["Категория"],
            orientation="h",
            text=labels,
            textposition="outside",
            textfont=dict(size=11),
            cliponaxis=False,
            marker=dict(color="#b7893f", line=dict(width=0)),
            hovertemplate="%{y}<br>%{x:,.0f} шт.<extra></extra>",
        )
    )
    fig.update_layout(
        title=title,
        height=max(350, 38 * len(grouped) + 120),
        margin=dict(l=20, r=82, t=55, b=38),
        bargap=0.28,
        showlegend=False,
    )
    fig.update_xaxes(
        range=_positive_axis_range(values, padding=0.20),
        rangemode="tozero",
        tickformat=",.0f",
        showgrid=False,
        zeroline=False,
    )
    fig.update_yaxes(showgrid=False)
    return fig


def stone_sku_chart(frame: pd.DataFrame) -> go.Figure:
    data = inventory_available(frame)
    records: list[str] = []
    for value in data.get("Камень", pd.Series(dtype=str)).fillna(""):
        records.extend(
            [part.strip() for part in str(value).split(";") if part.strip()]
        )
    if not records:
        return go.Figure().update_layout(title="Сувениры по камням · SKU")
    grouped = (
        pd.Series(records, name="Камень")
        .value_counts()
        .head(12)
        .sort_values()
    )
    values = grouped.values.astype(float)
    labels = [_formatted_integer(value) for value in values]
    fig = go.Figure(
        go.Bar(
            x=values,
            y=grouped.index,
            orientation="h",
            text=labels,
            textposition="outside",
            textfont=dict(size=11),
            cliponaxis=False,
            marker=dict(color="#6c5a3d", line=dict(width=0)),
            hovertemplate="%{y}<br>%{x:.0f} SKU<extra></extra>",
        )
    )
    fig.update_layout(
        title="Сувениры по камням · количество SKU",
        height=max(350, 38 * len(grouped) + 120),
        margin=dict(l=20, r=74, t=55, b=38),
        bargap=0.28,
        showlegend=False,
    )
    fig.update_xaxes(
        range=_positive_axis_range(values, padding=0.22),
        rangemode="tozero",
        dtick=1 if max(values, default=0) <= 12 else None,
        showgrid=False,
        zeroline=False,
    )
    fig.update_yaxes(showgrid=False)
    return fig


def movement_chart(operations: pd.DataFrame, days: int = 30) -> go.Figure:
    data = movement_window(operations, days)
    if data.empty:
        return go.Figure().update_layout(title=f"Движение за {days} дней")

    data = data.dropna(subset=["Дата"]).copy()
    if data.empty:
        return go.Figure().update_layout(title=f"Движение за {days} дней")

    data["День"] = data["Дата"].dt.normalize()
    data["Приход"] = data["Изменение"].clip(lower=0)
    data["Передано"] = -data["Изменение"].clip(upper=0)
    grouped = (
        data.groupby("День", as_index=False)[["Приход", "Передано"]]
        .sum()
        .sort_values("День")
    )

    today = pd.Timestamp.now().normalize()
    period_start = today - pd.Timedelta(days=max(days - 1, 0))
    period_end = today + pd.Timedelta(days=1)

    # Plotly interprets widths on a date axis in milliseconds. Explicitly using
    # a narrow width prevents a single active day from becoming one enormous
    # rectangle that fills almost the entire chart.
    bar_width_ms = 9 * 60 * 60 * 1000

    incoming_labels = [
        _formatted_integer(value) if float(value) > 0 else ""
        for value in grouped["Приход"]
    ]
    outgoing_labels = [
        _formatted_integer(value) if float(value) > 0 else ""
        for value in grouped["Передано"]
    ]

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=grouped["День"],
            y=grouped["Приход"],
            name="Приход",
            width=bar_width_ms,
            text=incoming_labels,
            textposition="outside",
            cliponaxis=False,
            marker=dict(color="#5f876d", line=dict(width=0)),
            hovertemplate="%{x|%d.%m.%Y}<br>%{y:,.0f} шт.<extra></extra>",
        )
    )
    fig.add_trace(
        go.Bar(
            x=grouped["День"],
            y=grouped["Передано"],
            name="Передано в бухгалтерию",
            width=bar_width_ms,
            text=outgoing_labels,
            textposition="outside",
            cliponaxis=False,
            marker=dict(color="#b7893f", line=dict(width=0)),
            hovertemplate="%{x|%d.%m.%Y}<br>%{y:,.0f} шт.<extra></extra>",
        )
    )
    fig.update_layout(
        title=f"Приход и передача за {days} дней",
        barmode="group",
        bargap=0.55,
        bargroupgap=0.12,
        height=330,
        margin=dict(l=20, r=54, t=55, b=48),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
        ),
    )
    fig.update_xaxes(
        type="date",
        range=[period_start, period_end],
        tickformat="%d.%m",
        nticks=8,
        showgrid=False,
        zeroline=False,
        title=None,
    )
    fig.update_yaxes(
        rangemode="tozero",
        tickformat=",.0f",
        gridcolor="rgba(90, 78, 62, 0.12)",
        zeroline=False,
        title=None,
    )
    return fig


def apply_inventory_filters(frame: pd.DataFrame, prefix: str) -> pd.DataFrame:
    if frame.empty:
        return frame.copy()
    result = frame.copy()
    first, second, third = st.columns(3)
    with first:
        categories = sorted(x for x in result["Категория"].dropna().unique().tolist() if x)
        selected_categories = st.multiselect(
            "Категория",
            categories,
            key=f"{prefix}_category",
        )
    with second:
        materials = sorted(
            dict.fromkeys(
                part.strip()
                for value in result["Материал"].fillna("")
                for part in str(value).split(";")
                if part.strip()
            )
        )
        selected_materials = st.multiselect(
            "Материал",
            materials,
            key=f"{prefix}_material",
        )
    with third:
        statuses = ["В наличии", "Заканчивается", "Ниже минимума", "Нет в наличии"]
        selected_statuses = st.multiselect(
            "Статус остатка",
            statuses,
            default=["В наличии", "Заканчивается", "Ниже минимума"],
            key=f"{prefix}_status",
        )

    stone_col, search_col = st.columns([1, 2])
    with stone_col:
        stones = sorted(
            dict.fromkeys(
                part.strip()
                for value in result["Камень"].fillna("")
                for part in str(value).split(";")
                if part.strip()
            )
        )
        selected_stones = st.multiselect(
            "Камень",
            stones,
            key=f"{prefix}_stone",
        )
    with search_col:
        search = st.text_input(
            "Артикул, коробка или текст",
            key=f"{prefix}_search",
            placeholder="Например: KH12, Agate или коробка 18",
        ).strip().casefold()

    if selected_categories:
        result = result[result["Категория"].isin(selected_categories)]
    if selected_materials:
        result = result[
            result["Материал"].fillna("").map(
                lambda value: any(
                    item in [part.strip() for part in str(value).split(";")]
                    for item in selected_materials
                )
            )
        ]
    if selected_stones:
        result = result[
            result["Камень"].fillna("").map(
                lambda value: any(
                    item in [part.strip() for part in str(value).split(";")]
                    for item in selected_stones
                )
            )
        ]
    if selected_statuses:
        result = result[result["Статус"].isin(selected_statuses)]
    else:
        result = result.iloc[0:0]
    if search:
        haystack = (
            result[["Артикул", "Категория", "Материал", "Камень", "Цвет", "Коробки", "Описание"]]
            .fillna("")
            .astype(str)
            .agg(" ".join, axis=1)
            .str.casefold()
        )
        result = result[haystack.str.contains(re.escape(search), regex=True)]
    return result.reset_index(drop=True)


def render_inventory_table(frame: pd.DataFrame, key: str) -> None:
    visible_columns = [
        "Фото", "Артикул", "Категория", "Материал", "Камень", "Цвет",
        "Коробки", "Остаток", "Минимальный остаток", "Статус", "Поставки",
    ]
    table = frame[[column for column in visible_columns if column in frame.columns]].copy()
    st.dataframe(
        table,
        width="stretch",
        hide_index=True,
        key=key,
        column_config={
            "Фото": st.column_config.ImageColumn("Фото", width="medium"),
            "Остаток": st.column_config.NumberColumn("Остаток", format="%d шт."),
            "Минимальный остаток": st.column_config.NumberColumn("Минимум", format="%d шт."),
        },
    )


def status_css(status: str) -> str:
    if status == "В наличии":
        return "wh-status-ok"
    if status == "Заканчивается":
        return "wh-status-warning"
    return "wh-status-critical"


def render_inventory_cards(
    frame: pd.DataFrame,
    *,
    config: WarehouseConfig,
    key: str,
) -> None:
    if frame.empty:
        st.info("По выбранным фильтрам позиций нет.")
        return

    page_size = st.segmented_control(
        "Карточек на странице",
        [6, 12, 18],
        default=12,
        key=f"{key}_page_size",
    ) or 12
    page_count = max(1, (len(frame) + int(page_size) - 1) // int(page_size))
    page = st.number_input(
        "Страница",
        min_value=1,
        max_value=page_count,
        value=1,
        step=1,
        key=f"{key}_page",
    )
    start = (int(page) - 1) * int(page_size)
    current = frame.iloc[start : start + int(page_size)]

    for row_start in range(0, len(current), 3):
        columns = st.columns(3)
        for column, (_, item) in zip(columns, current.iloc[row_start : row_start + 3].iterrows()):
            with column:
                image = fetch_image_bytes(str(item.get("Фото", "")), config.token)
                if image:
                    st.image(image, width="stretch")
                else:
                    st.markdown(
                        '<div class="wh-photo-placeholder">Нет фотографии</div>',
                        unsafe_allow_html=True,
                    )
                meta_parts = [
                    item.get("Категория", ""),
                    item.get("Материал", ""),
                    item.get("Камень", ""),
                    item.get("Цвет", ""),
                ]
                meta = " · ".join(str(x) for x in meta_parts if str(x).strip()) or "Характеристики не указаны"
                boxes = str(item.get("Коробки", "")).strip()
                boxes_line = f"<br>Коробки: {escape(boxes)}" if boxes else ""
                status = str(item.get("Статус", ""))
                card_html = (
                    '<div class="wh-stock-card">'
                    f'<div class="sku">{escape(str(item.get("Артикул", "")))}</div>'
                    f'<div class="meta">{escape(meta)}{boxes_line}</div>'
                    f'<div class="balance {status_css(status)}">'
                    f'{escape(status)} · {int(item.get("Остаток", 0)):,} шт.'
                    '</div></div>'
                )
                st.markdown(card_html, unsafe_allow_html=True)


def render_inventory_section(
    frame: pd.DataFrame,
    *,
    title: str,
    prefix: str,
    config: WarehouseConfig,
) -> None:
    st.markdown(
        f'<div class="wh-alert"><b>{escape(title)}</b><br>'
        'Показываются позиции, которые уже приняты в Baserow и ещё находятся на складе.</div>',
        unsafe_allow_html=True,
    )
    filtered = apply_inventory_filters(frame, prefix)
    c1, c2, c3 = st.columns(3)
    with c1:
        metric_card("Найдено SKU", f"{len(filtered):,}")
    with c2:
        metric_card("Единиц", f"{int(filtered['Остаток'].clip(lower=0).sum()) if not filtered.empty else 0:,}")
    with c3:
        metric_card("Остаток ≤ 15", f"{low_stock_count(filtered):,}", "Раннее предупреждение")

    view = st.segmented_control(
        "Вид",
        ["Карточки", "Таблица"],
        default="Карточки",
        key=f"{prefix}_view",
    ) or "Карточки"
    if view == "Таблица":
        render_inventory_table(filtered, key=f"{prefix}_table")
    else:
        render_inventory_cards(filtered, config=config, key=prefix)


def render_overview(bundle: WarehouseBundle) -> None:
    souvenirs = inventory_available(bundle.souvenirs)
    components = inventory_available(bundle.components)
    incoming, outgoing = operation_totals(bundle.operations, 30)
    attention_total = low_stock_count(bundle.souvenirs) + low_stock_count(bundle.components)
    partial_supplies = 0
    if not bundle.supplies.empty:
        partial_supplies = int(
            bundle.supplies["Статус"].fillna("").str.casefold().str.contains("частич").sum()
        )

    first = st.columns(4)
    with first[0]:
        metric_card("Сувениры", f"{len(souvenirs):,} SKU", f"{int(souvenirs['Остаток'].sum()) if not souvenirs.empty else 0:,} изделий")
    with first[1]:
        metric_card("Касты", f"{len(components):,} SKU", f"{int(components['Остаток'].sum()) if not components.empty else 0:,} единиц")
    with first[2]:
        metric_card("Остаток ≤ 15", f"{attention_total:,}", "Требует раннего внимания")
    with first[3]:
        metric_card("Частичные поставки", f"{partial_supplies:,}", "Ожидаются позиции")

    second = st.columns(4)
    with second[0]:
        metric_card("Ниже минимума", f"{below_minimum_count(bundle.souvenirs) + below_minimum_count(bundle.components):,}", "Минимум по умолчанию 10")
    with second[1]:
        metric_card("Нет в наличии", f"{int((bundle.souvenirs['Остаток'] <= 0).sum()) + int((bundle.components['Остаток'] <= 0).sum()):,}")
    with second[2]:
        metric_card("Принято за 30 дней", f"{incoming:,} шт.")
    with second[3]:
        metric_card("Передано за 30 дней", f"{outgoing:,} шт.", "Передача в бухгалтерию")

    charts = st.columns(2)
    with charts[0]:
        locked_chart(category_chart(bundle.souvenirs, "Остаток сувениров по категориям"), "warehouse_category_chart")
    with charts[1]:
        locked_chart(stone_sku_chart(bundle.souvenirs), "warehouse_stone_chart")
    locked_chart(movement_chart(bundle.operations, 30), "warehouse_movement_overview")
    st.caption(
        "В разрезе камней считается количество SKU, а не сумма изделий: одна карточка может содержать несколько камней."
    )


def render_attention(bundle: WarehouseBundle) -> None:
    combined = pd.concat([bundle.souvenirs, bundle.components], ignore_index=True)
    if combined.empty:
        st.info("Складские карточки пока отсутствуют.")
        return
    categories = [
        "Заканчивается",
        "Ниже минимума",
        "Нет в наличии",
        "Без фото",
        "Без категории",
    ]
    selected = st.segmented_control(
        "Проблема",
        categories,
        default="Заканчивается",
        key="warehouse_attention_type",
    ) or "Заканчивается"

    if selected == "Заканчивается":
        data = combined[
            (combined["Остаток"] > combined["Минимальный остаток"])
            & (combined["Остаток"] <= EARLY_WARNING_STOCK)
        ]
    elif selected == "Ниже минимума":
        data = combined[
            (combined["Остаток"] > 0)
            & (combined["Остаток"] <= combined["Минимальный остаток"])
        ]
    elif selected == "Нет в наличии":
        data = combined[combined["Остаток"] <= 0]
    elif selected == "Без фото":
        data = combined[combined["Фото"].fillna("").eq("")]
    else:
        data = combined[combined["Категория"].fillna("").eq("")]

    metric_card(selected, f"{len(data):,} позиций")
    render_inventory_table(data, key="warehouse_attention_table")


def render_movement(operations: pd.DataFrame) -> None:
    period = st.segmented_control(
        "Период",
        [7, 30, 90],
        default=30,
        key="warehouse_movement_period",
    ) or 30
    current = movement_window(operations, int(period))
    incoming, outgoing = operation_totals(operations, int(period))
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Приход", f"{incoming:,} шт.")
    with c2:
        metric_card("Передано", f"{outgoing:,} шт.")
    with c3:
        metric_card("Операций", f"{len(current):,}")
    with c4:
        batches = current["Batch ID"].replace("", pd.NA).dropna().nunique() if not current.empty else 0
        metric_card("Партий", f"{int(batches):,}")
    locked_chart(movement_chart(operations, int(period)), f"warehouse_movement_{period}")
    if current.empty:
        st.info("За выбранный период операций нет.")
        return
    table = current.copy()
    table["Дата"] = table["Дата"].dt.strftime("%d.%m.%Y %H:%M")
    st.dataframe(
        table[[
            "Дата", "Тип операции", "Раздел", "SKU", "Количество",
            "Изменение", "Batch ID", "Ответственный", "Комментарий",
        ]],
        width="stretch",
        hide_index=True,
        key="warehouse_operations_table",
    )


def render_supplies(supplies: pd.DataFrame) -> None:
    if supplies.empty:
        st.info("В Baserow пока нет данных для реестра поставок.")
        return
    partial = supplies[
        supplies["Статус"].fillna("").str.casefold().str.contains("частич")
    ]
    c1, c2, c3 = st.columns(3)
    with c1:
        metric_card("Поставок", f"{len(supplies):,}")
    with c2:
        metric_card("Частично приняты", f"{len(partial):,}")
    with c3:
        metric_card("Ожидается", f"{int(supplies['Ожидается'].sum()):,} шт.")
    table = supplies.drop(columns=["СтатусПорядок"], errors="ignore").copy()
    if "Дата" in table:
        table["Дата"] = table["Дата"].dt.strftime("%d.%m.%Y")
    st.dataframe(table, width="stretch", hide_index=True, key="warehouse_supplies_table")


def render_setup_help() -> None:
    st.error("Раздел склада пока не подключён к Baserow.")
    st.markdown(
        """
Создайте отдельный токен **Streamlit Warehouse Read Only** с правом `Read`
для четырёх таблиц, затем добавьте в **Streamlit Cloud → App settings → Secrets**:

```toml
[baserow]
url = "https://storage.princess-jewelry.com"
token = "READ_ONLY_TOKEN"
souvenirs_table_id = 642
components_table_id = 643
operations_table_id = 644
supplies_table_id = 645
```

Токен Supply Manager и пароль пользователя в сайт не передаются.
        """
    )


def render_warehouse_dashboard() -> None:
    st.markdown(WAREHOUSE_CSS, unsafe_allow_html=True)
    st.markdown(
        '<div class="warehouse-header">'
        '<div class="warehouse-header-kicker">Princess Jewelry · Warehouse</div>'
        '<h2>Сувениры и касты на складе</h2>'
        '<p>Изделия и компоненты, которые уже приняты в Baserow, но ещё не переданы в бухгалтерию и не вышли в зал.</p>'
        '</div>',
        unsafe_allow_html=True,
    )

    try:
        config = WarehouseConfig.load()
    except WarehouseConfigError:
        render_setup_help()
        return

    refresh_col, timestamp_col = st.columns([1, 3])
    with refresh_col:
        if st.button("Обновить данные", type="primary", width="stretch", key="warehouse_refresh"):
            clear_warehouse_cache()
            st.rerun()

    try:
        with st.spinner("Получаем актуальные остатки из Baserow…"):
            bundle = load_bundle(config)
    except WarehouseApiError as exc:
        st.error(str(exc))
        st.caption("Секретный токен не отображается и не передаётся в браузер пользователя.")
        return

    with timestamp_col:
        st.caption(f"Последнее обновление: {bundle.loaded_at:%d.%m.%Y %H:%M:%S} · кэш до 60 секунд")

    render_navigation()

    st.markdown('<div id="warehouse-overview"></div>', unsafe_allow_html=True)
    st.subheader("Обзор")
    render_overview(bundle)

    st.markdown('<div id="warehouse-souvenirs"></div>', unsafe_allow_html=True)
    st.subheader("Сувениры")
    render_inventory_section(
        bundle.souvenirs,
        title="Сувениры, находящиеся на складе",
        prefix="warehouse_souvenirs",
        config=config,
    )

    st.markdown('<div id="warehouse-components"></div>', unsafe_allow_html=True)
    st.subheader("Касты и комплектующие")
    render_inventory_section(
        bundle.components,
        title="Касты и комплектующие, находящиеся на складе",
        prefix="warehouse_components",
        config=config,
    )

    st.markdown('<div id="warehouse-attention"></div>', unsafe_allow_html=True)
    st.subheader("Требует внимания")
    render_attention(bundle)

    st.markdown('<div id="warehouse-movement"></div>', unsafe_allow_html=True)
    st.subheader("Движение склада")
    render_movement(bundle.operations)

    st.markdown('<div id="warehouse-supplies"></div>', unsafe_allow_html=True)
    st.subheader("Поставки")
    render_supplies(bundle.supplies)
