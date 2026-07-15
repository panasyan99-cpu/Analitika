from __future__ import annotations

import gc
import io
import re
import threading
from datetime import datetime
from dataclasses import dataclass
from typing import Any

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from src.currency import get_vnd_per_usd
from openpyxl import load_workbook
from PIL import Image

CATEGORY_LABELS = {
    "Bracelet": "Браслеты",
    "Earrings": "Серьги",
    "Ring": "Кольца",
    "Necklace": "Ожерелья",
    "Pendant": "Подвески",
    "Brooch": "Броши",
    "Other": "Другое",
}

# The first Sonu workbook was reviewed visually. These SKUs are the stable
# photo-based seed. New/unseen models are handled by the user's 50/50 rule.
SLIDER_BRACELETS = {
    "SSNB-AFG018-BS", "SSNB-000042-BS", "SSNB-000022-BS", "SSNB-000028-EM",
    "SSNB-000033-RUBY", "SSNB-000034-BS", "SSNB-000044-SWBT", "SSNB-000046-BS",
    "SSNB-000047-BS", "SSNB-000049-BS", "SSNB-AFG016-MOIS", "XB-SN-000019-BS",
    "XB-SN-000021-TANZ-BS", "XB-SN-KSB015-RDT", "XB-SN-KSB016-RUBY",
    "XB-SN-KSB079-BS", "SSNB-000035-GARN", "SSNB-000038-TANZ", "SSNB-000041-BS",
    "SSNB-000045-BS", "SSNB-AFG019-CIT", "XB-SN-000018-BS", "XB-SN-000752-RUBY",
    "XB-SN-KSB092-BS", "XB-SN-S4381L-BS", "SSNB-000035-RUBY", "SSNB-AFG011-EM",
    "XB-SN-000020-BS", "BB000312-RUBY", "BR001229-AMBER", "SSNB-000005-BS",
    "SSNB-000019-RUBY", "SSNB-000027-APAT", "SSNB-000029-TANZ", "SSNB-000037-TANZ",
    "SSNB-000048-EM", "SSNB-AFG015-BS", "SSNB-AFG020-APAT", "SSNB-AFG022-BS",
    "XB-SN-000024-FPW-RUBY", "XB-SN-000026-EM", "XB-SN-000175-BS",
    "XB-SN-000176-EM", "XB-SN-KSB088-BS", "XB-SN-KSB096-BS", "XB-SN-KST176-IOL",
    "XB-SN-KST177-BS",
}

FULL_CIRCLE_BRACELETS = {
    "XB-SN-KSB075-MOIS", "XB-SN-KSB077-57-MOIS", "XB-SN-KSB090-EM",
    "SSNB-AFG038-BS", "XB-SN-KSB041-BS", "XB-SN-KSB067-BS", "XB-SN-KSB069-CIT",
    "XB-SN-KSB077-37-MOIS", "XB-SN-KSB077-53-MOIS", "XB-SN-KSB078-OPAL",
    "XB-SN-KSB086-RUBY", "XB-SN-KSB087-BS", "XB-SN-KSB089-EM", "XB-SN-KSB103-AMST",
    "XB-SN-KSB104-RUBY", "XB-SN-KSB105-BS", "XB-SN-KSB108-BS", "SSNB-AFG001-BS",
    "SSNB-AFG003-BS", "SSNB-AFG007-BS", "XB-SN-KSB099-BS", "XB-SN-KSB101-EM",
    "XB-SN-KSB054-BS", "BB108-LBT", "XB-SN-KSB050-BS", "XB-SN-KSB064-RUBY",
    "XB-SN-KSB109-RUBY", "GSBM-25-1-MOIS", "BB-000003-EM", "BB100117-MYSTIC",
    "SSNB-AFG025-EM", "SSNB-AFG026-BS", "SSNB-AFG031-BS", "SSNB-AFG039-BS",
    "SSNB-AFG040-BS", "SSNB-AFG048-BS", "SSNB-AFG049-BS", "SSNB-AFG050-BS",
    "SSNB-AFG051-BS", "XB-SN-KSB040-RUBY", "XB-SN-KSB077-34-MOIS",
    "XB-SN-KSB098-BS", "XB-SN-KSB106-BS", "XB-SN-S121-SWBT",
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


@dataclass(frozen=True)
class SonuReport:
    data: pd.DataFrame
    period: str
    supplier: str
    bracelet_images: dict[str, bytes]


def _number(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        text = str(value).replace(" ", "").replace(",", ".")
        try:
            return float(text)
        except ValueError:
            return 0.0


def _clean_store(value: str) -> str:
    text = str(value or "").strip()
    compact = re.sub(r"\s+", "", text).upper()
    if compact.startswith("63"):
        return "63"
    for suffix in ("-RETAIL", "-TIMINGS", "-TIMING"):
        if compact.endswith(suffix):
            return text[: -len(suffix)].rstrip(" -")
    return text


def _extract_period(value: str) -> str:
    text = str(value or "")
    match = re.search(r"(\d{2}\.\d{2}\.\d{4})\s*[-–—]\s*(\d{2}\.\d{2}\.\d{4})", text)
    return f"{match.group(1)} – {match.group(2)}" if match else "Период не определён"


def _thumbnail(image_bytes: bytes, size: tuple[int, int] = (420, 320)) -> bytes:
    try:
        with Image.open(io.BytesIO(image_bytes)) as image:
            image = image.convert("RGB")
            image.thumbnail(size, Image.Resampling.LANCZOS)
            output = io.BytesIO()
            image.save(output, format="JPEG", quality=84, optimize=True)
            return output.getvalue()
    except Exception:
        return image_bytes


def parse_sonu_workbook(file_bytes: bytes) -> SonuReport:
    """Parse the expanded Sonu 1C hierarchy. Return rows are intentionally ignored."""
    workbook = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=False)
    try:
        sheet = workbook.active
        period = _extract_period(sheet.cell(1, 1).value)
        supplier_line = str(sheet.cell(2, 1).value or "").strip()
        supplier = supplier_line.split(":", 1)[-1].strip() if ":" in supplier_line else supplier_line
        if "SONU" not in supplier.upper():
            raise ValueError("Для раздела «Заказ Sonu» нужен отчет, отфильтрованный по поставщику Sonu.")

        hierarchy = str(sheet.cell(4, 1).value or "").upper()
        required = ("МАГАЗИН", "ТОВАР", "КАМЕНЬ", "ПРОБА", "НОМЕНКЛАТУРНАЯ ГРУППА")
        if not all(token in hierarchy for token in required):
            raise ValueError(
                "Неверная структура отчета. Нужны уровни: Магазин, Товар, Камень/вставка, "
                "Проба, Номенклатурная группа и Поставщик."
            )

        rows: list[dict[str, Any]] = []
        row_to_index: dict[int, int] = {}
        context: dict[str, str] = {}
        for row_no in range(7, sheet.max_row + 1):
            value = sheet.cell(row_no, 1).value
            if value is None or str(value).strip() == "":
                continue
            level = int(sheet.row_dimensions[row_no].outlineLevel or 0)
            text = str(value).strip()
            if level == 0:
                context = {"store": _clean_store(text)}
            elif level == 1:
                context["division"] = text
            elif level == 2:
                context["metal_group"] = text
            elif level == 3:
                context["category"] = text
            elif level == 4:
                if not context.get("store") or not context.get("category"):
                    continue
                shipped = _number(sheet.cell(row_no, 4).value)
                sold = _number(sheet.cell(row_no, 8).value)
                sales = _number(sheet.cell(row_no, 9).value)
                # The workbook has no independent current-stock field. Returns are ignored,
                # therefore the only honest balance available is shipped minus sold.
                remaining = max(shipped - sold, 0.0)
                item = {
                    "Магазин": context.get("store", ""),
                    "Раздел": context.get("division", ""),
                    "Металл": context.get("metal_group", ""),
                    "Категория": context.get("category", "Other"),
                    "Категория RU": CATEGORY_LABELS.get(context.get("category", "Other"), context.get("category", "Other")),
                    "SKU": text,
                    "Камень": "Не указан",
                    "Проба": "Не указана",
                    "Отгружено": shipped,
                    "Скорость продаж": sold,
                    "Продажи VND": sales,
                    "Средняя цена VND": sales / sold if sold else _number(sheet.cell(row_no, 6).value),
                    "Расчетный остаток": remaining,
                    "Фото": None,
                }
                rows.append(item)
                row_to_index[row_no] = len(rows) - 1
            elif level == 5 and rows:
                rows[-1]["Камень"] = text
            elif level == 6 and rows:
                rows[-1]["Проба"] = text

        if not rows:
            raise ValueError("В отчете не найдено товарных строк Sonu.")

        bracelet_rows = {
            row_no for row_no, idx in row_to_index.items() if rows[idx]["Категория"] == "Bracelet"
        }
        image_map: dict[str, bytes] = {}
        for image in getattr(sheet, "_images", []):
            try:
                anchor_row = int(image.anchor._from.row) + 1
                if anchor_row not in bracelet_rows:
                    continue
                idx = row_to_index.get(anchor_row)
                if idx is None:
                    continue
                sku = str(rows[idx]["SKU"]).strip().upper()
                data = _thumbnail(image._data())
                rows[idx]["Фото"] = data
                image_map.setdefault(sku, data)
            except Exception:
                continue

        frame = pd.DataFrame(rows)
        numeric_cols = [
            "Отгружено", "Скорость продаж", "Продажи VND", "Средняя цена VND", "Расчетный остаток"
        ]
        for column in numeric_cols:
            frame[column] = pd.to_numeric(frame[column], errors="coerce").fillna(0)
        return SonuReport(frame, period, supplier or "Sonu", image_map)
    finally:
        workbook.close()
        gc.collect()


@st.cache_resource(show_spinner=False)
def _sonu_excel_lock() -> threading.Lock:
    return threading.Lock()


@st.cache_data(show_spinner=False, ttl=1800, max_entries=2)
def cached_parse_sonu(file_bytes: bytes) -> SonuReport:
    with _sonu_excel_lock():
        return parse_sonu_workbook(file_bytes)


def _usd(value: float, rate: float) -> float:
    return float(value) / float(rate) if rate else 0.0


def add_usd_columns(frame: pd.DataFrame, rate: float) -> pd.DataFrame:
    result = frame.copy()
    result["Продажи USD"] = result["Продажи VND"].map(lambda value: _usd(value, rate))
    result["Средняя цена USD"] = result["Средняя цена VND"].map(lambda value: _usd(value, rate))
    return result


def aggregate_sonu(frame: pd.DataFrame, group_columns: list[str], rate: float) -> pd.DataFrame:
    grouped = frame.groupby(group_columns, as_index=False, dropna=False).agg(
        Отгружено=("Отгружено", "sum"),
        **{
            "Скорость продаж": ("Скорость продаж", "sum"),
            "Продажи VND": ("Продажи VND", "sum"),
            "Расчетный остаток": ("Расчетный остаток", "sum"),
            "Моделей": ("SKU", "nunique"),
        },
    )
    grouped["Продажи USD"] = grouped["Продажи VND"] / rate
    grouped["Средняя цена USD"] = grouped["Продажи USD"] / grouped["Скорость продаж"].replace(0, pd.NA)
    grouped["Средняя цена USD"] = grouped["Средняя цена USD"].fillna(0)
    total_sold = float(grouped["Скорость продаж"].sum())
    grouped["Доля продаж"] = grouped["Скорость продаж"] / total_sold if total_sold else 0.0
    return grouped.drop(columns=["Продажи VND"])


def classify_bracelets(frame: pd.DataFrame, rate: float) -> pd.DataFrame:
    bracelets = frame.loc[frame["Категория"] == "Bracelet"].copy()
    if bracelets.empty:
        return pd.DataFrame()
    grouped = bracelets.groupby("SKU", as_index=False).agg(
        Магазинов=("Магазин", "nunique"),
        Камень=("Камень", "first"),
        Проба=("Проба", "first"),
        Отгружено=("Отгружено", "sum"),
        **{
            "Скорость продаж": ("Скорость продаж", "sum"),
            "Продажи VND": ("Продажи VND", "sum"),
            "Расчетный остаток": ("Расчетный остаток", "sum"),
            "Фото": ("Фото", lambda values: next((value for value in values if isinstance(value, (bytes, bytearray))), None)),
        },
    )
    grouped["_sku"] = grouped["SKU"].astype(str).str.upper().str.strip()
    grouped["Тип браслета"] = pd.NA
    grouped["Источник классификации"] = pd.NA
    slider_mask = grouped["_sku"].isin(SLIDER_BRACELETS)
    full_mask = grouped["_sku"].isin(FULL_CIRCLE_BRACELETS)
    grouped.loc[slider_mask, "Тип браслета"] = "С затяжкой"
    grouped.loc[slider_mask, "Источник классификации"] = "Фото"
    grouped.loc[full_mask, "Тип браслета"] = "В круг с камнями"
    grouped.loc[full_mask, "Источник классификации"] = "Фото"

    ambiguous = grouped[grouped["Тип браслета"].isna()].sort_values(
        ["Скорость продаж", "Продажи VND", "SKU"], ascending=[False, False, True]
    )
    # User rule: disputed models are divided 50/50. The extra model (odd count)
    # and the better-selling half go to the slider group.
    slider_count = (len(ambiguous) + 1) // 2
    slider_indices = ambiguous.head(slider_count).index
    full_indices = ambiguous.iloc[slider_count:].index
    grouped.loc[slider_indices, "Тип браслета"] = "С затяжкой"
    grouped.loc[full_indices, "Тип браслета"] = "В круг с камнями"
    grouped.loc[ambiguous.index, "Источник классификации"] = "Правило 50/50"

    grouped["Продажи USD"] = grouped["Продажи VND"] / rate
    grouped["Средняя цена USD"] = grouped["Продажи USD"] / grouped["Скорость продаж"].replace(0, pd.NA)
    grouped["Средняя цена USD"] = grouped["Средняя цена USD"].fillna(0)
    return grouped.drop(columns=["_sku", "Продажи VND"])


def bracelet_type_summary(bracelets: pd.DataFrame) -> pd.DataFrame:
    if bracelets.empty:
        return pd.DataFrame()
    result = bracelets.groupby("Тип браслета", as_index=False).agg(
        Моделей=("SKU", "nunique"),
        Отгружено=("Отгружено", "sum"),
        **{
            "Скорость продаж": ("Скорость продаж", "sum"),
            "Продажи USD": ("Продажи USD", "sum"),
            "Расчетный остаток": ("Расчетный остаток", "sum"),
        },
    )
    result["Средняя цена USD"] = result["Продажи USD"] / result["Скорость продаж"].replace(0, pd.NA)
    result["Средняя цена USD"] = result["Средняя цена USD"].fillna(0)
    total = float(result["Скорость продаж"].sum())
    result["Доля продаж"] = result["Скорость продаж"] / total if total else 0.0
    return result.sort_values("Скорость продаж", ascending=False).reset_index(drop=True)


def bracelet_stone_summary(bracelets: pd.DataFrame) -> pd.DataFrame:
    """Aggregate stones separately inside each bracelet construction type."""
    if bracelets.empty:
        return pd.DataFrame()
    result = bracelets.groupby(["Тип браслета", "Камень"], as_index=False, dropna=False).agg(
        Моделей=("SKU", "nunique"),
        **{
            "Скорость продаж": ("Скорость продаж", "sum"),
            "Продажи USD": ("Продажи USD", "sum"),
            "Расчетный остаток": ("Расчетный остаток", "sum"),
        },
    )
    result["Средняя цена USD"] = result["Продажи USD"] / result["Скорость продаж"].replace(0, pd.NA)
    result["Средняя цена USD"] = result["Средняя цена USD"].fillna(0)
    totals = result.groupby("Тип браслета")["Скорость продаж"].transform("sum")
    result["Доля продаж внутри типа"] = result["Скорость продаж"] / totals.replace(0, pd.NA)
    result["Доля продаж внутри типа"] = result["Доля продаж внутри типа"].fillna(0)
    return result.sort_values(
        ["Тип браслета", "Скорость продаж", "Продажи USD"],
        ascending=[True, False, False],
    ).reset_index(drop=True)


def _money(value: float, digits: int = 0) -> str:
    return f"{float(value):,.{digits}f}".replace(",", " ")


def _kpi(label: str, value: str, note: str = "") -> None:
    st.markdown(
        f"""
        <div style="border:1px solid #e8e0d4;border-radius:16px;background:#fff;padding:18px;min-height:115px;box-shadow:0 8px 24px rgba(38,27,12,.05)">
          <div style="font-size:12px;text-transform:uppercase;letter-spacing:.08em;color:#8b6a36;font-weight:700">{label}</div>
          <div style="font-size:25px;font-weight:750;color:#17130f;margin-top:8px">{value}</div>
          <div style="font-size:12px;color:#777;margin-top:5px">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _locked_chart(figure: go.Figure, key: str) -> None:
    figure.update_layout(
        dragmode=False,
        clickmode="event",
        hovermode="closest",
        legend_itemclick=False,
        legend_itemdoubleclick=False,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Arial", color="#1c1813"),
    )
    figure.update_xaxes(fixedrange=True, automargin=True)
    figure.update_yaxes(fixedrange=True, automargin=True)
    st.plotly_chart(figure, width="stretch", key=key, config=LOCKED_CHART_CONFIG)


def _horizontal_chart(frame: pd.DataFrame, label: str, metric: str, title: str, suffix: str = "") -> go.Figure:
    data = frame.loc[frame[metric] > 0].sort_values(metric, ascending=True).copy()
    labels = [f"{_money(value)}{suffix}" for value in data[metric]]
    maximum = float(data[metric].max()) if not data.empty else 0.0
    fig = go.Figure(
        go.Bar(
            x=data[metric],
            y=data[label],
            orientation="h",
            marker_color="#b7893f",
            text=labels,
            textposition="outside",
            cliponaxis=False,
            hovertemplate="%{y}<br>%{text}<extra></extra>",
        )
    )
    fig.update_layout(
        title=title,
        height=max(340, len(data) * 40 + 100),
        margin=dict(l=20, r=130, t=55, b=35),
        xaxis=dict(range=[0, maximum * 1.3] if maximum else None, gridcolor="#ece8e1"),
        yaxis=dict(title=""),
    )
    return fig


def _table(frame: pd.DataFrame, key: str) -> None:
    config: dict[str, Any] = {}
    for column in frame.columns:
        if column in {
            "Продажи USD", "Средняя цена USD", "Средняя выручка в месяц USD",
            "Прогноз продаж USD",
        }:
            config[column] = st.column_config.NumberColumn(format="$%,.0f")
        elif column in {
            "Скорость продаж", "Расчетный остаток", "Отгружено", "Моделей",
            "Магазинов", "Рекомендованный заказ",
        }:
            config[column] = st.column_config.NumberColumn(format="%,.0f")
        elif column in {
            "Средние продажи в месяц", "Прогноз, шт.", "Ожидаемый остаток",
            "Потенциальный дефицит", "Целевой запас", "Запас, месяцев",
        }:
            config[column] = st.column_config.NumberColumn(format="%,.1f")
        elif column in {"Доля продаж", "Доля продаж внутри типа"}:
            config[column] = st.column_config.NumberColumn(format="percent")
        elif column == "Фото":
            config[column] = st.column_config.ImageColumn("Фото", width="small")
    st.dataframe(frame, width="stretch", hide_index=True, key=key, column_config=config)


def _persist_upload(uploaded_file) -> bytes | None:
    if uploaded_file is not None:
        current = bytes(uploaded_file.getvalue())
        previous = st.session_state.get("sonu_report_bytes")
        if previous != current:
            st.session_state["sonu_report_bytes"] = current
            st.session_state["sonu_report_name"] = uploaded_file.name
    return st.session_state.get("sonu_report_bytes")


def _render_store_section(frame: pd.DataFrame, rate: float) -> None:
    summary = aggregate_sonu(frame, ["Магазин"], rate).sort_values("Скорость продаж", ascending=False)
    st.markdown("### Магазины")
    left, right = st.columns(2)
    with left:
        _locked_chart(
            _horizontal_chart(summary, "Магазин", "Скорость продаж", "Продано за отчетный период", " шт."),
            "sonu_store_qty",
        )
    with right:
        _locked_chart(
            _horizontal_chart(summary, "Магазин", "Продажи USD", "Продажи по магазинам", " $"),
            "sonu_store_sales",
        )
    _table(summary, "sonu_store_table")


def _period_days(period: str) -> int:
    match = re.search(r"(\d{2}\.\d{2}\.\d{4})\s*[–—-]\s*(\d{2}\.\d{2}\.\d{4})", str(period))
    if not match:
        return 30
    try:
        start = datetime.strptime(match.group(1), "%d.%m.%Y")
        end = datetime.strptime(match.group(2), "%d.%m.%Y")
    except ValueError:
        return 30
    return max((end - start).days + 1, 1)


def _monthly_sales_matrix(frame: pd.DataFrame, rate: float, period_days: int) -> pd.DataFrame:
    data = aggregate_sonu(frame, ["Категория RU"], rate)
    factor = 30.0 / max(period_days, 1)
    data["Средние продажи в месяц"] = data["Скорость продаж"] * factor
    data["Средняя выручка в месяц USD"] = data["Продажи USD"] * factor
    data["Запас, месяцев"] = data["Расчетный остаток"] / data["Средние продажи в месяц"].replace(0, pd.NA)
    data["Запас, месяцев"] = data["Запас, месяцев"].fillna(0)
    return data.sort_values("Средние продажи в месяц", ascending=False).reset_index(drop=True)


@st.fragment
def _render_average_sales_section(frame: pd.DataFrame, rate: float, period_days: int) -> None:
    st.markdown("### Средние продажи")
    st.caption(
        f"Период отчета — {period_days} дн. Месячный темп приведен к 30 дням. "
        "Средняя цена считается взвешенно: общая выручка ÷ проданное количество."
    )
    summary = _monthly_sales_matrix(frame, rate, period_days)
    total_sold = float(frame["Скорость продаж"].sum())
    total_sales_usd = _usd(float(frame["Продажи VND"].sum()), rate)
    monthly_units = total_sold * 30 / max(period_days, 1)
    monthly_sales = total_sales_usd * 30 / max(period_days, 1)
    weighted_price = total_sales_usd / total_sold if total_sold else 0.0
    c1, c2, c3 = st.columns(3)
    with c1:
        _kpi("Средние продажи", f"{_money(monthly_units)} шт./мес.")
    with c2:
        _kpi("Средняя выручка", f"${_money(monthly_sales)}/мес.")
    with c3:
        _kpi("Взвешенная средняя цена", f"${_money(weighted_price)}")
    left, right = st.columns(2)
    with left:
        _locked_chart(
            _horizontal_chart(summary, "Категория RU", "Средние продажи в месяц", "Средние продажи по категориям", " шт./мес."),
            "sonu_average_sales_qty",
        )
    with right:
        _locked_chart(
            _horizontal_chart(summary, "Категория RU", "Средняя выручка в месяц USD", "Средняя выручка по категориям", " $/мес."),
            "sonu_average_sales_usd",
        )
    _table(
        summary[[
            "Категория RU", "Моделей", "Средние продажи в месяц",
            "Средняя выручка в месяц USD", "Средняя цена USD",
            "Расчетный остаток", "Запас, месяцев",
        ]],
        "sonu_average_sales_table",
    )


@st.fragment
def _render_forecast_section(frame: pd.DataFrame, rate: float, period_days: int) -> None:
    st.markdown("### Прогноз продаж")
    horizon = st.segmented_control(
        "Горизонт прогноза",
        [30, 60, 90],
        default=30,
        key="sonu_forecast_horizon",
    ) or 30
    data = aggregate_sonu(frame, ["Категория RU"], rate)
    factor = float(horizon) / max(period_days, 1)
    data["Прогноз, шт."] = data["Скорость продаж"] * factor
    data["Прогноз продаж USD"] = data["Продажи USD"] * factor
    data["Ожидаемый остаток"] = (data["Расчетный остаток"] - data["Прогноз, шт."]).clip(lower=0)
    data["Потенциальный дефицит"] = (data["Прогноз, шт."] - data["Расчетный остаток"]).clip(lower=0)
    data = data.sort_values("Прогноз, шт.", ascending=False).reset_index(drop=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        _kpi(f"Прогноз на {horizon} дней", f"{_money(data['Прогноз, шт.'].sum())} шт.")
    with c2:
        _kpi("Прогноз выручки", f"${_money(data['Прогноз продаж USD'].sum())}")
    with c3:
        _kpi("Потенциальный дефицит", f"{_money(data['Потенциальный дефицит'].sum())} шт.")
    _locked_chart(
        _horizontal_chart(data, "Категория RU", "Прогноз, шт.", f"Прогноз продаж на {horizon} дней", " шт."),
        f"sonu_forecast_{horizon}",
    )
    _table(data, f"sonu_forecast_table_{horizon}")


@st.fragment
def _render_order_matrix_section(frame: pd.DataFrame, rate: float, period_days: int) -> None:
    st.markdown("### Матрица заказа")
    target_months = st.slider(
        "Целевой запас после заказа, месяцев продаж",
        min_value=1.0,
        max_value=6.0,
        value=2.0,
        step=0.5,
        key="sonu_target_stock_months",
    )
    data = frame.groupby(["SKU", "Категория RU", "Камень", "Проба"], as_index=False, dropna=False).agg(
        Отгружено=("Отгружено", "sum"),
        **{
            "Скорость продаж": ("Скорость продаж", "sum"),
            "Продажи VND": ("Продажи VND", "sum"),
            "Расчетный остаток": ("Расчетный остаток", "sum"),
            "Магазинов": ("Магазин", "nunique"),
        },
    )
    factor = 30.0 / max(period_days, 1)
    data["Продажи USD"] = data["Продажи VND"] / rate
    data["Средние продажи в месяц"] = data["Скорость продаж"] * factor
    data["Целевой запас"] = data["Средние продажи в месяц"] * float(target_months)
    data["Рекомендованный заказ"] = (data["Целевой запас"] - data["Расчетный остаток"]).clip(lower=0).round().astype(int)
    data["Запас, месяцев"] = data["Расчетный остаток"] / data["Средние продажи в месяц"].replace(0, pd.NA)
    data["Запас, месяцев"] = data["Запас, месяцев"].fillna(0)
    data = data.drop(columns=["Продажи VND"]).sort_values(
        ["Рекомендованный заказ", "Средние продажи в месяц"], ascending=[False, False]
    )
    recommended = data.loc[data["Рекомендованный заказ"] > 0].copy()
    c1, c2, c3 = st.columns(3)
    with c1:
        _kpi("SKU к заказу", _money(recommended["SKU"].nunique()))
    with c2:
        _kpi("Рекомендовано", f"{_money(recommended['Рекомендованный заказ'].sum())} шт.")
    with c3:
        _kpi("Целевое покрытие", f"{target_months:g} мес.")
    st.caption(
        "Матрица использует расчетный остаток «отгружено − продано». "
        "До появления независимого фактического остатка заказ является ориентиром, а не финальным документом поставщику."
    )
    if recommended.empty:
        st.success("При выбранном целевом покрытии дефицитных моделей не найдено.")
    else:
        _table(
            recommended[[
                "SKU", "Категория RU", "Камень", "Проба", "Магазинов",
                "Скорость продаж", "Средние продажи в месяц", "Расчетный остаток",
                "Запас, месяцев", "Целевой запас", "Рекомендованный заказ", "Продажи USD",
            ]],
            "sonu_order_matrix_table",
        )


@st.fragment
def _render_recommendations_section(frame: pd.DataFrame, rate: float, period_days: int) -> None:
    st.markdown("### Рекомендации")
    monthly_units = float(frame["Скорость продаж"].sum()) * 30 / max(period_days, 1)
    remaining = float(frame["Расчетный остаток"].sum())
    coverage = remaining / monthly_units if monthly_units else 0.0
    matrix = frame.groupby(["SKU", "Категория RU", "Камень"], as_index=False, dropna=False).agg(
        **{
            "Скорость продаж": ("Скорость продаж", "sum"),
            "Расчетный остаток": ("Расчетный остаток", "sum"),
            "Продажи VND": ("Продажи VND", "sum"),
        }
    )
    matrix["Средние продажи в месяц"] = matrix["Скорость продаж"] * 30 / max(period_days, 1)
    matrix["Запас, месяцев"] = matrix["Расчетный остаток"] / matrix["Средние продажи в месяц"].replace(0, pd.NA)
    matrix["Запас, месяцев"] = matrix["Запас, месяцев"].fillna(0)
    matrix["Продажи USD"] = matrix["Продажи VND"] / rate
    shortage = matrix.loc[
        (matrix["Средние продажи в месяц"] > 0)
        & (matrix["Запас, месяцев"] < 1.0)
    ].sort_values(["Средние продажи в месяц", "Продажи USD"], ascending=False)
    slow = matrix.loc[
        (matrix["Расчетный остаток"] > 0)
        & ((matrix["Скорость продаж"] <= 0) | (matrix["Запас, месяцев"] > 6.0))
    ].sort_values("Расчетный остаток", ascending=False)

    if coverage >= 4:
        main_recommendation = (
            f"Расчетный запас покрывает около {coverage:.1f} месяца продаж. "
            "Крупный общий заказ не рекомендуется: формируйте точечную закупку только по дефицитным SKU."
        )
    elif coverage >= 2:
        main_recommendation = (
            f"Расчетный запас покрывает около {coverage:.1f} месяца продаж. "
            "Допустим ограниченный точечный заказ наиболее быстрых моделей."
        )
    else:
        main_recommendation = (
            f"Расчетный запас покрывает около {coverage:.1f} месяца продаж. "
            "Нужно проверить фактический остаток и подготовить приоритетный заказ по быстро продаваемым SKU."
        )
    st.markdown(
        f'<div style="border-left:4px solid #b7893f;background:#fffaf1;padding:16px 18px;border-radius:10px;margin:8px 0 18px">'
        f'<b>Главный вывод</b><br>{main_recommendation}</div>',
        unsafe_allow_html=True,
    )
    c1, c2, c3 = st.columns(3)
    with c1:
        _kpi("Запас", f"{coverage:.1f} мес.", "по расчетному остатку")
    with c2:
        _kpi("Дефицитных SKU", _money(len(shortage)), "покрытие менее месяца")
    with c3:
        _kpi("Медленных SKU", _money(len(slow)), "нет продаж или запас более 6 месяцев")
    left, right = st.columns(2)
    with left:
        st.markdown("#### Приоритет для точечного заказа")
        if shortage.empty:
            st.success("SKU с покрытием менее одного месяца не найдено.")
        else:
            _table(shortage.head(30).drop(columns=["Продажи VND"]), "sonu_shortage_recommendations")
    with right:
        st.markdown("#### Сначала перераспределить / реализовать")
        if slow.empty:
            st.success("Медленно оборачиваемые SKU по расчетным данным не найдены.")
        else:
            _table(slow.head(30).drop(columns=["Продажи VND"]), "sonu_slow_recommendations")


@st.fragment
def _render_category_section(frame: pd.DataFrame, rate: float) -> None:
    summary = aggregate_sonu(frame, ["Категория RU"], rate).sort_values("Скорость продаж", ascending=False)
    st.markdown("### Категории и расчетные остатки")
    left, right = st.columns(2)
    with left:
        _locked_chart(
            _horizontal_chart(summary, "Категория RU", "Скорость продаж", "Скорость продаж по категориям", " шт."),
            "sonu_category_sold",
        )
    with right:
        _locked_chart(
            _horizontal_chart(summary, "Категория RU", "Расчетный остаток", "Расчетный остаток по категориям", " шт."),
            "sonu_category_remaining",
        )
    _table(summary, "sonu_category_table")
    if float(summary["Расчетный остаток"].sum()) == 0:
        st.warning(
            "В текущем файле отгруженное количество полностью совпадает с проданным. "
            "Отдельной колонки фактического складского остатка в отчете нет, поэтому расчетный остаток равен нулю."
        )

    st.markdown("### Камни и пробы")
    mode = st.segmented_control(
        "Детализация",
        ["Камни", "Пробы"],
        default="Камни",
        key="sonu_category_detail",
    ) or "Камни"
    if mode == "Камни":
        detail = aggregate_sonu(frame, ["Камень"], rate).sort_values("Скорость продаж", ascending=False)
    else:
        detail = aggregate_sonu(frame, ["Проба"], rate).sort_values("Скорость продаж", ascending=False)
    _table(detail, f"sonu_{mode.lower()}_table")


@st.fragment
def _render_bracelet_section(frame: pd.DataFrame, rate: float) -> None:
    bracelets = classify_bracelets(frame, rate)
    if bracelets.empty:
        st.info("В отчете нет браслетов.")
        return
    summary = bracelet_type_summary(bracelets)
    st.markdown("### Браслеты: конструкция и продажи")
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        _kpi("Моделей браслетов", _money(bracelets["SKU"].nunique()))
    with k2:
        _kpi("Продано браслетов", f"{_money(bracelets['Скорость продаж'].sum())} шт.")
    with k3:
        _kpi("Продажи", f"${_money(bracelets['Продажи USD'].sum())}")
    with k4:
        auto_count = int((bracelets["Источник классификации"] == "Правило 50/50").sum())
        _kpi("Распределено 50/50", _money(auto_count), "неоднозначные или новые модели")

    left, right = st.columns(2)
    with left:
        _locked_chart(
            _horizontal_chart(summary, "Тип браслета", "Скорость продаж", "Продано по типу", " шт."),
            "sonu_bracelet_qty",
        )
    with right:
        _locked_chart(
            _horizontal_chart(summary, "Тип браслета", "Продажи USD", "Продажи по типу", " $"),
            "sonu_bracelet_sales",
        )
    _table(summary, "sonu_bracelet_type_table")

    st.markdown("### Камни внутри типов браслетов")
    stone_type = st.segmented_control(
        "Тип браслета для разбивки по камням",
        ["С затяжкой", "В круг с камнями"],
        default="С затяжкой",
        key="sonu_bracelet_stone_type",
    ) or "С затяжкой"
    stone_summary = bracelet_stone_summary(bracelets)
    stone_detail = stone_summary.loc[stone_summary["Тип браслета"] == stone_type].copy()
    if stone_detail.empty:
        st.info(f"Для типа «{stone_type}» нет данных по камням.")
    else:
        s1, s2, s3 = st.columns(3)
        with s1:
            _kpi("Камней", _money(stone_detail["Камень"].nunique()), stone_type)
        with s2:
            _kpi("Продано", f"{_money(stone_detail['Скорость продаж'].sum())} шт.", stone_type)
        with s3:
            _kpi("Продажи", f"${_money(stone_detail['Продажи USD'].sum())}", stone_type)
        left_stone, right_stone = st.columns(2)
        stone_key = "slider" if stone_type == "С затяжкой" else "circle"
        with left_stone:
            _locked_chart(
                _horizontal_chart(
                    stone_detail,
                    "Камень",
                    "Скорость продаж",
                    f"Камни · {stone_type} · продано",
                    " шт.",
                ),
                f"sonu_bracelet_stone_qty_{stone_key}",
            )
        with right_stone:
            _locked_chart(
                _horizontal_chart(
                    stone_detail,
                    "Камень",
                    "Продажи USD",
                    f"Камни · {stone_type} · продажи",
                    " $",
                ),
                f"sonu_bracelet_stone_sales_{stone_key}",
            )
        _table(
            stone_detail[
                [
                    "Камень", "Моделей", "Скорость продаж", "Продажи USD",
                    "Средняя цена USD", "Расчетный остаток", "Доля продаж внутри типа",
                ]
            ],
            f"sonu_bracelet_stone_table_{stone_key}",
        )

    st.markdown("### Модели браслетов")
    selected_type = st.selectbox(
        "Тип браслета",
        ["Все", "С затяжкой", "В круг с камнями"],
        key="sonu_bracelet_type_filter",
    )
    detail = bracelets.copy()
    if selected_type != "Все":
        detail = detail.loc[detail["Тип браслета"] == selected_type]
    detail = detail.sort_values(["Скорость продаж", "Продажи USD"], ascending=False)
    display_cols = [
        "Фото", "SKU", "Тип браслета", "Источник классификации", "Камень", "Проба",
        "Магазинов", "Скорость продаж", "Продажи USD", "Средняя цена USD", "Расчетный остаток",
    ]
    _table(detail[display_cols], "sonu_bracelet_model_table")

    image_rows = detail.loc[detail["Фото"].map(lambda value: isinstance(value, (bytes, bytearray)))].head(12)
    if not image_rows.empty:
        st.markdown("### Самые продаваемые модели с фото")
        columns = st.columns(4)
        for index, (_, row) in enumerate(image_rows.iterrows()):
            with columns[index % 4]:
                st.image(
                    row["Фото"],
                    caption=(
                        f"{row['SKU']}\n{row['Тип браслета']} · "
                        f"{_money(row['Скорость продаж'])} шт. · ${_money(row['Продажи USD'])}"
                    ),
                    width="stretch",
                )


@st.fragment
def _render_models_section(frame: pd.DataFrame, rate: float) -> None:
    data = add_usd_columns(frame, rate)
    st.markdown("### Детализация моделей")
    f1, f2, f3 = st.columns(3)
    with f1:
        stores = ["Все"] + sorted(data["Магазин"].dropna().astype(str).unique().tolist())
        store = st.selectbox("Магазин", stores, key="sonu_model_store")
    with f2:
        categories = ["Все"] + sorted(data["Категория RU"].dropna().astype(str).unique().tolist())
        category = st.selectbox("Категория", categories, key="sonu_model_category")
    with f3:
        stones = ["Все"] + sorted(data["Камень"].dropna().astype(str).unique().tolist())
        stone = st.selectbox("Камень", stones, key="sonu_model_stone")
    if store != "Все":
        data = data.loc[data["Магазин"] == store]
    if category != "Все":
        data = data.loc[data["Категория RU"] == category]
    if stone != "Все":
        data = data.loc[data["Камень"] == stone]
    columns = [
        "Магазин", "Категория RU", "SKU", "Камень", "Проба", "Скорость продаж",
        "Продажи USD", "Средняя цена USD", "Расчетный остаток",
    ]
    _table(data[columns].sort_values("Скорость продаж", ascending=False), "sonu_models_table")


def render_sonu_order_dashboard() -> None:
    """Streamlit entry point for the Sonu order analytics workspace."""
    st.markdown(
        """
        <section style="border:1px solid #e4d4bc;border-radius:22px;padding:26px 28px;background:linear-gradient(135deg,#fff 0%,#f8f1e7 100%);margin-bottom:18px">
          <div style="color:#b7893f;font-size:12px;font-weight:800;letter-spacing:.13em">SONU · ORDER ANALYTICS</div>
          <div style="font:42px Georgia,serif;color:#17130f;margin:6px 0">Заказ Sonu</div>
          <div style="color:#696158;max-width:850px">Продажи, магазины, категории, пробы и браслеты по специальной выгрузке 1С. Возвраты не учитываются. Скорость продаж — количество изделий, проданных за отчетный период.</div>
        </section>
        """,
        unsafe_allow_html=True,
    )

    rate = get_vnd_per_usd()
    st.sidebar.markdown("---")
    st.sidebar.markdown("**Заказ Sonu**")
    st.sidebar.markdown("[Сводка](#sonu-summary)")
    st.sidebar.markdown("[Аналитика](#sonu-detail)")

    uploaded = st.file_uploader(
        "Загрузите отчет Sonu",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
        key="sonu_upload_widget",
        help="Отчет должен быть отфильтрован по поставщику Sonu и содержать магазин, товар, камень, пробу и номенклатурную группу.",
    )
    file_bytes = _persist_upload(uploaded)
    if file_bytes is None:
        st.info(
            "Ожидается расширенная выгрузка 1С: Магазин → Товар → Камень/вставка → Проба → "
            "Номенклатурная группа → Поставщик."
        )
        return

    try:
        with st.spinner("Разбираем отчет и фотографии браслетов..."):
            report = cached_parse_sonu(file_bytes)
    except Exception as exc:
        st.error(str(exc))
        if st.button("Удалить загруженный файл Sonu", key="sonu_clear_invalid"):
            st.session_state.pop("sonu_report_bytes", None)
            st.session_state.pop("sonu_report_name", None)
            st.session_state.pop("sonu_upload_widget", None)
            st.rerun()
        return

    frame = report.data
    sold = float(frame["Скорость продаж"].sum())
    sales_usd = _usd(float(frame["Продажи VND"].sum()), rate)
    avg_usd = sales_usd / sold if sold else 0.0
    remaining = float(frame["Расчетный остаток"].sum())

    st.markdown('<div id="sonu-summary"></div>', unsafe_allow_html=True)
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k1:
        _kpi("Период", report.period)
    with k2:
        _kpi("Продано", f"{_money(sold)} шт.", "скорость продаж за период")
    with k3:
        _kpi("Продажи", f"${_money(sales_usd)}")
    with k4:
        _kpi("Средняя цена", f"${_money(avg_usd)}")
    with k5:
        _kpi("Расчетный остаток", f"{_money(remaining)} шт.", "отгружено − продано")
    with k6:
        _kpi("Моделей", _money(frame["SKU"].nunique()))

    st.caption(
        f"Файл: {st.session_state.get('sonu_report_name', 'Sonu.xlsx')} · Поставщик: {report.supplier} · "
        f"Курс: 1 USD = {_money(rate)} VND. Возвраты полностью исключены из расчетов."
    )
    if st.button("Загрузить другой отчет Sonu", key="sonu_replace_report"):
        st.session_state.pop("sonu_report_bytes", None)
        st.session_state.pop("sonu_report_name", None)
        st.session_state.pop("sonu_upload_widget", None)
        st.rerun()

    st.markdown('<div id="sonu-detail"></div>', unsafe_allow_html=True)
    st.markdown('<div class="block-navigation-title">Навигация по блокам</div>', unsafe_allow_html=True)
    sections = [
        "Продажи по магазинам",
        "Средние продажи",
        "Браслеты",
        "Прогноз продаж",
        "Матрица заказа",
        "Рекомендации",
    ]
    section = st.segmented_control(
        "Навигация по блокам Sonu",
        sections,
        default="Продажи по магазинам",
        key="sonu_section",
        label_visibility="collapsed",
    ) or "Продажи по магазинам"
    period_days = _period_days(report.period)
    if section == "Средние продажи":
        _render_average_sales_section(frame, rate, period_days)
    elif section == "Браслеты":
        _render_bracelet_section(frame, rate)
    elif section == "Прогноз продаж":
        _render_forecast_section(frame, rate, period_days)
    elif section == "Матрица заказа":
        _render_order_matrix_section(frame, rate, period_days)
    elif section == "Рекомендации":
        _render_recommendations_section(frame, rate, period_days)
    else:
        _render_store_section(frame, rate)
