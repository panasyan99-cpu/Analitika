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
from src.navigation import NavigationItem, render_mobile_navigation, render_sidebar
from src.report import COLORED_ORDER, PEARL_ORDER, TOP_ORDER, classify
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
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



SONU_SECTIONS = [
    ("sonu-stores", "Продажи по магазинам"),
    ("sonu-average-sales", "Средние продажи"),
    ("sonu-stones", "Камни"),
    ("sonu-bracelets", "Браслеты"),
    ("sonu-models", "Модели"),
]

STONE_GROUP_LABELS = {
    "TOP STONES": "Top Stones",
    "PEARLS": "Pearls",
    "COLORED STONES": "Other Stones",
}
STONE_GROUP_ORDER = ["Top Stones", "Pearls", "Other Stones"]
STONE_MEMBER_ORDER = {
    "Top Stones": TOP_ORDER,
    "Pearls": PEARL_ORDER,
    "Other Stones": COLORED_ORDER,
}

# Сокращения, используемые в артикулах и выгрузках Sonu. Порядок важен:
# он повторяет бизнес-приоритеты общей аналитики (Moissanite → Sapphire → Ruby...).
SONU_STONE_ALIASES: tuple[tuple[str, str, str], ...] = (
    (r"MOIS|MOISSANITE|MOISANITE", "MOISSANITE", "MOIS"),
    (r"SAPPHIRE|SAPPHRIE|(?:^|[-_/\s])BS(?:$|[-_/\s])", "SAPPHIRE", "BS"),
    (r"RUBY", "RUBY", "RUBY"),
    (r"LONDON|(?:^|[-_/\s])LBT(?:$|[-_/\s])", "LONDON TOPAZ", "LBT"),
    (r"SWISS|(?:^|[-_/\s])SWBT(?:$|[-_/\s])", "SWISS TOPAZ", "SWBT"),
    (r"WHITE TOPAZ|BLUE TOPAZ|(?:^|[-_/\s])W?BT(?:$|[-_/\s])", "BLUE TOPAZ", "BT"),
    (r"CREATED EMERALD|(?:^|[-_/\s])CE(?:$|[-_/\s])", "CREATED EMERALD", "CE"),
    (r"CHROME DIOPSIDE|(?:^|[-_/\s])CD(?:$|[-_/\s])", "CHROME DIOPSIDE", "CD"),
    (r"EMERALD|(?:^|[-_/\s])EM(?:$|[-_/\s])", "EMERALD", "EM"),
    (r"PERIDOT|(?:^|[-_/\s])PERI(?:$|[-_/\s])", "PERIDOT", "PERI"),
    (r"BAROQUE", "BAROQUE PEARL", "BAROQUE"),
    (r"AKOYA", "AKOYA", "AKOYA"),
    (r"TAHITI|TAHITIAN|(?:^|[-_/\s])TAH(?:$|[-_/\s])", "TAHITI", "TAH"),
    (r"SOUTH SEA|(?:^|[-_/\s])SSP(?:$|[-_/\s])", "SOUTH SEA PEARL", "SSP"),
    (r"FRESHWATER PEARL (?:PINK|ROSE|GRAY|GREY|BLACK)|(?:^|[-_/\s])FPC(?:$|[-_/\s])", "FRESHWATER PEARL PINK", "FPC"),
    (r"FRESHWATER|(?:^|[-_/\s])FPW(?:$|[-_/\s])|(?:^|[-_/\s])FWP(?:$|[-_/\s])", "FRESHWATER PEARL WHITE", "FPW"),
    (r"BLACK SPINEL|(?:^|[-_/\s])BSP(?:$|[-_/\s])", "BLACK SPINEL", "BSP"),
    (r"ONYX", "ONYX", "ONYX"),
    (r"OBSIDIAN", "OBSIDIAN", "OBSIDIAN"),
    (r"AMETHYST|(?:^|[-_/\s])AMST(?:$|[-_/\s])", "AMETHYST", "AMST"),
    (r"MYSTIC", "MYSTIC QUARTZ", "MYSTIC"),
    (r"CITRINE|(?:^|[-_/\s])CIT(?:$|[-_/\s])", "CITRINE", "CIT"),
    (r"SMOKY|SMOKEY|RAUCH|HONEY", "SMOKY QUARTZ", "SMOKY"),
    (r"QUARTZ", "QUARTZ", "QUARTZ"),
    (r"GARNET|(?:^|[-_/\s])GARN(?:$|[-_/\s])", "GARNET", "GARN"),
    (r"RHODOLITE|RODOLITE|(?:^|[-_/\s])RDT(?:$|[-_/\s])", "RHODOLITE", "RDT"),
    (r"GREEN AGATE", "GREEN AGATE", "GREEN AGATE"),
    (r"AGATE|(?:^|[-_/\s])AGAT(?:$|[-_/\s])", "AGATE", "AGATE"),
    (r"APATITE|(?:^|[-_/\s])APAT(?:$|[-_/\s])", "APATITE", "APAT"),
    (r"TANZANITE|(?:^|[-_/\s])TANZ(?:$|[-_/\s])", "TANZANITE", "TANZ"),
    (r"IOLITE|(?:^|[-_/\s])IOL(?:$|[-_/\s])", "IOLITE", "IOL"),
    (r"OPAL", "OPAL", "OPAL"),
    (r"AMBER", "AMBER", "AMBER"),
    (r"LAPIS", "LAPIS LAZURITE", "LAPIS"),
    (r"TURQUOISE", "TURQUOISE", "TURQUOISE"),
)



def sonu_navigation_items(has_report: bool) -> list[NavigationItem]:
    """Return the complete Sonu menu even before a workbook is uploaded."""
    definitions = [
        ("sonu-upload", "Загрузка отчета", "#sonu-upload", True),
        ("sonu-summary", "Сводка", "#sonu-summary", has_report),
        *[(anchor, label, f"#{anchor}", has_report) for anchor, label in SONU_SECTIONS],
        ("sonu-export", "Полная выгрузка", "#sonu-export", has_report),
        ("about", "О программе", "#about", True),
    ]
    return [
        NavigationItem(item_id=item_id, label=label, href=href, enabled=enabled)
        for item_id, label, href, enabled in definitions
    ]


def _sonu_sidebar_navigation(
    has_report: bool,
    *,
    status_text: str | None = None,
    status_tone: str | None = None,
    action_label: str | None = None,
    action_key: str | None = None,
):
    """Render Sonu through the same sidebar shell as every other workspace."""
    return render_sidebar(
        module_title="Заказ Sonu",
        navigation_title="Навигация по отчету",
        items=sonu_navigation_items(has_report),
        status_text=status_text or ("Отчет Sonu загружен" if has_report else "Ожидается отчет Sonu"),
        status_tone=(status_tone or ("success" if has_report else "neutral")),
        source_text="Источник: Excel · отчет Sonu",
        action_label=action_label,
        action_key=action_key,
    )


def _sonu_mobile_navigation(has_report: bool) -> None:
    """Sticky touch navigation using the shared component and states."""
    render_mobile_navigation(sonu_navigation_items(has_report))


def _anchor(anchor: str) -> None:
    st.markdown(f'<div id="{anchor}" class="report-anchor"></div>', unsafe_allow_html=True)


def _safe_excel_frame(frame: pd.DataFrame) -> pd.DataFrame:
    """Remove binary image payloads and normalize missing values before Excel export."""
    result = frame.copy()
    if "Фото" in result.columns:
        result["Фото"] = result["Фото"].map(
            lambda value: "Есть" if isinstance(value, (bytes, bytearray)) else ""
        )
    return result.replace({pd.NA: None})


def _classify_sonu_stone(raw_stone: Any, sku: Any) -> tuple[str, str, str]:
    """Map Sonu names/SKU abbreviations to the shared business stone groups."""
    raw = str(raw_stone or "").strip()
    sku_text = str(sku or "").strip()
    combined = re.sub(r"\s+", " ", f"{raw} {sku_text}".upper()).strip()
    for pattern, expanded_name, abbreviation in SONU_STONE_ALIASES:
        if re.search(pattern, combined):
            segment, member, _ = classify(expanded_name)
            return STONE_GROUP_LABELS.get(segment, segment), member, abbreviation

    if raw and raw.casefold() not in {"не указан", "не указана", "none", "nan"}:
        segment, member, _ = classify(raw)
        return STONE_GROUP_LABELS.get(segment, segment), member, raw
    return "Other Stones", "Other Colored Stones", "Не определено"


def add_stone_classification(frame: pd.DataFrame) -> pd.DataFrame:
    """Add group/member columns without replacing the original stone name."""
    result = frame.copy()
    classified = result.apply(
        lambda row: _classify_sonu_stone(row.get("Камень"), row.get("SKU")), axis=1
    )
    result[["Группа камня", "Камень группы", "Сокращение"]] = pd.DataFrame(
        classified.tolist(), index=result.index
    )
    return result


def stone_group_summary(frame: pd.DataFrame, rate: float) -> pd.DataFrame:
    enriched = add_stone_classification(frame)
    result = enriched.groupby("Группа камня", as_index=False, dropna=False).agg(
        Моделей=("SKU", "nunique"),
        Магазинов=("Магазин", "nunique"),
        **{
            "Скорость продаж": ("Скорость продаж", "sum"),
            "Продажи VND": ("Продажи VND", "sum"),
        },
    )
    result["Продажи USD"] = result["Продажи VND"] / rate
    result["Средняя цена USD"] = result["Продажи USD"] / result["Скорость продаж"].replace(0, pd.NA)
    result["Средняя цена USD"] = result["Средняя цена USD"].fillna(0)
    total_qty = float(result["Скорость продаж"].sum())
    total_sales = float(result["Продажи USD"].sum())
    result["Доля количества"] = result["Скорость продаж"] / total_qty if total_qty else 0.0
    result["Доля выручки"] = result["Продажи USD"] / total_sales if total_sales else 0.0
    order = {name: index for index, name in enumerate(STONE_GROUP_ORDER)}
    result["_order"] = result["Группа камня"].map(order).fillna(99)
    return result.drop(columns=["Продажи VND"]).sort_values("_order").drop(columns="_order").reset_index(drop=True)


def stone_member_summary(frame: pd.DataFrame, rate: float) -> pd.DataFrame:
    enriched = add_stone_classification(frame)

    def abbreviations(values: pd.Series) -> str:
        unique = sorted({str(value).strip() for value in values if str(value).strip()})
        return ", ".join(unique[:8])

    result = enriched.groupby(
        ["Группа камня", "Камень группы"], as_index=False, dropna=False
    ).agg(
        Сокращения=("Сокращение", abbreviations),
        Моделей=("SKU", "nunique"),
        Магазинов=("Магазин", "nunique"),
        **{
            "Скорость продаж": ("Скорость продаж", "sum"),
            "Продажи VND": ("Продажи VND", "sum"),
        },
    )
    result["Продажи USD"] = result["Продажи VND"] / rate
    result["Средняя цена USD"] = result["Продажи USD"] / result["Скорость продаж"].replace(0, pd.NA)
    result["Средняя цена USD"] = result["Средняя цена USD"].fillna(0)
    group_qty = result.groupby("Группа камня")["Скорость продаж"].transform("sum")
    group_sales = result.groupby("Группа камня")["Продажи USD"].transform("sum")
    result["Доля количества внутри группы"] = result["Скорость продаж"] / group_qty.replace(0, pd.NA)
    result["Доля выручки внутри группы"] = result["Продажи USD"] / group_sales.replace(0, pd.NA)
    result[["Доля количества внутри группы", "Доля выручки внутри группы"]] = result[
        ["Доля количества внутри группы", "Доля выручки внутри группы"]
    ].fillna(0)
    group_order = {name: index for index, name in enumerate(STONE_GROUP_ORDER)}
    member_order = {
        (group, member): index
        for group, members in STONE_MEMBER_ORDER.items()
        for index, member in enumerate(members)
    }
    result["_group_order"] = result["Группа камня"].map(group_order).fillna(99)
    result["_member_order"] = [
        member_order.get((group, member), 99)
        for group, member in zip(result["Группа камня"], result["Камень группы"])
    ]
    return result.drop(columns=["Продажи VND"]).sort_values(
        ["_group_order", "_member_order", "Продажи USD"], ascending=[True, True, False]
    ).drop(columns=["_group_order", "_member_order"]).reset_index(drop=True)


def stone_store_summary(frame: pd.DataFrame, rate: float) -> pd.DataFrame:
    enriched = add_stone_classification(frame)
    result = enriched.groupby(
        ["Магазин", "Группа камня"], as_index=False, dropna=False
    ).agg(
        Моделей=("SKU", "nunique"),
        **{
            "Скорость продаж": ("Скорость продаж", "sum"),
            "Продажи VND": ("Продажи VND", "sum"),
        },
    )
    result["Продажи USD"] = result["Продажи VND"] / rate
    return result.drop(columns="Продажи VND").sort_values(
        ["Магазин", "Продажи USD"], ascending=[True, False]
    ).reset_index(drop=True)


@st.cache_data(show_spinner=False, ttl=1800, max_entries=4)
def build_full_sonu_export(
    frame: pd.DataFrame,
    period: str,
    supplier: str,
    rate: float,
) -> bytes:
    """Build a complete sales workbook without stock-based recommendations."""
    period_days = _period_days(period)
    stores = aggregate_sonu(frame, ["Магазин"], rate).sort_values("Скорость продаж", ascending=False)
    averages = _monthly_sales_matrix(frame, rate, period_days)
    groups = stone_group_summary(frame, rate)
    members = stone_member_summary(frame, rate)
    stone_stores = stone_store_summary(frame, rate)
    bracelets = classify_bracelets(frame, rate)
    bracelet_types = bracelet_type_summary(bracelets)
    bracelet_stones = bracelet_stone_summary(bracelets)
    enriched = add_stone_classification(frame)
    models = add_usd_columns(enriched, rate).sort_values(
        ["Скорость продаж", "Продажи USD"], ascending=False
    )

    sold = float(frame["Скорость продаж"].sum())
    sales_usd = _usd(float(frame["Продажи VND"].sum()), rate)
    monthly_units = sold * 30 / max(period_days, 1)
    monthly_sales = sales_usd * 30 / max(period_days, 1)
    summary = pd.DataFrame(
        [
            ("Период", period),
            ("Дней в периоде", period_days),
            ("Поставщик", supplier),
            ("Курс VND за 1 USD", rate),
            ("Продано, шт.", sold),
            ("Продажи, USD", sales_usd),
            ("Взвешенная средняя цена, USD", sales_usd / sold if sold else 0.0),
            ("Средние продажи, шт./мес.", monthly_units),
            ("Средняя выручка, USD/мес.", monthly_sales),
            ("Моделей", int(frame["SKU"].nunique())),
            ("Магазинов", int(frame["Магазин"].nunique())),
        ],
        columns=["Показатель", "Значение"],
    )

    source_columns = [column for column in models.columns if column != "Расчетный остаток"]
    sheets: list[tuple[str, pd.DataFrame]] = [
        ("Сводка", summary),
        ("Магазины", stores),
        ("Средние продажи", averages),
        ("Группы камней", groups),
        ("Камни по группам", members),
        ("Камни по магазинам", stone_stores),
        ("Типы браслетов", bracelet_types),
        ("Модели браслетов", bracelets),
        ("Камни браслетов", bracelet_stones),
        ("Все модели", models[source_columns]),
        ("Исходные данные", models[source_columns]),
    ]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, sheet_frame in sheets:
            _safe_excel_frame(sheet_frame).to_excel(writer, sheet_name=sheet_name[:31], index=False)

        workbook = writer.book
        header_fill = PatternFill("solid", fgColor="2B2115")
        header_font = Font(color="F6D899", bold=True)
        thin_gold = Side(style="thin", color="C59A52")
        border = Border(bottom=thin_gold)
        money_columns = {
            "Продажи USD", "Средняя цена USD", "Средняя выручка в месяц USD",
        }
        decimal_columns = {"Средние продажи в месяц"}
        percent_columns = {
            "Доля продаж", "Доля количества", "Доля выручки",
            "Доля продаж внутри типа", "Доля количества внутри группы",
            "Доля выручки внутри группы",
        }

        for worksheet in workbook.worksheets:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            worksheet.sheet_view.showGridLines = False
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = border
            worksheet.row_dimensions[1].height = 30
            headers = {cell.column: str(cell.value or "") for cell in worksheet[1]}
            for column_idx, header in headers.items():
                max_length = len(header)
                for cell in worksheet[get_column_letter(column_idx)][1:]:
                    value = cell.value
                    max_length = max(max_length, len(str(value)) if value is not None else 0)
                    if header in money_columns or "USD" in header:
                        cell.number_format = '$#,##0.00'
                    elif header in percent_columns:
                        cell.number_format = '0.0%'
                    elif header in decimal_columns or "месяц" in header.lower():
                        cell.number_format = '0.0'
                    elif isinstance(value, (int, float)):
                        cell.number_format = '#,##0.00' if isinstance(value, float) and not float(value).is_integer() else '#,##0'
                worksheet.column_dimensions[get_column_letter(column_idx)].width = min(max(max_length + 2, 12), 42)

    output.seek(0)
    return output.getvalue()


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
    bracelets = add_stone_classification(frame.loc[frame["Категория"] == "Bracelet"].copy())
    if bracelets.empty:
        return pd.DataFrame()
    grouped = bracelets.groupby("SKU", as_index=False).agg(
        Магазинов=("Магазин", "nunique"),
        Камень=("Камень", "first"),
        **{
            "Группа камня": ("Группа камня", "first"),
            "Камень группы": ("Камень группы", "first"),
            "Сокращение": ("Сокращение", "first"),
        },
        Проба=("Проба", "first"),
        Отгружено=("Отгружено", "sum"),
        **{
            "Скорость продаж": ("Скорость продаж", "sum"),
            "Продажи VND": ("Продажи VND", "sum"),
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
    # Неоднозначные модели делятся 50/50; лишняя и верхняя по продажам половина
    # относятся к браслетам с затяжкой.
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
        },
    )
    result["Средняя цена USD"] = result["Продажи USD"] / result["Скорость продаж"].replace(0, pd.NA)
    result["Средняя цена USD"] = result["Средняя цена USD"].fillna(0)
    total = float(result["Скорость продаж"].sum())
    result["Доля продаж"] = result["Скорость продаж"] / total if total else 0.0
    return result.sort_values("Скорость продаж", ascending=False).reset_index(drop=True)


def bracelet_stone_summary(bracelets: pd.DataFrame) -> pd.DataFrame:
    """Aggregate standardized stone groups inside each bracelet construction type."""
    if bracelets.empty:
        return pd.DataFrame()
    result = bracelets.groupby(
        ["Тип браслета", "Группа камня", "Камень группы"],
        as_index=False,
        dropna=False,
    ).agg(
        Сокращения=("Сокращение", lambda values: ", ".join(sorted({str(v) for v in values if str(v).strip()}))),
        Моделей=("SKU", "nunique"),
        **{
            "Скорость продаж": ("Скорость продаж", "sum"),
            "Продажи USD": ("Продажи USD", "sum"),
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
        }:
            config[column] = st.column_config.NumberColumn(format="$%,.0f")
        elif column in {
            "Скорость продаж", "Отгружено", "Моделей", "Магазинов",
        }:
            config[column] = st.column_config.NumberColumn(format="%,.0f")
        elif column in {"Средние продажи в месяц"}:
            config[column] = st.column_config.NumberColumn(format="%,.1f")
        elif column in {
            "Доля продаж", "Доля продаж внутри типа", "Доля количества",
            "Доля выручки", "Доля количества внутри группы", "Доля выручки внутри группы",
        }:
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
            "Средняя выручка в месяц USD", "Средняя цена USD", "Доля продаж",
        ]],
        "sonu_average_sales_table",
    )


@st.fragment
def _render_stones_section(frame: pd.DataFrame, rate: float) -> None:
    st.markdown("### Камни по нашим группам")
    st.caption(
        "Названия и сокращения из выгрузки и SKU приводятся к единому стандарту: "
        "Top Stones, Pearls и Other Stones, затем — к участнику соответствующей группы."
    )
    groups = stone_group_summary(frame, rate)
    members = stone_member_summary(frame, rate)

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        _kpi("Групп", _money(groups["Группа камня"].nunique()))
    with k2:
        _kpi("Участников групп", _money(members["Камень группы"].nunique()))
    with k3:
        _kpi("Продано", f"{_money(groups['Скорость продаж'].sum())} шт.")
    with k4:
        _kpi("Продажи", f"${_money(groups['Продажи USD'].sum())}")

    left, right = st.columns(2)
    with left:
        _locked_chart(
            _horizontal_chart(groups, "Группа камня", "Скорость продаж", "Группы камней · количество", " шт."),
            "sonu_stone_group_qty",
        )
    with right:
        _locked_chart(
            _horizontal_chart(groups, "Группа камня", "Продажи USD", "Группы камней · продажи", " $"),
            "sonu_stone_group_sales",
        )
    _table(groups, "sonu_stone_group_table")

    selected_group = st.segmented_control(
        "Участники группы",
        STONE_GROUP_ORDER,
        default="Top Stones",
        key="sonu_stone_group_selected",
    ) or "Top Stones"
    detail = members.loc[members["Группа камня"] == selected_group].copy()
    if detail.empty:
        st.info(f"В группе «{selected_group}» продаж нет.")
        return

    d1, d2, d3 = st.columns(3)
    with d1:
        _kpi("Участников", _money(detail["Камень группы"].nunique()), selected_group)
    with d2:
        _kpi("Продано", f"{_money(detail['Скорость продаж'].sum())} шт.", selected_group)
    with d3:
        _kpi("Продажи", f"${_money(detail['Продажи USD'].sum())}", selected_group)

    left_detail, right_detail = st.columns(2)
    safe_key = re.sub(r"[^a-z]+", "_", selected_group.lower()).strip("_")
    with left_detail:
        _locked_chart(
            _horizontal_chart(detail, "Камень группы", "Скорость продаж", f"{selected_group} · количество", " шт."),
            f"sonu_stone_member_qty_{safe_key}",
        )
    with right_detail:
        _locked_chart(
            _horizontal_chart(detail, "Камень группы", "Продажи USD", f"{selected_group} · продажи", " $"),
            f"sonu_stone_member_sales_{safe_key}",
        )
    _table(
        detail[[
            "Камень группы", "Сокращения", "Моделей", "Магазинов",
            "Скорость продаж", "Продажи USD", "Средняя цена USD",
            "Доля количества внутри группы", "Доля выручки внутри группы",
        ]],
        f"sonu_stone_member_table_{safe_key}",
    )


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
            _kpi("Камней", _money(stone_detail["Камень группы"].nunique()), stone_type)
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
                    "Камень группы",
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
                    "Камень группы",
                    "Продажи USD",
                    f"Камни · {stone_type} · продажи",
                    " $",
                ),
                f"sonu_bracelet_stone_sales_{stone_key}",
            )
        _table(
            stone_detail[
                [
                    "Группа камня", "Камень группы", "Сокращения", "Моделей",
                    "Скорость продаж", "Продажи USD", "Средняя цена USD",
                    "Доля продаж внутри типа",
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
        "Фото", "SKU", "Тип браслета", "Источник классификации", "Группа камня",
        "Камень группы", "Сокращение", "Камень", "Проба", "Магазинов",
        "Скорость продаж", "Продажи USD", "Средняя цена USD",
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
    data = add_usd_columns(add_stone_classification(frame), rate)
    st.markdown("### Детализация моделей")
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        stores = ["Все"] + sorted(data["Магазин"].dropna().astype(str).unique().tolist())
        store = st.selectbox("Магазин", stores, key="sonu_model_store")
    with f2:
        categories = ["Все"] + sorted(data["Категория RU"].dropna().astype(str).unique().tolist())
        category = st.selectbox("Категория", categories, key="sonu_model_category")
    with f3:
        groups = ["Все"] + STONE_GROUP_ORDER
        stone_group = st.selectbox("Группа камня", groups, key="sonu_model_stone_group")
    with f4:
        available_members = sorted(data["Камень группы"].dropna().astype(str).unique().tolist())
        member = st.selectbox("Камень", ["Все"] + available_members, key="sonu_model_stone")
    if store != "Все":
        data = data.loc[data["Магазин"] == store]
    if category != "Все":
        data = data.loc[data["Категория RU"] == category]
    if stone_group != "Все":
        data = data.loc[data["Группа камня"] == stone_group]
    if member != "Все":
        data = data.loc[data["Камень группы"] == member]
    columns = [
        "Магазин", "Категория RU", "SKU", "Группа камня", "Камень группы",
        "Сокращение", "Камень", "Проба", "Скорость продаж", "Продажи USD",
        "Средняя цена USD",
    ]
    _table(data[columns].sort_values(["Скорость продаж", "Продажи USD"], ascending=False), "sonu_models_table")


def render_sonu_order_dashboard() -> None:
    """Streamlit entry point for the complete Sonu order analytics workspace."""
    rate = get_vnd_per_usd()
    _anchor("sonu-upload")
    uploaded = st.file_uploader(
        "Загрузите отчет Sonu",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
        key="sonu_upload_widget",
        help="Отчет должен быть отфильтрован по поставщику Sonu и содержать магазин, товар, камень, пробу и номенклатурную группу.",
    )
    file_bytes = _persist_upload(uploaded)
    if file_bytes is None:
        _sonu_sidebar_navigation(False)
        _sonu_mobile_navigation(False)
        st.info(
            "Ожидается расширенная выгрузка 1С: Магазин → Товар → Камень/вставка → Проба → "
            "Номенклатурная группа → Поставщик."
        )
        return

    try:
        with st.spinner("Разбираем отчет и фотографии браслетов..."):
            report = cached_parse_sonu(file_bytes)
    except Exception as exc:
        navigation = _sonu_sidebar_navigation(
            False,
            status_text="Файл Sonu не распознан",
            status_tone="error",
            action_label="Удалить загруженный файл",
            action_key="sonu_clear_invalid",
        )
        _sonu_mobile_navigation(False)
        st.error(str(exc))
        if navigation.action_clicked:
            st.session_state.pop("sonu_report_bytes", None)
            st.session_state.pop("sonu_report_name", None)
            st.session_state.pop("sonu_upload_widget", None)
            st.rerun()
        return

    navigation = _sonu_sidebar_navigation(
        True,
        action_label="Загрузить другой отчет",
        action_key="sonu_replace_report",
    )
    _sonu_mobile_navigation(True)
    if navigation.action_clicked:
        st.session_state.pop("sonu_report_bytes", None)
        st.session_state.pop("sonu_report_name", None)
        st.session_state.pop("sonu_upload_widget", None)
        st.rerun()
    frame = report.data
    sold = float(frame["Скорость продаж"].sum())
    sales_usd = _usd(float(frame["Продажи VND"].sum()), rate)
    avg_usd = sales_usd / sold if sold else 0.0
    period_days = _period_days(report.period)

    _anchor("sonu-summary")
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        _kpi("Период", report.period)
    with k2:
        _kpi("Продано", f"{_money(sold)} шт.", "скорость продаж за период")
    with k3:
        _kpi("Продажи", f"${_money(sales_usd)}")
    with k4:
        _kpi("Средняя цена", f"${_money(avg_usd)}")
    with k5:
        _kpi("Моделей", _money(frame["SKU"].nunique()))

    st.caption(
        f"Файл: {st.session_state.get('sonu_report_name', 'Sonu.xlsx')} · Поставщик: {report.supplier} · "
        f"Курс: 1 USD = {_money(rate)} VND. Возвраты полностью исключены из расчетов."
    )
    # The full report is rendered from top to bottom. Sidebar buttons only move
    # the viewport and never hide or rebuild analytical blocks.
    _anchor("sonu-stores")
    _render_store_section(frame, rate)

    _anchor("sonu-average-sales")
    _render_average_sales_section(frame, rate, period_days)

    _anchor("sonu-stones")
    _render_stones_section(frame, rate)

    _anchor("sonu-bracelets")
    _render_bracelet_section(frame, rate)

    _anchor("sonu-models")
    _render_models_section(frame, rate)

    _anchor("sonu-export")
    st.markdown("### Полная выгрузка Sonu")
    export_bytes = build_full_sonu_export(
        frame,
        report.period,
        report.supplier,
        rate,
    )
    safe_period = re.sub(r"[^0-9]+", "_", report.period).strip("_") or "period"
    st.download_button(
        "Скачать полный отчет Sonu",
        data=export_bytes,
        file_name=f"Sonu_full_report_{safe_period}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        key="sonu_full_export",
        help=(
            "Один Excel-файл со сводкой, магазинами, средними продажами, группами и "
            "участниками камней, браслетами, моделями и исходными данными."
        ),
    )
    st.caption(
        "Выгрузка формируется полностью: группы Top Stones, Pearls и Other Stones, "
        "участники групп, магазины, средние продажи, браслеты, все модели и исходные данные."
    )

