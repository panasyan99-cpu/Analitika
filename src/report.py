from __future__ import annotations

import calendar
import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

APP_VERSION = "1.1.0"

SEGMENTS = ("TOP STONES", "PEARLS", "OTHER STONES")
TOP_ORDER = (
    "Blue Sapphire", "Ruby", "Moissanite", "London Topaz",
    "Swiss Topaz", "Other Topaz", "Green Stones",
)
PEARL_ORDER = (
    "Round White Freshwater Pearl", "White Freshwater Pearl",
    "Baroque Pearl", "Colored Freshwater Pearl", "Sea Pearl",
)
PRODUCT_ORDER = (
    "Earrings", "Ring", "Pendant", "Bracelet", "Necklace", "Brooch",
    "Pearl Necklace", "Pearl Bracelet", "Pearl Chain", "Chain", "Other",
)
PRODUCT_ALIASES = {
    "EARRINGS": "Earrings", "EARRING": "Earrings", "RING": "Ring",
    "PENDANT": "Pendant", "BRACELET": "Bracelet", "NECKLACE": "Necklace",
    "BROOCH": "Brooch", "PEARL NECKLACE": "Pearl Necklace",
    "PEARL BRACELET": "Pearl Bracelet", "PEARL CHAIN": "Pearl Chain",
    "CHAIN": "Chain", "OTHER": "Other",
}
SEA_WORDS = ("SEA PEARL", "AKOYA", "TAHITI", "TAHITIAN", "SOUTH SEA", "MABE", "GALATEA", "FACETED")
COLOR_WORDS = (
    "ROSE", "PINK", "GRAY", "GREY", "BLACK", "PURPLE", "LAVENDER",
    "PEACH", "GOLD", "GOLDEN", "MULTI", "COLORED", "COLOUR", "BLUE",
)

# Ordered aliases. Earlier entries have higher priority.
RULES: list[tuple[str, str, str]] = [
    # Top stones
    ("MOISSANITE", "TOP STONES", "Moissanite"),
    ("MOISANITE", "TOP STONES", "Moissanite"),
    ("MOSSANITE", "TOP STONES", "Moissanite"),
    ("MUSSONITE", "TOP STONES", "Moissanite"),
    ("MUSANITE", "TOP STONES", "Moissanite"),
    ("SAPPHIRE", "TOP STONES", "Blue Sapphire"),
    ("SAPPHRIE", "TOP STONES", "Blue Sapphire"),
    ("SAPPHIRE", "TOP STONES", "Blue Sapphire"),
    ("RUBY", "TOP STONES", "Ruby"),
    ("LONDON BLUE TOPAZ", "TOP STONES", "London Topaz"),
    ("LONDON TOPAZ", "TOP STONES", "London Topaz"),
    ("LONDON BT", "TOP STONES", "London Topaz"),
    ("LONDON BLUE T", "TOP STONES", "London Topaz"),
    ("SWISS BLUE TOPAZ", "TOP STONES", "Swiss Topaz"),
    ("SWISS TOPAZ", "TOP STONES", "Swiss Topaz"),
    ("SWISS BT", "TOP STONES", "Swiss Topaz"),
    ("SWISS BLUE T", "TOP STONES", "Swiss Topaz"),
    ("SIVS TOPAZ", "TOP STONES", "Swiss Topaz"),
    ("VISTOPAZ", "TOP STONES", "Swiss Topaz"),
    ("MLBT", "TOP STONES", "Other Topaz"),
    ("MULTI BLUE TOPAZ", "TOP STONES", "Other Topaz"),
    ("MULTI BT", "TOP STONES", "Other Topaz"),
    ("SKY BLUE TOPAZ", "TOP STONES", "Other Topaz"),
    ("SKY TOPAZ", "TOP STONES", "Other Topaz"),
    ("WHITE TOPAZ", "TOP STONES", "Other Topaz"),
    ("BLUE TOPAZ", "TOP STONES", "Other Topaz"),
    ("CREATED EMERALD", "TOP STONES", "Green Stones"),
    ("CREATE EMERALD", "TOP STONES", "Green Stones"),
    ("CREATED EMEALD", "TOP STONES", "Green Stones"),
    ("EMERALD", "TOP STONES", "Green Stones"),
    ("EMREAL", "TOP STONES", "Green Stones"),
    ("CHROME DIOPSIDE", "TOP STONES", "Green Stones"),
    ("CHROME DIOPOSIDE", "TOP STONES", "Green Stones"),
    ("GREEN AGATE", "TOP STONES", "Green Stones"),
    ("GREEN AGAT", "TOP STONES", "Green Stones"),
    ("PERIDOT", "TOP STONES", "Green Stones"),
    # Other stones combined groups
    ("MYSTIC TOPAZ", "OTHER STONES", "Mystic"),
    ("MYSTIC MB", "OTHER STONES", "Mystic"),
    ("MYST MB", "OTHER STONES", "Mystic"),
    ("MYSTIC QUARTZ", "OTHER STONES", "Mystic"),
    ("SMOKY", "OTHER STONES", "Rauch Topaz"),
    ("SMOKEY", "OTHER STONES", "Rauch Topaz"),
    ("HONEY", "OTHER STONES", "Rauch Topaz"),
    ("RAUCH", "OTHER STONES", "Rauch Topaz"),
    ("RHODOLITE", "OTHER STONES", "Garnet"),
    ("RODOLITE", "OTHER STONES", "Garnet"),
    ("GARNET", "OTHER STONES", "Garnet"),
    ("GRANADA", "OTHER STONES", "Garnet"),
    ("GRANATE", "OTHER STONES", "Garnet"),
    ("PADALITE", "OTHER STONES", "Garnet"),
    ("CUBIC ZIRCONIA", "OTHER STONES", "Color CZ"),
    ("ZIRCONIA", "OTHER STONES", "Color CZ"),
    ("ALL CZ", "OTHER STONES", "Color CZ"),
    ("BLACK ONYX", "OTHER STONES", "Onyx"),
    ("MATT ONYX", "OTHER STONES", "Onyx"),
    ("MATTE ONYX", "OTHER STONES", "Onyx"),
    ("ONYX", "OTHER STONES", "Onyx"),
    ("PICTURE JASPER", "OTHER STONES", "Jasper"),
    ("JASPER", "OTHER STONES", "Jasper"),
]

OTHER_GROUPS: list[tuple[tuple[str, ...], str]] = [
    (("MOTHER OF PEARL", "MOP"), "Mother of Pearl"),
    (("ABALONE", "HELIOTIS"), "Abalone"),
    (("AMETHYST",), "Amethyst"), (("AQUAMARINE",), "Aquamarine"),
    (("APATITE",), "Apatite"), (("AGATE",), "Agate"),
    (("AMBER",), "Amber"), (("AMMOLITE",), "Ammolite"),
    (("BISMUTH",), "Bismuth"), (("DIAMOND",), "Diamond"),
    (("SPINEL",), "Spinel"), (("CARNELIAN",), "Carnelian"),
    (("CHALCEDONY",), "Chalcedony"), (("CHRYSOPRASE",), "Chrysoprase"),
    (("CITRINE",), "Citrine"), (("CORAL",), "Coral"),
    (("CORUNDUM", "CORONDUM"), "Corundum"), (("FLUORITE",), "Fluorite"),
    (("HEMATITE",), "Hematite"), (("HYPERSTHENE",), "Hypersthene"),
    (("IOLITE", "IHOLIT"), "Iolite"), (("JADE",), "Jade"),
    (("KYANITE", "KYN"), "Kyanite"), (("LABRADORITE",), "Labradorite"),
    (("LAPIS", "LAPISE", "LAZURITE"), "Lapis"), (("LARIMAR",), "Larimar"),
    (("MALACHITE",), "Malachite"), (("METEORITE",), "Meteorite"),
    (("MOONSTONE",), "Moonstone"), (("MORGANITE",), "Morganite"),
    (("OPAL",), "Opal"), (("PREHNITE", "PRENITE"), "Prehnite"),
    (("PYRITE",), "Pyrite"), (("QUARTZ",), "Quartz"),
    (("RUBELLITE",), "Rubellite"), (("SULTANITE",), "Sultanite"),
    (("SUNSTONE", "SUN STONE"), "Sunstone"), (("TANZANITE",), "Tanzanite"),
    (("TERAHERTZ",), "Terahertz"), (("TIGER EYE",), "Tiger Eye"),
    (("TOURMALINE",), "Tourmaline"), (("TSAVORITE",), "Tsavorite"),
    (("TURQUOISE",), "Turquoise"), (("HOWLITE",), "Howlite"),
]

STORE_PATTERNS = (
    (re.compile(r"(?:^|[^A-Z0-9])NTR\s*2(?:[^A-Z0-9]|$)", re.I), "NTR2"),
    (re.compile(r"(?:^|[^A-Z0-9])NTR\s*1(?:[^A-Z0-9]|$)", re.I), "NTR1"),
    (re.compile(r"(?:^|[^A-Z0-9])AB(?:[^A-Z0-9]|$)", re.I), "AB"),
    (re.compile(r"(?:^|[^A-Z0-9])SCR(?:[^A-Z0-9]|$)", re.I), "SCR"),
    (re.compile(r"(?:^|[^A-Z0-9])TT(?:[^A-Z0-9]|$)", re.I), "TT"),
    (re.compile(r"(?:^|[^0-9])20(?:[^0-9]|$)", re.I), "20"),
    (re.compile(r"(?:^|[^0-9])63(?:[.\s_-]*[12])?(?:[^0-9]|$)", re.I), "63"),
)


def _clean(text: object) -> str:
    value = str(text or "").upper().replace("Ё", "Е")
    for old, new in (("FRESH WATER", "FRESHWATER"), ("FWP", "FRESHWATER PEARL"),
                     ("LAB-CREATED", "LAB CREATED"), ("/", " "), ("-", " "),
                     (",", " "), (";", " ")):
        value = value.replace(old, new)
    value = re.sub(r"\bNATURAL\b", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _title(value: str) -> str:
    return value.title().replace("Cz", "CZ").replace("Mlbt", "MLBT").replace("Bt", "BT")


def classify(raw: object) -> tuple[str, str, str]:
    text = _clean(raw)
    # Absolute moissanite priority, including common misspellings.
    if re.search(r"MO+I?S+A?N+I?T|MOSSANIT|MUS+ONIT|MUSANIT", text):
        return "TOP STONES", "Moissanite", "moissanite priority"
    # Pearl logic must run before generic rules.
    pearl_hint = "PEARL" in text or "PARL" in text or any(x in text for x in ("AKOYA", "TAHITI", "SOUTH SEA", "MABE"))
    if pearl_hint:
        if "BAROQUE" in text:
            return "PEARLS", "Baroque Pearl", "baroque"
        if any(x in text for x in SEA_WORDS):
            return "PEARLS", "Sea Pearl", "sea pearl"
        if any(x in text for x in COLOR_WORDS):
            return "PEARLS", "Colored Freshwater Pearl", "colored freshwater"
        if "ROUND" in text:
            return "PEARLS", "Round White Freshwater Pearl", "round white freshwater"
        return "PEARLS", "White Freshwater Pearl", "white freshwater"
    # Ordered explicit aliases.
    for alias, segment, category in RULES:
        if alias in text:
            # Ruby in zoisite is not a top ruby.
            if alias == "RUBY" and any(x in text for x in ("ZOISITE", "CIOSITE")):
                continue
            return segment, category, f"alias: {alias}"
    if text == "BT":
        return "TOP STONES", "Other Topaz", "BT"
    if re.search(r"\b(?:CZ|ZIRCON)\b", text):
        return "OTHER STONES", "Color CZ", "CZ/zircon"
    for aliases, category in OTHER_GROUPS:
        if any(alias in text for alias in aliases):
            return "OTHER STONES", category, f"group: {aliases[0]}"
    if "MYST" in text:
        return "OTHER STONES", "Mystic", "mystic"
    # Nothing is discarded: every remaining stone gets its own column.
    return "OTHER STONES", _title(text or "Other"), "own column"


def normalize_product(value: object) -> str:
    text = re.sub(r"\s+", " ", str(value or "").strip()).upper()
    return PRODUCT_ALIASES.get(text, _title(text) if text else "Other")


def extract_period(ws) -> tuple[datetime, datetime] | None:
    months = {
        "ЯНВАРЬ": 1, "ЯНВАРЯ": 1, "ФЕВРАЛЬ": 2, "ФЕВРАЛЯ": 2,
        "МАРТ": 3, "МАРТА": 3, "АПРЕЛЬ": 4, "АПРЕЛЯ": 4,
        "МАЙ": 5, "МАЯ": 5, "ИЮНЬ": 6, "ИЮНЯ": 6,
        "ИЮЛЬ": 7, "ИЮЛЯ": 7, "АВГУСТ": 8, "АВГУСТА": 8,
        "СЕНТЯБРЬ": 9, "СЕНТЯБРЯ": 9, "ОКТЯБРЬ": 10, "ОКТЯБРЯ": 10,
        "НОЯБРЬ": 11, "НОЯБРЯ": 11, "ДЕКАБРЬ": 12, "ДЕКАБРЯ": 12,
    }
    for row in ws.iter_rows(min_row=1, max_row=min(15, ws.max_row), min_col=1, max_col=min(8, ws.max_column)):
        for cell in row:
            text = str(cell.value or "").strip()
            if not text:
                continue
            match = re.search(r"(\d{2}\.\d{2}\.\d{4})\s*[-–—]\s*(\d{2}\.\d{2}\.\d{4})", text)
            if match:
                return datetime.strptime(match.group(1), "%d.%m.%Y"), datetime.strptime(match.group(2), "%d.%m.%Y")
            upper = text.upper().replace("Ё", "Е")
            match = re.search(r"([А-Я]+)\s+(20\d{2})\s*Г?\.?\s*[-–—]\s*([А-Я]+)\s+(20\d{2})", upper)
            if match and match.group(1) in months and match.group(3) in months:
                sm, sy = months[match.group(1)], int(match.group(2))
                em, ey = months[match.group(3)], int(match.group(4))
                return datetime(sy, sm, 1), datetime(ey, em, calendar.monthrange(ey, em)[1])
    return None


def period_text(period: tuple[datetime, datetime] | None) -> str:
    if not period:
        return "Период не найден"
    return f"{period[0]:%d.%m.%Y} — {period[1]:%d.%m.%Y}"


def _is_consolidated(ws) -> bool:
    header = " ".join(str(ws.cell(r, 1).value or "") for r in range(1, min(8, ws.max_row) + 1)).upper()
    return "МАГАЗИН" in header and "КАМЕНЬ" in header and "НОМЕНКЛАТУР" in header


def _normalize_store(value: object) -> str | None:
    text = re.sub(r"[^A-ZА-Я0-9]", "", str(value or "").upper())
    if "63NDC" in text or text.startswith("63RETAIL") or text.startswith("63TIM"):
        return "63"
    if text.startswith("NTR2"):
        return "NTR2"
    if text.startswith("NTR1"):
        return "NTR1"
    if text.startswith("AB"):
        return "AB"
    if text.startswith("SCR"):
        return "SCR"
    if text == "TT" or text.startswith("STOCKTT") or text.startswith("ALLSALESTT"):
        return "TT"
    if text.startswith("20"):
        return "20"
    if text.startswith("63"):
        return "63"
    return None


def detect_store(path: Path) -> str:
    path = Path(path)
    if path.suffix.lower() == ".xlsx":
        try:
            wb = load_workbook(path, read_only=False, data_only=True)
            ws = wb.active
            consolidated = _is_consolidated(ws)
            wb.close()
            if consolidated:
                return "Все магазины"
        except Exception:
            pass
    compact = re.sub(r"[^A-Z0-9]", "", path.stem.upper())
    if compact.startswith("NTR2"):
        return "NTR2"
    if compact.startswith("NTR1") or compact == "NTR":
        return "NTR1"
    if compact.startswith("AB"):
        return "AB"
    if compact.startswith("SCR"):
        return "SCR"
    if compact.startswith("TT") or "ALLSALESTT" in compact:
        return "TT"
    if compact.startswith("20"):
        return "20"
    if compact.startswith("63") or "MUINE" in compact:
        return "63"
    for pattern, store in STORE_PATTERNS:
        if pattern.search(path.stem):
            return store
    raise ValueError(f"Не удалось определить магазин по имени файла: {path.name}")


@dataclass
class ReportUnit:
    store: str
    period: tuple[datetime, datetime] | None
    source: str
    records: dict[tuple[str, str, str], dict[str, float]] = field(default_factory=lambda: defaultdict(lambda: {"qty": 0.0, "amount": 0.0}))
    raw_rules: dict[str, tuple[str, str, str]] = field(default_factory=dict)

    def add(self, raw: object, product: object, qty: object, amount: object) -> None:
        try:
            q = float(qty or 0)
        except Exception:
            q = 0.0
        try:
            a = float(amount or 0)
        except Exception:
            a = 0.0
        if q == 0 and a == 0:
            return
        segment, stone, rule = classify(raw)
        product_name = normalize_product(product)
        row = self.records[(segment, stone, product_name)]
        row["qty"] += q
        row["amount"] += a
        self.raw_rules[str(raw or "").strip()] = (segment, stone, rule)

    @property
    def total_qty(self) -> float:
        return sum(v["qty"] for v in self.records.values())

    @property
    def total_amount(self) -> float:
        return sum(v["amount"] for v in self.records.values())

    def segment_total(self, segment: str) -> tuple[float, float]:
        qty = amount = 0.0
        for (seg, _, _), values in self.records.items():
            if seg == segment:
                qty += values["qty"]
                amount += values["amount"]
        return qty, amount


def parse_consolidated(path: Path) -> list[ReportUnit]:
    wb = load_workbook(path, data_only=True, read_only=False)
    ws = wb.active
    period = extract_period(ws)
    units: dict[str, ReportUnit] = {}
    current_store: str | None = None
    current_stone: str | None = None
    for row in range(6, ws.max_row + 1):
        value = ws.cell(row, 1).value
        if value is None:
            continue
        text = str(value).strip()
        if not text:
            continue
        indent = float(ws.cell(row, 1).alignment.indent or 0)
        outline = int(ws.row_dimensions[row].outlineLevel or 0)
        level = outline if outline else (0 if indent < 1 else 1 if indent < 3 else 2)
        if level == 0:
            store = _normalize_store(text)
            current_store = store
            current_stone = None
            if store and store not in units:
                units[store] = ReportUnit(store=store, period=period, source=path.name)
            continue
        if not current_store:
            continue
        if level == 1:
            current_stone = text
            continue
        if level >= 2 and current_stone:
            # Sold columns only. Return columns 10-11 are deliberately ignored.
            units[current_store].add(current_stone, text, ws.cell(row, 8).value, ws.cell(row, 9).value)
    wb.close()
    return [unit for unit in units.values() if unit.records]


def parse_legacy(path: Path) -> ReportUnit:
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Поддерживается формат .xlsx. Сохраните старый .xls как .xlsx")
    wb = load_workbook(path, data_only=True, read_only=False)
    ws = wb.active
    store = detect_store(path)
    period = extract_period(ws)
    unit = ReportUnit(store=store, period=period, source=path.name)
    current_stone: str | None = None
    current_group: str | None = None
    # Two common legacy layouts are supported:
    # A) stone at indent 2, product at indent 4, qty/sales in H/I;
    # B) product group heading in A, stone rows underneath, qty/sales in H/I.
    for row in range(6, ws.max_row + 1):
        value = ws.cell(row, 1).value
        if value is None:
            continue
        text = str(value).strip()
        if not text:
            continue
        upper = text.upper()
        indent = float(ws.cell(row, 1).alignment.indent or 0)
        outline = int(ws.row_dimensions[row].outlineLevel or 0)
        level = outline if outline else (0 if indent < 1 else 1 if indent < 3 else 2)
        if upper.startswith("ИТОГО") or upper.startswith("TOTAL") or "VLADIMIR PANASIAN" in upper:
            continue
        if upper in PRODUCT_ALIASES:
            if level >= 2 and current_stone:
                unit.add(current_stone, text, ws.cell(row, 8).value, ws.cell(row, 9).value)
            else:
                current_group = text
            continue
        if level == 1:
            current_stone = text
            continue
        qty = ws.cell(row, 8).value
        amount = ws.cell(row, 9).value
        if (qty not in (None, 0) or amount not in (None, 0)) and current_group:
            unit.add(text, current_group, qty, amount)
    wb.close()
    if not unit.records:
        raise ValueError("В файле не найдены строки продаж в поддерживаемом формате")
    return unit


def parse_file(path: Path) -> list[ReportUnit]:
    path = Path(path)
    wb = load_workbook(path, data_only=True, read_only=False)
    consolidated = _is_consolidated(wb.active)
    wb.close()
    return parse_consolidated(path) if consolidated else [parse_legacy(path)]


def _sheet_name(base: str, used: set[str]) -> str:
    clean = re.sub(r"[\\/*?:\[\]]", " ", base).strip()[:31] or "Report"
    name = clean
    counter = 2
    while name in used:
        suffix = f" ({counter})"
        name = clean[:31 - len(suffix)] + suffix
        counter += 1
    used.add(name)
    return name


def _period_short(period: tuple[datetime, datetime] | None) -> str:
    if not period:
        return "no period"
    if period[0].year == period[1].year and period[0].month == period[1].month:
        return period[0].strftime("%m.%Y")
    return f"{period[0]:%m.%y}-{period[1]:%m.%y}"


def _ordered_stones(unit: ReportUnit) -> list[tuple[str, str]]:
    present = {(seg, stone) for seg, stone, _ in unit.records}
    ordered: list[tuple[str, str]] = []
    for stone in TOP_ORDER:
        if ("TOP STONES", stone) in present:
            ordered.append(("TOP STONES", stone))
    for stone in PEARL_ORDER:
        if ("PEARLS", stone) in present:
            ordered.append(("PEARLS", stone))
    ordered.extend(sorted((seg, stone) for seg, stone in present if seg == "OTHER STONES"))
    return ordered


def _totals(unit: ReportUnit, segment: str | None = None, stone: str | None = None, product: str | None = None) -> tuple[float, float]:
    qty = amount = 0.0
    for (seg, st, prod), values in unit.records.items():
        if segment and seg != segment:
            continue
        if stone and st != stone:
            continue
        if product and prod != product:
            continue
        qty += values["qty"]
        amount += values["amount"]
    return qty, amount


def build_report(units: list[ReportUnit], output: Path) -> Path:
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "SUMMARY"
    used = {"SUMMARY"}

    thin = Side(style="thin", color="404040")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fills = {
        "title": PatternFill("solid", fgColor="111111"),
        "TOP STONES": PatternFill("solid", fgColor="6F42C1"),
        "PEARLS": PatternFill("solid", fgColor="B58A24"),
        "OTHER STONES": PatternFill("solid", fgColor="3D7A44"),
        "total": PatternFill("solid", fgColor="E2F0D9"),
        "compare_up": PatternFill("solid", fgColor="D9EAD3"),
        "compare_down": PatternFill("solid", fgColor="F4CCCC"),
    }
    white = Font(color="FFFFFF", bold=True)
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def style(cell, *, fill=None, font=None, align=center, number_format=None):
        cell.border = border
        cell.alignment = align
        if fill:
            cell.fill = fill
        if font:
            cell.font = font
        if number_format:
            cell.number_format = number_format

    headers = [
        "Store", "Period", "Source file", "PCS", "Sales",
        "Top Stones PCS %", "Top Stones Sales %", "Pearls PCS %",
        "Pearls Sales %", "Other Stones PCS %", "Other Stones Sales %",
    ]
    ws.append(headers)
    for cell in ws[1]:
        style(cell, fill=fills["title"], font=white)
    summary_rows: dict[int, ReportUnit] = {}
    for unit in units:
        row = [unit.store, period_text(unit.period), unit.source, unit.total_qty, unit.total_amount]
        for segment in SEGMENTS:
            qty, amount = unit.segment_total(segment)
            row.extend([qty / unit.total_qty if unit.total_qty else 0, amount / unit.total_amount if unit.total_amount else 0])
        ws.append(row)
        summary_rows[ws.max_row] = unit
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            style(cell, align=left if cell.column <= 3 else center)
        row[3].number_format = "#,##0.##"
        row[4].number_format = "#,##0"
        for cell in row[5:]:
            cell.number_format = "0.00%"
    widths = [12, 24, 34, 12, 16, 18, 18, 16, 16, 20, 20]
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

    if units:
        chart = BarChart()
        chart.title = "Sales by store and period"
        chart.y_axis.title = "Sales"
        chart.x_axis.title = "Store / period"
        data = Reference(ws, min_col=5, min_row=1, max_row=1 + len(units))
        cats = Reference(ws, min_col=1, min_row=2, max_row=1 + len(units))
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.height = 8
        chart.width = 15
        ws.add_chart(chart, "M2")

        qty_chart = BarChart()
        qty_chart.title = "PCS by store and period"
        qty_chart.y_axis.title = "PCS"
        qty_chart.add_data(Reference(ws, min_col=4, min_row=1, max_row=1 + len(units)), titles_from_data=True)
        qty_chart.set_categories(cats)
        qty_chart.height = 8
        qty_chart.width = 15
        ws.add_chart(qty_chart, "M18")

    for unit in units:
        title = f"{unit.store} {_period_short(unit.period)}"
        sheet = wb.create_sheet(_sheet_name(title, used))
        stones = _ordered_stones(unit)
        max_col = max(5, 1 + 2 * len(stones))
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
        sheet.cell(1, 1).value = f"{unit.store} — Stone Sales Report"
        style(sheet.cell(1, 1), fill=fills["title"], font=Font(color="FFFFFF", bold=True, size=16))
        sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
        sheet.cell(2, 1).value = f"Report period: {period_text(unit.period)}"
        style(sheet.cell(2, 1), fill=fills["total"], font=bold, align=left)
        sheet.merge_cells(start_row=3, start_column=1, end_row=3, end_column=max_col)
        sheet.cell(3, 1).value = f"Source file: {unit.source}"
        style(sheet.cell(3, 1), align=left)
        sheet.cell(4, 1).value = "Product / Total"
        style(sheet.cell(4, 1), fill=fills["title"], font=white)

        col = 2
        index = 0
        while index < len(stones):
            segment = stones[index][0]
            start_col = col
            while index < len(stones) and stones[index][0] == segment:
                index += 1
                col += 2
            end_col = col - 1
            sheet.merge_cells(start_row=4, start_column=start_col, end_row=4, end_column=end_col)
            sheet.cell(4, start_col).value = segment
            for current in range(start_col, end_col + 1):
                style(sheet.cell(4, current), fill=fills[segment], font=white)

        col = 2
        for segment, stone in stones:
            sheet.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)
            sheet.cell(5, col).value = stone
            for current in (col, col + 1):
                style(sheet.cell(5, current), fill=fills[segment], font=white)
            sheet.cell(6, col).value = "PCS"
            sheet.cell(6, col + 1).value = "Sales"
            for current in (col, col + 1):
                style(sheet.cell(6, current), fill=fills[segment], font=white)
            col += 2

        row = 7
        products = list(PRODUCT_ORDER)
        extra_products = sorted({prod for _, _, prod in unit.records if prod not in products})
        products.extend(extra_products)
        for product in products:
            if not any(_totals(unit, seg, stone, product) != (0, 0) for seg, stone in stones):
                continue
            sheet.cell(row, 1).value = product
            style(sheet.cell(row, 1), font=bold, align=left)
            col = 2
            for segment, stone in stones:
                qty, amount = _totals(unit, segment, stone, product)
                sheet.cell(row, col).value = qty if qty else None
                sheet.cell(row, col + 1).value = amount if amount else None
                style(sheet.cell(row, col), number_format="#,##0.##")
                style(sheet.cell(row, col + 1), number_format="#,##0")
                col += 2
            row += 1

        for label, kind in (("TOTAL PCS", "qty"), ("TOTAL SALES", "amount"), ("% OF STORE PCS", "qty_pct"), ("% OF STORE SALES", "amount_pct")):
            sheet.cell(row, 1).value = label
            style(sheet.cell(row, 1), fill=fills["total"], font=bold, align=left)
            col = 2
            for segment, stone in stones:
                qty, amount = _totals(unit, segment, stone)
                values = {
                    "qty": (qty, None), "amount": (None, amount),
                    "qty_pct": (qty / unit.total_qty if unit.total_qty else 0, None),
                    "amount_pct": (None, amount / unit.total_amount if unit.total_amount else 0),
                }[kind]
                for offset, value in enumerate(values):
                    cell = sheet.cell(row, col + offset)
                    cell.value = value
                    fmt = "0.00%" if "pct" in kind else "#,##0.##" if kind == "qty" else "#,##0"
                    style(cell, fill=fills["total"], font=bold, number_format=fmt)
                col += 2
            row += 1

        analysis_start = row + 2
        sheet.cell(analysis_start, 1).value = "Segment analysis"
        sheet.cell(analysis_start, 1).font = Font(bold=True, size=13)
        headers2 = ("Segment", "PCS", "Sales", "PCS %", "Sales %")
        for c, text in enumerate(headers2, 1):
            sheet.cell(analysis_start + 1, c).value = text
            style(sheet.cell(analysis_start + 1, c), fill=fills["title"], font=white)
        for offset, segment in enumerate(SEGMENTS, 2):
            qty, amount = unit.segment_total(segment)
            values = (segment, qty, amount, qty / unit.total_qty if unit.total_qty else 0, amount / unit.total_amount if unit.total_amount else 0)
            for c, value in enumerate(values, 1):
                sheet.cell(analysis_start + offset, c).value = value
                style(sheet.cell(analysis_start + offset, c), fill=fills[segment] if c == 1 else None, font=white if c == 1 else None, align=left if c == 1 else center)
            sheet.cell(analysis_start + offset, 2).number_format = "#,##0.##"
            sheet.cell(analysis_start + offset, 3).number_format = "#,##0"
            sheet.cell(analysis_start + offset, 4).number_format = "0.00%"
            sheet.cell(analysis_start + offset, 5).number_format = "0.00%"

        pie = PieChart()
        pie.title = "Sales structure"
        pie.add_data(Reference(sheet, min_col=3, min_row=analysis_start + 1, max_row=analysis_start + 4), titles_from_data=True)
        pie.set_categories(Reference(sheet, min_col=1, min_row=analysis_start + 2, max_row=analysis_start + 4))
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        pie.height = 8
        pie.width = 11
        sheet.add_chart(pie, "G" + str(analysis_start))

        sheet.freeze_panes = "B7"
        sheet.sheet_view.showGridLines = False
        sheet.column_dimensions["A"].width = 24
        for current in range(2, max_col + 1):
            sheet.column_dimensions[get_column_letter(current)].width = 13 if current % 2 == 0 else 16

    # Compare periods for stores that occur more than once.
    by_store: dict[str, list[ReportUnit]] = defaultdict(list)
    for unit in units:
        by_store[unit.store].append(unit)
    for store, store_units in by_store.items():
        if len(store_units) < 2:
            continue
        store_units.sort(key=lambda item: item.period[0] if item.period else datetime.min)
        sheet = wb.create_sheet(_sheet_name(f"COMPARE {store}", used))
        compare_headers = ["Period", "Source", "PCS", "Sales", "PCS Δ", "PCS Δ %", "Sales Δ", "Sales Δ %"]
        sheet.append(compare_headers)
        for cell in sheet[1]:
            style(cell, fill=fills["title"], font=white)
        previous: ReportUnit | None = None
        for unit in store_units:
            if previous:
                qty_delta = unit.total_qty - previous.total_qty
                amount_delta = unit.total_amount - previous.total_amount
                qty_pct = qty_delta / previous.total_qty if previous.total_qty else 0
                amount_pct = amount_delta / previous.total_amount if previous.total_amount else 0
            else:
                qty_delta = amount_delta = qty_pct = amount_pct = 0
            sheet.append([period_text(unit.period), unit.source, unit.total_qty, unit.total_amount, qty_delta, qty_pct, amount_delta, amount_pct])
            previous = unit
        for row_cells in sheet.iter_rows(min_row=2):
            for cell in row_cells:
                style(cell, align=left if cell.column <= 2 else center)
            for col_index in (3, 5):
                row_cells[col_index - 1].number_format = "#,##0.##"
            for col_index in (4, 7):
                row_cells[col_index - 1].number_format = "#,##0"
            for col_index in (6, 8):
                row_cells[col_index - 1].number_format = "0.00%"
            for col_index in (5, 6, 7, 8):
                value = row_cells[col_index - 1].value or 0
                row_cells[col_index - 1].fill = fills["compare_up"] if value >= 0 else fills["compare_down"]
        for index, width in enumerate((24, 32, 14, 18, 14, 14, 18, 14), 1):
            sheet.column_dimensions[get_column_letter(index)].width = width
        chart = BarChart()
        chart.title = f"{store}: Sales by period"
        chart.add_data(Reference(sheet, min_col=4, min_row=1, max_row=1 + len(store_units)), titles_from_data=True)
        chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=1 + len(store_units)))
        chart.height = 8
        chart.width = 15
        sheet.add_chart(chart, "J2")
        qty_chart = BarChart()
        qty_chart.title = f"{store}: PCS by period"
        qty_chart.add_data(Reference(sheet, min_col=3, min_row=1, max_row=1 + len(store_units)), titles_from_data=True)
        qty_chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=1 + len(store_units)))
        qty_chart.height = 8
        qty_chart.width = 15
        sheet.add_chart(qty_chart, "J18")
        sheet.freeze_panes = "A2"
        sheet.sheet_view.showGridLines = False

    rules = wb.create_sheet("RULES")
    rules.append(["RAW / alias", "Segment", "Final column", "Priority / note"])
    for cell in rules[1]:
        style(cell, fill=fills["title"], font=white)
    seen: set[tuple[str, str, str]] = set()
    for alias, segment, category in RULES:
        key = (alias, segment, category)
        if key not in seen:
            rules.append([alias, segment, category, "explicit alias"])
            seen.add(key)
    for aliases, category in OTHER_GROUPS:
        for alias in aliases:
            key = (alias, "OTHER STONES", category)
            if key not in seen:
                rules.append([alias, "OTHER STONES", category, "group alias"])
                seen.add(key)
    raw_start = rules.max_row + 3
    rules.cell(raw_start, 1).value = "Names actually found in uploaded files"
    rules.cell(raw_start, 1).font = Font(bold=True, size=13)
    rules.append(["RAW name", "Segment", "Final column", "Applied rule"])
    for cell in rules[rules.max_row]:
        style(cell, fill=fills["title"], font=white)
    actual: dict[str, tuple[str, str, str]] = {}
    for unit in units:
        actual.update(unit.raw_rules)
    for raw, (segment, category, rule) in sorted(actual.items(), key=lambda item: item[0].upper()):
        rules.append([raw, segment, category, rule])
    for row in rules.iter_rows(min_row=2):
        for cell in row:
            style(cell, align=left)
    for index, width in enumerate((38, 20, 34, 30), 1):
        rules.column_dimensions[get_column_letter(index)].width = width
    rules.freeze_panes = "A2"
    rules.sheet_view.showGridLines = False

    wb.save(output)
    return output


def run_files(files: list[Path], output_file: Path) -> Path:
    paths = [Path(path) for path in files if Path(path).suffix.lower() in {".xlsx", ".xls"} and not Path(path).name.startswith("~$")]
    if not paths:
        raise RuntimeError("Не выбраны Excel-файлы")
    units: list[ReportUnit] = []
    errors: list[str] = []
    for path in paths:
        try:
            units.extend(parse_file(path))
        except Exception as exc:
            errors.append(f"{path.name}: {exc}")
    if errors:
        raise RuntimeError("Ошибки чтения файлов:\n" + "\n".join(errors))
    if not units:
        raise RuntimeError("В выбранных файлах не найдены продажи")
    # Merge only the special 63.1 + 63.2 case when period is identical.
    merged: list[ReportUnit] = []
    groups: dict[tuple[str, object], list[ReportUnit]] = defaultdict(list)
    for unit in units:
        if unit.store == "63" and re.search(r"63[.\s_-]*[12]", unit.source, re.I):
            groups[(unit.store, unit.period)].append(unit)
        else:
            merged.append(unit)
    for (_, period), group in groups.items():
        if len(group) == 1:
            merged.extend(group)
            continue
        combined = ReportUnit("63", period, " + ".join(item.source for item in group))
        for item in group:
            for (segment, stone, product), values in item.records.items():
                row = combined.records[(segment, stone, product)]
                row["qty"] += values["qty"]
                row["amount"] += values["amount"]
            combined.raw_rules.update(item.raw_rules)
        merged.append(combined)
    merged.sort(key=lambda item: (item.store, item.period[0] if item.period else datetime.min, item.source))
    return build_report(merged, Path(output_file))
