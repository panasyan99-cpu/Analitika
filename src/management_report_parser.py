from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import BinaryIO, Iterable, Optional
import calendar
import re
import xml.etree.ElementTree as ET
import zipfile

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

KNOWN_CATEGORIES = {
    "bracelet", "bracelets", "chain", "earrings", "necklace", "pendant", "ring",
    "pearl on chain", "pearl necklaces/bracelets", "jewelry service", "pearls",
    "souvenirs", "brooch", "stone", "stones&pearls", "other", "other metal",
}


@dataclass(frozen=True)
class Metrics:
    quantity: float = 0.0
    discount_pct: float = 0.0
    average_price: float = 0.0
    revenue: float = 0.0
    return_quantity: float = 0.0
    return_amount: float = 0.0

    @property
    def net_revenue(self) -> float:
        return self.revenue - self.return_amount


@dataclass(frozen=True)
class ProductFact:
    row_number: int
    store: str
    manager: str
    top_group: str
    product_section: str
    category: str
    sku: str
    stone: str
    assay: str
    note: str
    metrics: Metrics


@dataclass(frozen=True)
class ReportMeta:
    source_file: str
    title: str
    period_label: str
    period_start: Optional[str]
    period_end: Optional[str]
    period_days: int
    generated_at: Optional[str]
    generated_by: str
    grouping_label: str


@dataclass
class ParsedReport:
    meta: ReportMeta
    totals: Metrics
    facts: list[ProductFact]
    stores: dict[str, Metrics]
    validation: dict[str, float | str | bool]

    def to_dict(self) -> dict:
        return {
            "meta": asdict(self.meta),
            "totals": asdict(self.totals),
            "facts": [{**asdict(fact), "metrics": asdict(fact.metrics)} for fact in self.facts],
            "stores": {key: asdict(value) for key, value in self.stores.items()},
            "validation": self.validation,
        }


@dataclass
class _Row:
    number: int
    level: int
    height: float
    values: dict[str, str]


def _float(value: str | None) -> float:
    if value in (None, ""):
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return [
        "".join(node.text or "" for node in item.findall(".//m:t", NS))
        for item in root.findall("m:si", NS)
    ]


def _iter_rows(source: str | Path | bytes | bytearray | BinaryIO) -> Iterable[_Row]:
    if isinstance(source, (bytes, bytearray)):
        archive_source: str | Path | BinaryIO = BytesIO(source)
    else:
        archive_source = source
    with zipfile.ZipFile(archive_source) as zf:
        shared = _load_shared_strings(zf)
        sheet_name = "xl/worksheets/sheet1.xml"
        if sheet_name not in zf.namelist():
            raise ValueError("В Excel не найден первый рабочий лист.")
        with zf.open(sheet_name) as stream:
            for _, row in ET.iterparse(stream, events=("end",)):
                if not row.tag.endswith("}row"):
                    continue
                values: dict[str, str] = {}
                for cell in row.findall("m:c", NS):
                    reference = cell.attrib.get("r", "")
                    match = re.match(r"([A-Z]+)", reference)
                    if not match:
                        continue
                    column = match.group(1)
                    cell_type = cell.attrib.get("t")
                    value_node = cell.find("m:v", NS)
                    value = ""
                    if value_node is not None and value_node.text is not None:
                        value = value_node.text
                        if cell_type == "s":
                            try:
                                value = shared[int(value)]
                            except (ValueError, IndexError):
                                pass
                    inline = cell.find("m:is", NS)
                    if inline is not None:
                        value = "".join(node.text or "" for node in inline.findall(".//m:t", NS))
                    values[column] = value
                yield _Row(
                    number=int(row.attrib.get("r", "0") or 0),
                    level=int(row.attrib.get("outlineLevel", "0") or 0),
                    height=float(row.attrib.get("ht", "0") or 0),
                    values=values,
                )
                row.clear()


def _metrics(row: _Row) -> Metrics:
    return Metrics(
        quantity=_float(row.values.get("H") or row.values.get("D")),
        discount_pct=_float(row.values.get("E")),
        average_price=_float(row.values.get("F")),
        revenue=_float(row.values.get("I") or row.values.get("G")),
        return_quantity=_float(row.values.get("J")),
        return_amount=_float(row.values.get("K")),
    )


def _period_from_title(title: str) -> tuple[Optional[date], Optional[date], str]:
    exact = re.search(
        r"(\d{2})\.(\d{2})\.(20\d{2})\s*[-–—]\s*(\d{2})\.(\d{2})\.(20\d{2})",
        title,
    )
    if exact:
        start = date(int(exact.group(3)), int(exact.group(2)), int(exact.group(1)))
        end = date(int(exact.group(6)), int(exact.group(5)), int(exact.group(4)))
        month_labels = (
            "", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
        )
        is_full_month = (
            start.year == end.year
            and start.month == end.month
            and start.day == 1
            and end.day == calendar.monthrange(end.year, end.month)[1]
        )
        label = f"{month_labels[start.month]} {start.year}" if is_full_month else f"{start:%d.%m.%Y}–{end:%d.%m.%Y}"
        return start, end, label

    months = {
        "январь": 1, "февраль": 2, "март": 3, "апрель": 4, "май": 5,
        "июнь": 6, "июль": 7, "август": 8, "сентябрь": 9,
        "октябрь": 10, "ноябрь": 11, "декабрь": 12,
    }
    lowered = title.casefold()
    year_match = re.search(r"(20\d{2})", title)
    year = int(year_match.group(1)) if year_match else None
    if year:
        for label, month in months.items():
            if label not in lowered:
                continue
            start = date(year, month, 1)
            next_month = date(year + 1, 1, 1) if month == 12 else date(year, month + 1, 1)
            end = next_month - timedelta(days=1)
            return start, end, f"{label.capitalize()} {year}"
    return None, None, title.strip()


def _looks_like_sku(value: str) -> bool:
    value = value.strip()
    return bool(re.search(r"\d", value)) or ("-" in value and " " not in value)


def _normalize_category(value: str) -> str:
    text = " ".join(str(value or "").strip().split())
    aliases = {
        "bracelets": "Bracelet",
        "pearl necklace": "Pearl necklace",
        "pearl bracelet": "Pearl bracelet",
        "pearl on chain": "Pearl on chain",
        "jewelry service": "Service",
        "service": "Service",
    }
    return aliases.get(text.casefold(), text)


def _choose_category(ancestors: list[str], sku: str) -> tuple[str, str]:
    cleaned = [" ".join(item.strip().split()) for item in ancestors if item and item.strip()]
    if "service" in sku.strip().casefold():
        return "Service", cleaned[0] if cleaned else ""
    for value in reversed(cleaned):
        if value.casefold() in KNOWN_CATEGORIES:
            return _normalize_category(value), cleaned[0] if cleaned else ""
    if cleaned:
        return _normalize_category(cleaned[-1]), cleaned[0]
    if _looks_like_sku(sku):
        return "Без номенклатурной группы", ""
    return _normalize_category(sku) or "Без номенклатурной группы", ""


def parse_report(
    source: str | Path | bytes | bytearray | BinaryIO,
    *,
    source_name: str | None = None,
) -> ParsedReport:
    rows = list(_iter_rows(source))
    if len(rows) < 8:
        raise ValueError("Файл не похож на поддерживаемую выгрузку продаж 1С.")

    title = rows[0].values.get("A", "")
    grouping = next((row.values.get("A", "") for row in rows[:8] if ";" in row.values.get("A", "")), "")
    required_groupings = ("Магазин", "Менеджер", "Товар", "Камень/вставка", "Проба", "Номенклатурная группа")
    if not grouping or not all(item.casefold() in grouping.casefold() for item in required_groupings):
        raise ValueError(
            "Неверная структура отчета. Нужны группировки: Магазин → Менеджер → Товар → "
            "Камень/вставка → Проба → Номенклатурная группа."
        )

    period_start, period_end, period_label = _period_from_title(title)
    if period_start is None or period_end is None:
        raise ValueError("Не удалось определить период из заголовка выгрузки.")
    period_days = (period_end - period_start).days + 1

    generated_raw = rows[-1].values.get("A", "") if rows else ""
    generated_at = None
    generated_by = ""
    generated_match = re.match(r"(\d{2}\.\d{2}\.\d{4}\s+\d{1,2}:\d{2}:\d{2})\s*(.*)", generated_raw)
    if generated_match:
        try:
            generated_at = datetime.strptime(generated_match.group(1), "%d.%m.%Y %H:%M:%S").isoformat()
        except ValueError:
            generated_at = generated_match.group(1)
        generated_by = generated_match.group(2).strip()

    total_row = next((row for row in reversed(rows) if row.values.get("A", "").strip().startswith("Итого")), None)
    if total_row is None:
        raise ValueError("В выгрузке не найдена итоговая строка «Итого».")
    totals = _metrics(total_row)

    stores: dict[str, Metrics] = {}
    for row in rows[6:]:
        label = row.values.get("A", "").strip()
        if label.startswith("Итого"):
            break
        if label and row.level == 0:
            stores[label] = _metrics(row)

    stack: dict[int, str] = {}
    facts: list[ProductFact] = []
    index = 6
    while index < len(rows):
        row = rows[index]
        label = row.values.get("A", "").strip()
        if label.startswith("Итого"):
            break
        for level in list(stack):
            if level >= row.level:
                stack.pop(level, None)

        row_metrics = _metrics(row)
        is_product = (
            row.height >= 50
            and row.level >= 2
            and bool(label or row_metrics.revenue or row_metrics.quantity or row_metrics.return_amount)
        )
        if is_product:
            category, section = _choose_category(
                [stack[level] for level in sorted(stack) if 2 < level < row.level],
                label,
            )
            children: list[_Row] = []
            child_index = index + 1
            while child_index < len(rows) and rows[child_index].level > row.level:
                children.append(rows[child_index])
                child_index += 1
            stone = ""
            assay = ""
            if children:
                stone_row = next((child for child in children if child.level == row.level + 1), None)
                assay_row = next((child for child in children if child.level == row.level + 2), None)
                stone = stone_row.values.get("A", "").strip() if stone_row else ""
                assay = assay_row.values.get("A", "").strip() if assay_row else ""
            sku_label = label or f"(без артикула, строка {row.number})"
            facts.append(ProductFact(
                row_number=row.number,
                store=stack.get(0, ""),
                manager=stack.get(1, ""),
                top_group=stack.get(2, ""),
                product_section=section,
                category=category,
                sku=sku_label,
                stone=stone,
                assay=assay,
                note=row.values.get("L", "").strip(),
                metrics=row_metrics,
            ))
            stack[row.level] = sku_label
            index += 1
            continue
        if label:
            stack[row.level] = label
        index += 1

    product_revenue = sum(fact.metrics.revenue for fact in facts)
    product_returns = sum(fact.metrics.return_amount for fact in facts)
    store_revenue = sum(metric.revenue for metric in stores.values())
    store_quantity = sum(metric.quantity for metric in stores.values())
    validation = {
        "facts_count": len(facts),
        "product_revenue_matches_total": abs(product_revenue - totals.revenue) < 0.5,
        "product_revenue_difference": product_revenue - totals.revenue,
        "product_return_difference": product_returns - totals.return_amount,
        "store_revenue_difference": store_revenue - totals.revenue,
        "store_quantity_difference": store_quantity - totals.quantity,
        "quantity_note": (
            "Верхний KPI количества использует явную строку «Итого»; разрезы строятся по строкам товаров."
        ),
    }

    if source_name:
        file_name = source_name
    elif isinstance(source, (str, Path)):
        file_name = Path(source).name
    else:
        file_name = "report.xlsx"

    return ParsedReport(
        meta=ReportMeta(
            source_file=file_name,
            title=title,
            period_label=period_label,
            period_start=period_start.isoformat() if period_start else None,
            period_end=period_end.isoformat() if period_end else None,
            period_days=period_days,
            generated_at=generated_at,
            generated_by=generated_by,
            grouping_label=grouping,
        ),
        totals=totals,
        facts=facts,
        stores=stores,
        validation=validation,
    )
