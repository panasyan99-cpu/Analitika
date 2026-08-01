from __future__ import annotations

from collections import defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import BinaryIO, Mapping
import math
import re

import pandas as pd

from src.management_report_parser import Metrics, ReportMeta, _iter_rows, _metrics, _period_from_title
from src.store_normalization import analytics_store_name


SALES = "sales"
CONSULTANTS = "consultants"
SUPPLIERS = "suppliers"
UNKNOWN = "Не определен"

KIND_LABELS = {
    SALES: "Продажи по магазинам",
    CONSULTANTS: "Продажи по консультантам",
    SUPPLIERS: "Продажи по поставщикам",
}

EXPECTED_GROUPINGS = {
    SALES: ("Магазин", "Камень/вставка", "Проба", "Номенклатурная группа"),
    CONSULTANTS: ("Менеджер", "Проба", "Номенклатурная группа"),
    SUPPLIERS: ("Номенклатурная группа", "Поставщик"),
}

PRIMARY_DIMENSION = {
    SALES: "stores",
    CONSULTANTS: "consultants",
    SUPPLIERS: "suppliers",
}


@dataclass(frozen=True)
class HierarchyFact:
    row_number: int
    level: int
    parent: str
    label: str
    metrics: Metrics


@dataclass
class BlockReport:
    kind: str
    meta: ReportMeta
    totals: Metrics
    dimensions: dict[str, dict[str, Metrics]]
    facts: list[HierarchyFact]
    validation: dict[str, float | str | bool]

    def to_dict(self) -> dict:
        return {
            "kind": self.kind,
            "meta": asdict(self.meta),
            "totals": asdict(self.totals),
            "dimensions": {
                dimension: {name: asdict(metrics) for name, metrics in rows.items()}
                for dimension, rows in self.dimensions.items()
            },
            "facts": [
                {**asdict(fact), "metrics": asdict(fact.metrics)}
                for fact in self.facts
            ],
            "validation": dict(self.validation),
        }


def _clean(value: object, fallback: str = "Не указано") -> str:
    text = " ".join(str(value or "").strip().split())
    return text or fallback


def _normalize_supplier(value: object) -> str:
    text = _clean(value, UNKNOWN)
    folded = text.casefold().replace("ё", "е")
    if folded.startswith("own production") or folded.startswith("ownproduction"):
        return "Own production"
    if folded in {"не определен", "не определён", "unknown", "(не указано)", "не указано"}:
        return UNKNOWN
    return text


def _add_metric(target: dict[str, dict[str, float]], key: str, metric: Metrics) -> None:
    bucket = target.setdefault(
        key,
        {
            "quantity": 0.0,
            "revenue": 0.0,
            "return_quantity": 0.0,
            "return_amount": 0.0,
            "discount_weight": 0.0,
            "discount_base": 0.0,
        },
    )
    bucket["quantity"] += metric.quantity
    bucket["revenue"] += metric.revenue
    bucket["return_quantity"] += metric.return_quantity
    bucket["return_amount"] += metric.return_amount
    if metric.revenue:
        bucket["discount_weight"] += metric.discount_pct * metric.revenue
        bucket["discount_base"] += metric.revenue


def _finish_metrics(raw: Mapping[str, Mapping[str, float]]) -> dict[str, Metrics]:
    result: dict[str, Metrics] = {}
    for key, bucket in raw.items():
        quantity = float(bucket.get("quantity", 0.0))
        revenue = float(bucket.get("revenue", 0.0))
        discount_base = float(bucket.get("discount_base", 0.0))
        result[key] = Metrics(
            quantity=quantity,
            discount_pct=(float(bucket.get("discount_weight", 0.0)) / discount_base if discount_base else 0.0),
            average_price=(revenue / quantity if quantity else 0.0),
            revenue=revenue,
            return_quantity=float(bucket.get("return_quantity", 0.0)),
            return_amount=float(bucket.get("return_amount", 0.0)),
        )
    return result


def _source_name(source: str | Path | bytes | bytearray | BinaryIO, explicit: str | None) -> str:
    if explicit:
        return explicit
    if isinstance(source, (str, Path)):
        return Path(source).name
    return "report.xlsx"


def _generated(rows) -> tuple[str | None, str]:
    value = rows[-1].values.get("A", "") if rows else ""
    match = re.match(r"(\d{2}\.\d{2}\.\d{4}\s+\d{1,2}:\d{2}:\d{2})\s*(.*)", value)
    if not match:
        return None, ""
    try:
        generated_at = datetime.strptime(match.group(1), "%d.%m.%Y %H:%M:%S").isoformat()
    except ValueError:
        generated_at = match.group(1)
    return generated_at, match.group(2).strip()


def _detect_grouping(rows) -> str:
    return next((row.values.get("A", "") for row in rows[:8] if ";" in row.values.get("A", "")), "")


def _validate_grouping(grouping: str, kind: str) -> None:
    required = EXPECTED_GROUPINGS[kind]
    folded = grouping.casefold()
    if not grouping or not all(part.casefold() in folded for part in required):
        expected = " → ".join(required)
        raise ValueError(f"Неверный тип файла. Для блока «{KIND_LABELS[kind]}» нужны группировки: {expected}.")
    foreign = {
        SALES: ("Менеджер", "Поставщик"),
        CONSULTANTS: ("Магазин", "Поставщик"),
        SUPPLIERS: ("Магазин", "Менеджер", "Камень/вставка"),
    }
    if any(marker.casefold() in folded for marker in foreign[kind]):
        raise ValueError(f"Файл относится к другому блоку, а не к «{KIND_LABELS[kind]}».")


def parse_block_report(
    source: str | Path | bytes | bytearray | BinaryIO,
    *,
    kind: str,
    source_name: str | None = None,
) -> BlockReport:
    if kind not in KIND_LABELS:
        raise ValueError(f"Неизвестный тип управленческого отчета: {kind}")
    rows = list(_iter_rows(source))
    if len(rows) < 8:
        raise ValueError("Файл не похож на выгрузку продаж 1С.")

    title = rows[0].values.get("A", "")
    grouping = _detect_grouping(rows)
    _validate_grouping(grouping, kind)
    period_start, period_end, period_label = _period_from_title(title)
    if period_start is None or period_end is None:
        raise ValueError("Не удалось определить период из заголовка выгрузки.")

    total_row = next((row for row in reversed(rows) if row.values.get("A", "").strip().startswith("Итого")), None)
    if total_row is None:
        raise ValueError("В выгрузке не найдена итоговая строка «Итого». ")
    totals = _metrics(total_row)

    raw_dimensions: dict[str, dict[str, dict[str, float]]] = defaultdict(dict)
    facts: list[HierarchyFact] = []
    stack: dict[int, str] = {}
    supplier_container = False

    for row in rows[6:]:
        label_raw = row.values.get("A", "").strip()
        if label_raw.startswith("Итого"):
            break
        metric = _metrics(row)
        if not label_raw and not any((metric.quantity, metric.revenue, metric.return_quantity, metric.return_amount)):
            continue

        for level in tuple(stack):
            if level >= row.level:
                stack.pop(level, None)
        parent = stack.get(row.level - 1, "")
        label = _clean(label_raw)

        if kind == SALES:
            dimension_by_level = {0: "stores", 1: "stones", 2: "assays", 3: "categories"}
            dimension = dimension_by_level.get(row.level)
            if dimension:
                if dimension == "stores":
                    label = analytics_store_name(label_raw) or "Без магазина"
                elif dimension == "stones" and not label_raw:
                    label = "Без вставки"
                elif dimension == "assays" and not label_raw:
                    label = "Проба не указана"
                elif dimension == "categories" and not label_raw:
                    label = "Группа не указана"
                _add_metric(raw_dimensions[dimension], label, metric)
        elif kind == CONSULTANTS:
            dimension_by_level = {0: "consultants", 1: "assays", 2: "categories"}
            dimension = dimension_by_level.get(row.level)
            if dimension:
                if dimension == "consultants" and not label_raw:
                    label = "Менеджер не указан"
                elif dimension == "assays" and not label_raw:
                    label = "Проба не указана"
                elif dimension == "categories" and not label_raw:
                    label = "Группа не указана"
                _add_metric(raw_dimensions[dimension], label, metric)
        else:
            if row.level == 0:
                label = _clean(label_raw, "Группа не указана")
                _add_metric(raw_dimensions["categories"], label, metric)
                supplier_container = False
            elif row.level == 1:
                supplier_container = label_raw.casefold() == "поставщики"
                if not label_raw:
                    _add_metric(raw_dimensions["suppliers"], UNKNOWN, metric)
            elif row.level == 2 and supplier_container:
                label = _normalize_supplier(label_raw)
                _add_metric(raw_dimensions["suppliers"], label, metric)
                category = stack.get(0, "Группа не указана")
                _add_metric(raw_dimensions["supplier_categories"], f"{label} · {category}", metric)

        facts.append(HierarchyFact(row.number, row.level, parent, label, metric))
        stack[row.level] = label_raw or label

    dimensions = {name: _finish_metrics(values) for name, values in raw_dimensions.items()}
    primary = dimensions.get(PRIMARY_DIMENSION[kind], {})
    primary_quantity = sum(item.quantity for item in primary.values())
    primary_revenue = sum(item.revenue for item in primary.values())
    primary_return_quantity = sum(item.return_quantity for item in primary.values())
    primary_return_amount = sum(item.return_amount for item in primary.values())
    quantity_difference = primary_quantity - totals.quantity
    revenue_difference = primary_revenue - totals.revenue
    return_quantity_difference = primary_return_quantity - totals.return_quantity
    return_amount_difference = primary_return_amount - totals.return_amount

    # 1C can round grouped quantities to whole units while keeping the exact weighted
    # quantity in the final row. Revenue and returns must still match exactly.
    validation: dict[str, float | str | bool] = {
        "primary_dimension": PRIMARY_DIMENSION[kind],
        "primary_rows": len(primary),
        "revenue_difference": revenue_difference,
        "quantity_difference": quantity_difference,
        "return_quantity_difference": return_quantity_difference,
        "return_amount_difference": return_amount_difference,
        "revenue_matches_total": abs(revenue_difference) < 0.5,
        "returns_match_total": abs(return_amount_difference) < 0.5 and abs(return_quantity_difference) < 0.0005,
        "quantity_matches_with_rounding": abs(quantity_difference) <= 1.05,
        "valid": (
            abs(revenue_difference) < 0.5
            and abs(return_amount_difference) < 0.5
            and abs(return_quantity_difference) < 0.0005
            and abs(quantity_difference) <= 1.05
        ),
        "quantity_note": (
            "Строка «Итого» хранит точное количество, а групповые строки 1С могут быть округлены; "
            "допустима разница не более 1,05 единицы."
        ),
    }

    generated_at, generated_by = _generated(rows)
    meta = ReportMeta(
        source_file=_source_name(source, source_name),
        title=title,
        period_label=period_label,
        period_start=period_start.isoformat(),
        period_end=period_end.isoformat(),
        period_days=(period_end - period_start).days + 1,
        generated_at=generated_at,
        generated_by=generated_by,
        grouping_label=grouping,
    )
    return BlockReport(kind, meta, totals, dimensions, facts, validation)


def safe_pct(old: float, new: float) -> float | None:
    if abs(old) < 1e-12:
        return None if abs(new) < 1e-12 else math.inf
    return (new - old) / old * 100.0


def compare_metric_maps(old: Mapping[str, Metrics], new: Mapping[str, Metrics]) -> pd.DataFrame:
    old_total = sum(metric.revenue for metric in old.values())
    new_total = sum(metric.revenue for metric in new.values())
    total_delta = new_total - old_total
    rows: list[dict[str, object]] = []
    for name in sorted(set(old) | set(new), key=str.casefold):
        first = old.get(name, Metrics())
        second = new.get(name, Metrics())
        delta = second.revenue - first.revenue
        rows.append({
            "Позиция": name,
            "Количество · Период 1": first.quantity,
            "Количество · Период 2": second.quantity,
            "Δ количества": second.quantity - first.quantity,
            "Δ количества, %": safe_pct(first.quantity, second.quantity),
            "Выручка · Период 1": first.revenue,
            "Выручка · Период 2": second.revenue,
            "Δ выручки": delta,
            "Δ выручки, %": safe_pct(first.revenue, second.revenue),
            "Чистая выручка · Период 1": first.net_revenue,
            "Чистая выручка · Период 2": second.net_revenue,
            "Δ чистой выручки": second.net_revenue - first.net_revenue,
            "Количество возвратов · Период 1": first.return_quantity,
            "Количество возвратов · Период 2": second.return_quantity,
            "Δ количества возвратов": second.return_quantity - first.return_quantity,
            "Доля · Период 1, %": first.revenue / old_total * 100.0 if old_total else 0.0,
            "Доля · Период 2, %": second.revenue / new_total * 100.0 if new_total else 0.0,
            "Δ доли, п.п.": (
                second.revenue / new_total * 100.0 if new_total else 0.0
            ) - (
                first.revenue / old_total * 100.0 if old_total else 0.0
            ),
            "Средняя цена · Период 1": first.average_price,
            "Средняя цена · Период 2": second.average_price,
            "Δ средней цены, %": safe_pct(first.average_price, second.average_price),
            "Скидка · Период 1, %": first.discount_pct,
            "Скидка · Период 2, %": second.discount_pct,
            "Возвраты · Период 1": first.return_amount,
            "Возвраты · Период 2": second.return_amount,
            "Доля возвратов · Период 1, %": first.return_amount / first.revenue * 100.0 if first.revenue else 0.0,
            "Доля возвратов · Период 2, %": second.return_amount / second.revenue * 100.0 if second.revenue else 0.0,
            "Вклад в изменение, %": delta / total_delta * 100.0 if total_delta else 0.0,
        })
    frame = pd.DataFrame(rows)
    if frame.empty:
        return frame
    return frame.sort_values("Δ выручки", ascending=False).reset_index(drop=True)


def compare_totals(first: Metrics, second: Metrics) -> dict[str, float | None | Metrics]:
    return {
        "first": first,
        "second": second,
        "revenue_pct": safe_pct(first.revenue, second.revenue),
        "quantity_pct": safe_pct(first.quantity, second.quantity),
        "average_price_pct": safe_pct(first.average_price, second.average_price),
        "discount_delta_pp": second.discount_pct - first.discount_pct,
        "return_amount_pct": safe_pct(first.return_amount, second.return_amount),
        "return_quantity_pct": safe_pct(first.return_quantity, second.return_quantity),
        "net_revenue_pct": safe_pct(first.net_revenue, second.net_revenue),
        "return_share_first": first.return_amount / first.revenue * 100.0 if first.revenue else 0.0,
        "return_share_second": second.return_amount / second.revenue * 100.0 if second.revenue else 0.0,
    }


def cross_block_validation(period_reports: Mapping[str, BlockReport]) -> pd.DataFrame:
    reference = period_reports[SALES].totals
    rows: list[dict[str, object]] = []
    for kind in (SALES, CONSULTANTS, SUPPLIERS):
        report = period_reports[kind]
        total = report.totals
        rows.append({
            "Блок": KIND_LABELS[kind],
            "Количество": total.quantity,
            "Выручка": total.revenue,
            "Возвратов, шт.": total.return_quantity,
            "Возвраты": total.return_amount,
            "Δ количества к продажам по магазинам": total.quantity - reference.quantity,
            "Δ выручки к продажам по магазинам": total.revenue - reference.revenue,
            "Δ возвратов к продажам по магазинам": total.return_amount - reference.return_amount,
            "Структура блока сходится": bool(report.validation.get("valid", False)),
        })
    return pd.DataFrame(rows)


def validate_period_bundle(period_reports: Mapping[str, BlockReport]) -> list[str]:
    errors: list[str] = []
    missing = [kind for kind in (SALES, CONSULTANTS, SUPPLIERS) if kind not in period_reports]
    if missing:
        return ["Не загружены блоки: " + ", ".join(KIND_LABELS[kind] for kind in missing)]
    starts = {period_reports[kind].meta.period_start for kind in period_reports}
    ends = {period_reports[kind].meta.period_end for kind in period_reports}
    if len(starts) != 1 or len(ends) != 1:
        errors.append("Внутри одного периода три загруженных отчета имеют разные даты.")
    reference = period_reports[SALES].totals
    for kind in (CONSULTANTS, SUPPLIERS):
        total = period_reports[kind].totals
        if abs(total.revenue - reference.revenue) >= 0.5:
            errors.append(f"Выручка в блоке {KIND_LABELS[kind]} не совпадает с блоком «Продажи по магазинам».")
        if abs(total.return_amount - reference.return_amount) >= 0.5 or abs(total.return_quantity - reference.return_quantity) >= 0.0005:
            errors.append(f"Возвраты в блоке {KIND_LABELS[kind]} не совпадают с блоком «Продажи по магазинам».")
        if abs(total.quantity - reference.quantity) > 0.0005:
            errors.append(f"Итоговое количество в блоке {KIND_LABELS[kind]} не совпадает с блоком «Продажи по магазинам».")
    for kind, report in period_reports.items():
        if not report.validation.get("valid", False):
            errors.append(f"Внутренняя сверка блока {KIND_LABELS[kind]} не пройдена.")
    return errors
