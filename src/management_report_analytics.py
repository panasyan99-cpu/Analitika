from __future__ import annotations

from collections import defaultdict
from dataclasses import asdict
from datetime import date
from typing import Callable, Iterable
import math
import re

import pandas as pd

from src.management_report_parser import Metrics, ParsedReport, ProductFact
from src.management_report_suppliers import SupplierCatalog, UNKNOWN_SUPPLIER
from src.store_normalization import analytics_store_name


TECHNICAL_MANAGER_MARKERS = (
    "admin", "administrator", "cashier", "кассир", "cafe", "gift", "vietnamese staff",
)


def safe_pct(old: float, new: float) -> float | None:
    if abs(old) < 1e-12:
        return None if abs(new) < 1e-12 else math.inf
    return (new - old) / old * 100.0


def canonical_store(value: str) -> str:
    text = analytics_store_name(value)
    folded = text.casefold().replace("ё", "е")
    if "gift" in folded and ("tt" in folded or "тт" in folded):
        return "Gifts-TT"
    if folded in {"cafe", "cafe tt", "кафе"}:
        return "Cafe"
    return text or "Без магазина"


def is_technical_manager(value: str) -> bool:
    folded = " ".join(str(value or "").strip().casefold().split())
    if not folded:
        return True
    return any(marker in folded for marker in TECHNICAL_MANAGER_MARKERS)


def category_label(fact: ProductFact) -> str:
    value = " ".join(str(fact.category or "").strip().split())
    aliases = {
        "bracelets": "Bracelet",
        "jewelry service": "Service",
        "service": "Service",
        "pearl necklaces/bracelets": "Pearl necklaces/bracelets",
        "без номенклатурной группы": "Other",
    }
    return aliases.get(value.casefold(), value or "Other")


def stone_group(value: str) -> str:
    normalized = " ".join(str(value or "").strip().upper().split())
    if not normalized:
        return "Без вставки"
    if any(token in normalized for token in ("SEA PEARL", "AKOYA", "TAHITI", "SOUTH SEA", "GALATEA", "FACETED PEARL")):
        return "Морской жемчуг"
    if any(token in normalized for token in ("FRESH WATER", "FRESHWATER", "FRPW")):
        return "Пресноводный жемчуг"
    if "SAPPHIRE" in normalized:
        return "Сапфиры"
    if "RUBY" in normalized:
        return "Рубины"
    if "MOISSANITE" in normalized:
        return "Муассаниты"
    if any(token in normalized for token in ("TOPAZ", "LONDON BT", "SWISS BT", "WHITE TOPAZ")):
        return "Топазы"
    if any(token in normalized for token in ("EMERALD", "CHROME DIOPSIDE", "GREEN AGATE", "PERIDOT", "GREEN AMETHYST")):
        return "Зеленые камни"
    if any(token in normalized for token in ("CITRINE", "LEMON", "ROSE QUARTZ", "MYSTIC", "QUARTZ")):
        return "Кварц и цитрин"
    if any(token in normalized for token in ("ONYX", "BLACK SPINEL", "BLACK AGATE")):
        return "Черные камни"
    if "GARNET" in normalized or "RHODOLITE" in normalized:
        return "Гранат и родолит"
    if "CZ" in normalized:
        return "CZ"
    return "Прочие вставки"


def assay_group(value: str) -> str:
    normalized = " ".join(str(value or "").strip().upper().split())
    if normalized.startswith("B 925") or normalized in {"925", "AG 925"}:
        return "Серебро"
    if normalized.startswith("AU") or normalized.startswith("PT"):
        return "Золото и платина"
    return "Другое"


def _aggregate(
    facts: Iterable[ProductFact],
    key_fn: Callable[[ProductFact], str],
) -> dict[str, Metrics]:
    buckets: dict[str, dict[str, float]] = defaultdict(
        lambda: {
            "quantity": 0.0,
            "revenue": 0.0,
            "return_quantity": 0.0,
            "return_amount": 0.0,
            "discount_weight": 0.0,
            "discount_base": 0.0,
        }
    )
    for fact in facts:
        key = " ".join(str(key_fn(fact) or "").strip().split()) or "(не указано)"
        metric = fact.metrics
        bucket = buckets[key]
        bucket["quantity"] += metric.quantity
        bucket["revenue"] += metric.revenue
        bucket["return_quantity"] += metric.return_quantity
        bucket["return_amount"] += metric.return_amount
        if metric.revenue:
            bucket["discount_weight"] += metric.discount_pct * metric.revenue
            bucket["discount_base"] += metric.revenue
    result: dict[str, Metrics] = {}
    for key, bucket in buckets.items():
        quantity = bucket["quantity"]
        revenue = bucket["revenue"]
        result[key] = Metrics(
            quantity=quantity,
            discount_pct=(bucket["discount_weight"] / bucket["discount_base"] if bucket["discount_base"] else 0.0),
            average_price=(revenue / quantity if quantity else 0.0),
            revenue=revenue,
            return_quantity=bucket["return_quantity"],
            return_amount=bucket["return_amount"],
        )
    return result


def _compare_maps(old: dict[str, Metrics], new: dict[str, Metrics]) -> pd.DataFrame:
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


def _supplier_maps(report: ParsedReport, catalog: SupplierCatalog) -> tuple[dict[str, Metrics], dict[str, float], pd.DataFrame]:
    resolved_rows: list[tuple[ProductFact, str, str]] = []
    source_revenue: dict[str, float] = defaultdict(float)
    unknown: dict[str, dict[str, object]] = defaultdict(lambda: {"Количество": 0.0, "Выручка": 0.0, "Камень": set(), "Категория": set()})
    supplier_by_fact_id: dict[int, str] = {}
    for fact in report.facts:
        resolution = catalog.resolve(fact.sku)
        resolved_rows.append((fact, resolution.supplier, resolution.source))
        supplier_by_fact_id[id(fact)] = resolution.supplier
        source_revenue[resolution.source] += fact.metrics.revenue
        if resolution.supplier == UNKNOWN_SUPPLIER:
            row = unknown[fact.sku]
            row["Количество"] = float(row["Количество"]) + fact.metrics.quantity
            row["Выручка"] = float(row["Выручка"]) + fact.metrics.revenue
            row["Камень"].add(fact.stone or "Без вставки")
            row["Категория"].add(category_label(fact))

    maps = _aggregate((fact for fact, _, _ in resolved_rows), lambda fact: supplier_by_fact_id[id(fact)])
    unknown_rows = []
    for sku, row in unknown.items():
        unknown_rows.append({
            "SKU": sku,
            "Количество": row["Количество"],
            "Выручка": row["Выручка"],
            "Камень": "; ".join(sorted(row["Камень"])),
            "Категория": "; ".join(sorted(row["Категория"])),
        })
    unknown_frame = pd.DataFrame(unknown_rows)
    if not unknown_frame.empty:
        unknown_frame = unknown_frame.sort_values("Выручка", ascending=False).reset_index(drop=True)
    return maps, dict(source_revenue), unknown_frame


def _date_from_iso(value: str | None) -> date | None:
    if not value:
        return None
    try:
        return date.fromisoformat(value)
    except ValueError:
        return None


def _overall(first: ParsedReport, second: ParsedReport) -> dict[str, object]:
    old = first.totals
    new = second.totals
    first_days = first.meta.period_days or 1
    second_days = second.meta.period_days or 1
    return {
        "old": asdict(old),
        "new": asdict(new),
        "revenue_delta": new.revenue - old.revenue,
        "revenue_pct": safe_pct(old.revenue, new.revenue),
        "quantity_delta": new.quantity - old.quantity,
        "quantity_pct": safe_pct(old.quantity, new.quantity),
        "average_price_delta": new.average_price - old.average_price,
        "average_price_pct": safe_pct(old.average_price, new.average_price),
        "discount_delta_pp": new.discount_pct - old.discount_pct,
        "return_amount_delta": new.return_amount - old.return_amount,
        "return_amount_pct": safe_pct(old.return_amount, new.return_amount),
        "return_share_old": old.return_amount / old.revenue * 100.0 if old.revenue else 0.0,
        "return_share_new": new.return_amount / new.revenue * 100.0 if new.revenue else 0.0,
        "net_revenue_old": old.net_revenue,
        "net_revenue_new": new.net_revenue,
        "net_revenue_delta": new.net_revenue - old.net_revenue,
        "net_revenue_pct": safe_pct(old.net_revenue, new.net_revenue),
        "daily": {
            "old_revenue": old.revenue / first_days,
            "new_revenue": new.revenue / second_days,
            "revenue_pct": safe_pct(old.revenue / first_days, new.revenue / second_days),
            "old_net_revenue": old.net_revenue / first_days,
            "new_net_revenue": new.net_revenue / second_days,
            "net_revenue_pct": safe_pct(old.net_revenue / first_days, new.net_revenue / second_days),
            "old_quantity": old.quantity / first_days,
            "new_quantity": new.quantity / second_days,
            "quantity_pct": safe_pct(old.quantity / first_days, new.quantity / second_days),
            "old_return_quantity": old.return_quantity / first_days,
            "new_return_quantity": new.return_quantity / second_days,
            "return_quantity_pct": safe_pct(old.return_quantity / first_days, new.return_quantity / second_days),
            "old_return_amount": old.return_amount / first_days,
            "new_return_amount": new.return_amount / second_days,
            "return_pct": safe_pct(old.return_amount / first_days, new.return_amount / second_days),
        },
    }


def build_management_snapshot(
    first: ParsedReport,
    second: ParsedReport,
    catalog: SupplierCatalog,
) -> dict[str, object]:
    first_start = _date_from_iso(first.meta.period_start)
    second_start = _date_from_iso(second.meta.period_start)
    if first_start and second_start and first_start > second_start:
        first, second = second, first

    first_suppliers, first_supplier_sources, first_unknown = _supplier_maps(first, catalog)
    second_suppliers, second_supplier_sources, second_unknown = _supplier_maps(second, catalog)

    first_manager_facts = [fact for fact in first.facts if not is_technical_manager(fact.manager)]
    second_manager_facts = [fact for fact in second.facts if not is_technical_manager(fact.manager)]

    dimensions = {
        "stores": _compare_maps(
            _aggregate(first.facts, lambda fact: canonical_store(fact.store)),
            _aggregate(second.facts, lambda fact: canonical_store(fact.store)),
        ),
        "managers": _compare_maps(
            _aggregate(first_manager_facts, lambda fact: fact.manager),
            _aggregate(second_manager_facts, lambda fact: fact.manager),
        ),
        "suppliers": _compare_maps(first_suppliers, second_suppliers),
        "categories": _compare_maps(
            _aggregate(first.facts, category_label),
            _aggregate(second.facts, category_label),
        ),
        "stone_groups": _compare_maps(
            _aggregate(first.facts, lambda fact: stone_group(fact.stone)),
            _aggregate(second.facts, lambda fact: stone_group(fact.stone)),
        ),
        "stones": _compare_maps(
            _aggregate(first.facts, lambda fact: fact.stone or "Без вставки"),
            _aggregate(second.facts, lambda fact: fact.stone or "Без вставки"),
        ),
        "assay_groups": _compare_maps(
            _aggregate(first.facts, lambda fact: assay_group(fact.assay)),
            _aggregate(second.facts, lambda fact: assay_group(fact.assay)),
        ),
        "assays": _compare_maps(
            _aggregate(first.facts, lambda fact: fact.assay or "Не указана"),
            _aggregate(second.facts, lambda fact: fact.assay or "Не указана"),
        ),
        "sku": _compare_maps(
            _aggregate(first.facts, lambda fact: fact.sku),
            _aggregate(second.facts, lambda fact: fact.sku),
        ),
    }

    first_classified = first.totals.revenue - first_suppliers.get(UNKNOWN_SUPPLIER, Metrics()).revenue
    second_classified = second.totals.revenue - second_suppliers.get(UNKNOWN_SUPPLIER, Metrics()).revenue
    supplier_quality = {
        "old_revenue_coverage_pct": first_classified / first.totals.revenue * 100.0 if first.totals.revenue else 0.0,
        "new_revenue_coverage_pct": second_classified / second.totals.revenue * 100.0 if second.totals.revenue else 0.0,
        "old_source_revenue": first_supplier_sources,
        "new_source_revenue": second_supplier_sources,
        "old_unknown": first_unknown,
        "new_unknown": second_unknown,
    }

    outlet_members = {"TT", "Gifts-TT", "Cafe"}
    outlet_first = _aggregate((fact for fact in first.facts if canonical_store(fact.store) in outlet_members), lambda _: "Outlet").get("Outlet", Metrics())
    outlet_second = _aggregate((fact for fact in second.facts if canonical_store(fact.store) in outlet_members), lambda _: "Outlet").get("Outlet", Metrics())

    return {
        "first": first,
        "second": second,
        "overall": _overall(first, second),
        "dimensions": dimensions,
        "supplier_quality": supplier_quality,
        "outlet": _compare_maps({"Outlet": outlet_first}, {"Outlet": outlet_second}),
        "validation": {"first": first.validation, "second": second.validation},
    }


def significant_rows(frame: pd.DataFrame, *, positive: bool, limit: int = 5) -> pd.DataFrame:
    if frame.empty:
        return frame
    scoped = frame.loc[frame["Δ выручки"] > 0] if positive else frame.loc[frame["Δ выручки"] < 0]
    scoped = scoped.copy()
    scoped["_importance"] = scoped["Δ выручки"].abs()
    return scoped.nlargest(limit, "_importance").drop(columns="_importance")


def new_and_lost_sku(frame: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    if frame.empty:
        return frame, frame
    new = frame.loc[(frame["Выручка · Период 1"] == 0) & (frame["Выручка · Период 2"] > 0)].copy()
    lost = frame.loc[(frame["Выручка · Период 1"] > 0) & (frame["Выручка · Период 2"] == 0)].copy()
    return new.sort_values("Выручка · Период 2", ascending=False), lost.sort_values("Выручка · Период 1", ascending=False)
