from __future__ import annotations

import io
from zipfile import ZipFile

from src.order_workflow import (
    CATEGORY_MEDIUM,
    CATEGORY_TOP,
    CATEGORY_WEAK,
    ORDER_MODE_STONES,
    RECOMMENDATION_BASE,
    RECOMMENDATION_SEASONAL,
    OrderItem,
    OrderSet,
    _annotate_pendant_duplicates,
    build_order_recommendation,
    earring_lock_code,
    ordinary_transferable_stock,
    report_month_count,
)


def item(
    sku: str,
    *,
    group: str,
    sales: int = 0,
    tt: int = 0,
    stores: dict[str, int] | None = None,
    working: int | None = None,
    stock_tt_warehouse: int = 0,
    category: str = CATEGORY_WEAK,
    set_id: str = "Set# 1101",
    eligible_store_count: int = 4,
    row: int = 10,
    duplicate_status: str | None = None,
) -> tuple[OrderItem, OrderSet]:
    stores = dict(stores or {"TT": tt})
    if stock_tt_warehouse and "Stock TT" not in stores:
        stores["Stock TT"] = stock_tt_warehouse
    if working is None:
        working = sum(stores.values())
    current = OrderItem(
        row=row,
        set_id=set_id,
        sku=sku,
        stone="Ruby",
        group=group,
        sales=sales,
        stock_63=0,
        stock_20=0,
        stores=stores,
        total_stock=working,
        working_stock=working,
        ntr2_stock=0,
        ntr2_calculated=False,
        tvp_raw=0,
        stock_tt=tt,
        stock_tt_warehouse=stock_tt_warehouse,
        report_months=4,
        eligible_store_count=eligible_store_count,
        duplicate_status=duplicate_status,
    )
    current_set = OrderSet(
        key="1101|set",
        set_id=set_id,
        stone="Ruby",
        items=(current,),
        category=category,
        driver_sku=sku,
        max_sales=sales,
        has_positive_tvp=False,
        has_negative_tvp=False,
    )
    return current, current_set


def quantity(pair: tuple[OrderItem, OrderSet], profile: str = RECOMMENDATION_BASE) -> int:
    current, current_set = pair
    return build_order_recommendation(current, current_set, ORDER_MODE_STONES, profile).quantity


def test_report_duration_and_seasonal_minimum_order() -> None:
    assert report_month_count("01.03.2026 — 30.06.2026") == 4

    current, current_set = item("PD-SEASON", group="Pendant", sales=8, tt=3, working=0)
    assert build_order_recommendation(current, current_set, ORDER_MODE_STONES, RECOMMENDATION_BASE).quantity == 4
    assert build_order_recommendation(current, current_set, ORDER_MODE_STONES, RECOMMENDATION_SEASONAL).quantity == 3

    current2, current_set2 = item("PD-SEASON-2", group="Pendant", sales=4, tt=3, working=0)
    assert build_order_recommendation(current2, current_set2, ORDER_MODE_STONES, RECOMMENDATION_BASE).quantity == 3
    assert build_order_recommendation(current2, current_set2, ORDER_MODE_STONES, RECOMMENDATION_SEASONAL).quantity == 0


def test_earrings_tt_zero_matrix() -> None:
    assert quantity(item("ER-12-0", group="Earrings", sales=12, tt=0, working=0)) == 6
    assert quantity(item(
        "ER-12-FOUR",
        group="Earrings",
        sales=12,
        tt=0,
        stores={"TT": 0, "AB": 1, "NTR1": 1, "NTR2": 1, "SCR": 1},
        working=4,
    )) == 5
    assert quantity(item(
        "ER-12-MOVE",
        group="Earrings",
        sales=12,
        tt=0,
        stores={"TT": 0, "AB": 2, "NTR1": 1, "NTR2": 1, "SCR": 1},
        working=5,
    )) == 3
    assert quantity(item(
        "ER-8-STOCK",
        group="Earrings",
        sales=8,
        tt=0,
        stores={"TT": 0, "Stock TT": 3},
        working=3,
        stock_tt_warehouse=3,
    )) == 5
    assert quantity(item("ER-4", group="Earrings", sales=4, tt=0, working=4)) == 3
    assert quantity(item("ER-2", group="Earrings", sales=2, tt=0, working=2)) == 3


def test_earrings_tt_one_two_and_three_matrix() -> None:
    assert quantity(item("ER-TT1-ACTIVE", group="Earrings", sales=12, tt=1, working=1)) == 5
    assert quantity(item(
        "ER-TT1-MOVE",
        group="Earrings",
        sales=12,
        tt=1,
        stores={"TT": 1, "AB": 2, "NTR1": 1},
        working=4,
    )) == 3
    assert quantity(item(
        "ER-TT1-WH",
        group="Earrings",
        sales=12,
        tt=1,
        stores={"TT": 1, "Stock TT": 4},
        working=5,
        stock_tt_warehouse=4,
    )) == 5
    assert quantity(item("ER-TT1-2", group="Earrings", sales=2, tt=1, working=1)) == 3
    assert quantity(item("ER-TT1-1", group="Earrings", sales=1, tt=1, working=1)) == 0

    assert quantity(item("ER-TT2-3", group="Earrings", sales=3, tt=2, working=2)) == 0
    assert quantity(item("ER-TT2-6", group="Earrings", sales=6, tt=2, working=2)) == 3
    assert quantity(item(
        "ER-TT2-MOVE",
        group="Earrings",
        sales=6,
        tt=2,
        stores={"TT": 2, "Stock TT": 1},
        working=3,
        stock_tt_warehouse=1,
    )) == 0
    assert quantity(item("ER-TT2-8", group="Earrings", sales=8, tt=2, working=2)) == 3
    assert quantity(item("ER-TT3", group="Earrings", sales=12, tt=3, working=5)) == 0


def test_last_shop_unit_is_not_transferable_but_stock_tt_is() -> None:
    current, _ = item(
        "ER-TRANSFER",
        group="Earrings",
        stores={"TT": 0, "AB": 1, "NTR1": 2, "Stock TT": 3},
        working=6,
        stock_tt_warehouse=3,
    )
    assert ordinary_transferable_stock(current) == 1


def test_ring_completeness_and_network_fill() -> None:
    assert quantity(item(
        "RG-TOP-EMPTY",
        group="Ring",
        sales=6,
        tt=0,
        working=0,
        category=CATEGORY_TOP,
        eligible_store_count=4,
    )) == 6
    assert quantity(item(
        "RG-WEAK-EMPTY",
        group="Ring",
        sales=1,
        tt=0,
        working=0,
        category=CATEGORY_WEAK,
    )) == 3
    assert quantity(item(
        "RG-TOP-TT1",
        group="Ring",
        sales=6,
        tt=1,
        working=1,
        category=CATEGORY_TOP,
    )) == 3
    assert quantity(item(
        "RG-MID-TT2",
        group="Ring",
        sales=3,
        tt=2,
        working=2,
        category=CATEGORY_MEDIUM,
    )) == 0


def test_pendant_matrix() -> None:
    assert quantity(item("PD-ONE", group="Pendant", sales=1, tt=0, working=0)) == 3
    assert quantity(item("PD-TWO", group="Pendant", sales=2, tt=0, working=0)) == 3
    assert quantity(item(
        "PD-STOCK-TT",
        group="Pendant",
        sales=2,
        tt=0,
        stores={"TT": 0, "Stock TT": 1},
        working=1,
        stock_tt_warehouse=1,
    )) == 3
    assert quantity(item("PD-TOP-ZERO", group="Pendant", sales=0, tt=0, working=0, category=CATEGORY_TOP)) == 3
    assert quantity(item("PD-TT1-ONE", group="Pendant", sales=1, tt=1, working=1)) == 0
    assert quantity(item("PD-TT1-TWO", group="Pendant", sales=2, tt=1, working=1)) == 3
    assert quantity(item("PD-TT2-FOUR", group="Pendant", sales=4, tt=2, working=2)) == 0
    assert quantity(item("PD-TT2-SIX", group="Pendant", sales=6, tt=2, working=2)) == 3
    assert quantity(item("PD-TT2-EIGHT", group="Pendant", sales=8, tt=2, working=2)) == 3


def test_cross_targets_are_total_stock_targets() -> None:
    assert quantity(item(
        "KP210147",
        group="Pendant",
        sales=8,
        tt=3,
        working=2,
        category=CATEGORY_TOP,
        set_id="Кресты BS",
    )) == 3
    assert quantity(item(
        "KP210148",
        group="Pendant",
        sales=1,
        tt=1,
        working=1,
        set_id="Кресты BS",
    )) == 3
    assert quantity(item(
        "KP210149",
        group="Pendant",
        sales=1,
        tt=2,
        working=2,
        set_id="Кресты BS",
    )) == 0


def test_duplicate_selection_prefers_active_then_newer_model() -> None:
    old, _ = item("KP160147", group="Pendant", sales=0, tt=0, working=0, set_id="Кресты BS", row=10)
    active, _ = item("SKP210147", group="Pendant", sales=2, tt=0, working=1, set_id="Кресты BS", row=11)
    with ZipFile(io.BytesIO(), "w") as archive:
        annotated = _annotate_pendant_duplicates(archive, [old, active])
    old_result = next(value for value in annotated if value.sku == old.sku)
    active_result = next(value for value in annotated if value.sku == active.sku)
    assert old_result.duplicate_status == "suppress"
    assert active_result.duplicate_status == "preferred"

    old2, _ = item("KP160222", group="Pendant", sales=0, tt=0, working=0, set_id="Кресты BS", row=12)
    new2, _ = item("SKP210222", group="Pendant", sales=0, tt=0, working=0, set_id="Кресты BS", row=13)
    with ZipFile(io.BytesIO(), "w") as archive:
        annotated2 = _annotate_pendant_duplicates(archive, [old2, new2])
    assert next(value for value in annotated2 if value.sku == old2.sku).duplicate_status == "suppress"
    assert next(value for value in annotated2 if value.sku == new2.sku).duplicate_status == "preferred"


def test_earring_lock_code_is_read_after_year() -> None:
    assert earring_lock_code("SKE17A004") == "A"
    assert earring_lock_code("SKE23B001B") == "B"
    assert earring_lock_code("SKE20C046") == "C"
    assert earring_lock_code("SKE21D040") == "D"
    assert earring_lock_code("KE21D040") == "D"
