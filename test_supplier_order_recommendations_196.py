from __future__ import annotations

import pytest

from src.order_workflow import (
    CATEGORY_MEDIUM,
    CATEGORY_TOP,
    ORDER_MODE_PEARLS,
    ORDER_MODE_STONES,
    OrderItem,
    OrderSet,
    build_order_recommendation,
)


def item(
    sku: str,
    *,
    stone: str = "FRESH WATER PEARL - BLACK",
    group: str,
    sales: int,
    stock: int,
    tt: int,
    stock63: int = 0,
    tvp: int = 0,
    row: int = 10,
) -> OrderItem:
    return OrderItem(
        row=row,
        set_id="Set# 2 020",
        sku=sku,
        stone=stone,
        group=group,
        sales=sales,
        stock_63=stock63,
        stock_20=0,
        stores={"TT": tt, "63": stock63},
        total_stock=stock + stock63,
        working_stock=stock,
        ntr2_stock=0,
        ntr2_calculated=False,
        tvp_raw=tvp,
        stock_tt=tt,
    )


def order_set(*items: OrderItem, category: str = CATEGORY_MEDIUM) -> OrderSet:
    driver = max(items, key=lambda value: (value.sales, -value.row))
    return OrderSet(
        key="196|set",
        set_id="Set# 2 020",
        stone=items[0].stone,
        items=tuple(items),
        category=category,
        driver_sku=driver.sku,
        max_sales=driver.sales,
        has_positive_tvp=any(value.tvp_raw > 0 for value in items),
        has_negative_tvp=False,
        is_ungrouped=False,
    )


@pytest.mark.parametrize(
    ("category", "expected"),
    [(CATEGORY_MEDIUM, 3), (CATEGORY_TOP, 4)],
)
def test_weak_ring_inside_stronger_set_gets_minimum_assortment_batch(category: str, expected: int):
    driver = item("KE16E210B-FPB", group="Earrings", sales=5 if category == CATEGORY_TOP else 3, stock=1, tt=0)
    ring = item("KR160210B-FPB", group="Ring", sales=2, stock=1, tt=1, row=11)
    current_set = order_set(driver, ring, category=category)

    rec = build_order_recommendation(ring, current_set, ORDER_MODE_PEARLS)

    assert rec.quantity == expected
    assert rec.rule == "weak_item"
    assert any("Слабая позиция внутри комплекта" in reason for reason in rec.reasons)


def test_actual_kr160210b_fpb_scenario_is_not_suppressed_as_one_unit_shortage():
    earrings = item("KE16E210B-FPB", group="Earrings", sales=3, stock=0, tt=0, stock63=1)
    pendant = item("KP160210B-FPB", group="Pendant", sales=2, stock=1, tt=0, row=11)
    ring = item("KR160210B-FPB", group="Ring", sales=2, stock=1, tt=1, row=12)
    current_set = order_set(earrings, pendant, ring, category=CATEGORY_MEDIUM)

    rec = build_order_recommendation(ring, current_set, ORDER_MODE_PEARLS)

    assert rec.quantity == 3
    assert rec.rule == "weak_item"
    assert not any("1 шт.; такую автоматическую партию" in reason for reason in rec.reasons)


def test_weak_item_with_five_units_and_tt_presence_is_not_reordered():
    driver = item("KE-DRIVER", group="Earrings", sales=3, stock=6, tt=3)
    ring = item("KR-STOCKED", group="Ring", sales=2, stock=5, tt=1, row=11)
    current_set = order_set(driver, ring, category=CATEGORY_MEDIUM)

    rec = build_order_recommendation(ring, current_set, ORDER_MODE_PEARLS)

    assert rec.quantity == 0
    assert rec.rule == "none"


def test_other_colour_weak_ring_in_stronger_set_still_finishes_at_three():
    driver = item("ER-CITRINE", stone="Citrine", group="Earrings", sales=3, stock=6, tt=3)
    ring = item("RG-CITRINE", stone="Citrine", group="Ring", sales=2, stock=1, tt=1, row=11)
    current_set = order_set(driver, ring, category=CATEGORY_MEDIUM)

    rec = build_order_recommendation(ring, current_set, ORDER_MODE_STONES)

    assert rec.quantity == 3  # base 3 -> colour correction 2 -> minimum batch 3
    assert rec.rule == "weak_item"
    assert any("3 → 2" in reason for reason in rec.reasons)
    assert any("минимальная автоматическая партия" in reason.lower() for reason in rec.reasons)


def test_positive_tvp_still_blocks_the_same_weak_item():
    driver = item("KE-DRIVER", group="Earrings", sales=3, stock=1, tt=0)
    ring = item("KR-TVP", group="Ring", sales=2, stock=1, tt=1, tvp=3, row=11)
    current_set = order_set(driver, ring, category=CATEGORY_MEDIUM)

    rec = build_order_recommendation(ring, current_set, ORDER_MODE_PEARLS)

    assert rec.quantity == 0
    assert rec.blocked_by_tvp is True
    assert rec.rule == "tvp"
