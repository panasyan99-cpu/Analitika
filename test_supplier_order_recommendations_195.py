from __future__ import annotations

import pytest

from src.order_workflow import (
    CATEGORY_TOP,
    CATEGORY_WEAK,
    CATEGORY_ZERO,
    ORDER_MODE_PEARLS,
    ORDER_MODE_STONES,
    OrderItem,
    OrderSet,
    build_order_recommendation,
)


def item(
    sku: str,
    *,
    stone: str,
    group: str,
    sales: int = 0,
    stock: int = 0,
    tt: int = 0,
    stock63: int = 0,
    tvp: int = 0,
    row: int = 10,
    ungrouped: bool = False,
) -> OrderItem:
    return OrderItem(
        row=row,
        set_id="Без комплекта" if ungrouped else "Set# 195",
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
        ungrouped=ungrouped,
    )


def order_set(
    *items: OrderItem,
    category: str = CATEGORY_WEAK,
    ungrouped: bool = False,
) -> OrderSet:
    return OrderSet(
        key="195|set",
        set_id="Без комплекта" if ungrouped else "Set# 195",
        stone=items[0].stone,
        items=tuple(items),
        category=category,
        driver_sku=max(items, key=lambda value: value.sales).sku,
        max_sales=max(value.sales for value in items),
        has_positive_tvp=any(value.tvp_raw > 0 for value in items),
        has_negative_tvp=False,
        is_ungrouped=ungrouped,
    )


@pytest.mark.parametrize(
    "stone",
    [
        "Ruby",
        "Blue Sapphire",
        "Fancy Sapphire",
        "London Topaz",
        "Mystic Topaz",
        "Moissanite",
        "Emerald",
        "Created Emerald",
        "Onyx",
        "Black Spinel",
        "Green Agate",
    ],
)
def test_priority_stones_keep_full_pearl_scale(stone: str):
    earrings = item("ER-FULL", stone=stone, group="Earrings", sales=2, stock=1, tt=0)
    rec = build_order_recommendation(earrings, order_set(earrings), ORDER_MODE_STONES)
    assert rec.quantity == 5
    assert not any("уменьшена на 1" in reason for reason in rec.reasons)


def test_other_coloured_stones_are_reduced_after_base_calculation():
    earrings = item("ER-AM", stone="Amethyst", group="Earrings", sales=2, stock=1, tt=0)
    top_ring = item("RG-AM", stone="Amethyst", group="Ring", sales=6, stock=0, tt=0)

    earrings_rec = build_order_recommendation(earrings, order_set(earrings), ORDER_MODE_STONES)
    ring_rec = build_order_recommendation(top_ring, order_set(top_ring, category=CATEGORY_TOP), ORDER_MODE_STONES)

    assert earrings_rec.quantity == 4  # base 5, coloured-stone correction to 4
    assert ring_rec.quantity == 3  # base 4, correction to 3
    assert any("уменьшена на 1" in reason for reason in earrings_rec.reasons)


def test_final_batch_rule_suppresses_one_and_converts_two_to_three():
    # Full-scale standalone pendant with stock 2 and TT 0 reaches raw compact
    # top-up 1, which must be suppressed.
    one = item(
        "PD-ONE",
        stone="Ruby",
        group="Pendant",
        sales=0,
        stock=2,
        tt=0,
        ungrouped=True,
    )
    one_rec = build_order_recommendation(
        one,
        order_set(one, category=CATEGORY_ZERO, ungrouped=True),
        ORDER_MODE_STONES,
    )
    assert one_rec.quantity == 0
    assert any("1 шт." in reason for reason in one_rec.reasons)

    # A full-scale pendant with zero stock has raw recommendation 2 and is
    # rounded to the minimum automatic batch of 3.
    two = item(
        "PD-TWO",
        stone="Ruby",
        group="Pendant",
        sales=0,
        stock=0,
        tt=0,
        ungrouped=True,
    )
    two_rec = build_order_recommendation(
        two,
        order_set(two, category=CATEGORY_ZERO, ungrouped=True),
        ORDER_MODE_STONES,
    )
    assert two_rec.quantity == 3
    assert any("минимальная автоматическая партия" in reason.lower() for reason in two_rec.reasons)


def test_other_colour_reduction_is_applied_before_minimum_batch_rule():
    # Base 3 -> colour correction 2 -> minimum automatic batch 3.
    ring = item("RG-COLOR", stone="Citrine", group="Ring", sales=2, stock=1, tt=0)
    ring_rec = build_order_recommendation(ring, order_set(ring), ORDER_MODE_STONES)
    assert ring_rec.quantity == 3
    assert any("3 → 2" in reason for reason in ring_rec.reasons)
    assert any("2 шт." in reason for reason in ring_rec.reasons)

    # Base 2 -> colour correction 1 -> no automatic order.
    pendant = item(
        "PD-COLOR",
        stone="Citrine",
        group="Pendant",
        sales=0,
        stock=0,
        tt=0,
        ungrouped=True,
    )
    pendant_rec = build_order_recommendation(
        pendant,
        order_set(pendant, category=CATEGORY_ZERO, ungrouped=True),
        ORDER_MODE_STONES,
    )
    assert pendant_rec.quantity == 0
    assert any("2 → 1" in reason for reason in pendant_rec.reasons)


def test_pearls_always_keep_full_scale_and_minimum_batch():
    pendant = item(
        "PD-PEARL",
        stone="Freshwater Pearl - White",
        group="Pendant",
        sales=0,
        stock=0,
        tt=0,
        ungrouped=True,
    )
    rec = build_order_recommendation(
        pendant,
        order_set(pendant, category=CATEGORY_ZERO, ungrouped=True),
        ORDER_MODE_PEARLS,
    )
    assert rec.quantity == 3
    assert not any("цветных камней" in reason for reason in rec.reasons)


def test_tvp_set_balance_is_also_reduced_for_other_colours():
    earrings = item("ER-TVP10", stone="Amethyst", group="Earrings", tvp=10)
    ring = item("RG-MISSING", stone="Amethyst", group="Ring", stock=0, tt=0, row=11)
    current_set = order_set(earrings, ring, category=CATEGORY_ZERO)

    earrings_rec = build_order_recommendation(earrings, current_set, ORDER_MODE_STONES)
    ring_rec = build_order_recommendation(ring, current_set, ORDER_MODE_STONES)

    assert earrings_rec.quantity == 0
    assert earrings_rec.blocked_by_tvp is True
    assert ring_rec.quantity == 5  # balance base 6, then colour correction to 5
    assert any("серьги находятся в пути" in reason.lower() for reason in ring_rec.reasons)
