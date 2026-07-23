from __future__ import annotations

from src.order_workflow import (
    CATEGORY_WEAK,
    CATEGORY_ZERO,
    ORDER_MODE_PEARLS,
    OrderItem,
    OrderSet,
    build_order_recommendation,
)


def item(
    sku: str,
    *,
    group: str,
    sales: int = 0,
    stock: int = 0,
    tt: int = 0,
    stock63: int = 0,
    tvp: int = 0,
    row: int = 10,
    set_id: str = "Set# 286",
    ungrouped: bool = False,
) -> OrderItem:
    return OrderItem(
        row=row,
        set_id=set_id,
        sku=sku,
        stone="Freshwater Pearl - White",
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


def order_set(*items: OrderItem, category: str = CATEGORY_WEAK, ungrouped: bool = False) -> OrderSet:
    return OrderSet(
        key="pearls|set",
        set_id="Без комплекта" if ungrouped else "Set# 286",
        stone="Freshwater Pearl - White",
        items=tuple(items),
        category=category,
        driver_sku=max(items, key=lambda value: value.sales).sku,
        max_sales=max(value.sales for value in items),
        has_positive_tvp=any(value.tvp_raw > 0 for value in items),
        has_negative_tvp=False,
        is_ungrouped=ungrouped,
    )


def test_weak_real_set_recommends_five_earrings_and_three_rings():
    earrings = item("ER-WEAK", group="Earrings", sales=2, stock=0, stock63=1, tt=0)
    ring = item("RG-WEAK", group="Ring", sales=2, stock=0, stock63=1, tt=0, row=11)
    current_set = order_set(earrings, ring)

    earrings_rec = build_order_recommendation(earrings, current_set, ORDER_MODE_PEARLS)
    ring_rec = build_order_recommendation(ring, current_set, ORDER_MODE_PEARLS)

    assert earrings_rec.quantity == 5
    assert ring_rec.quantity == 3
    assert any("Слабый комплект" in reason for reason in earrings_rec.reasons)


def test_weak_sales_with_tt_one_or_two_still_replenish_but_five_total_does_not():
    tt_one = item("ER-TT1", group="Earrings", sales=2, stock=1, tt=1)
    tt_two = item("ER-TT2", group="Earrings", sales=2, stock=2, tt=2)
    enough = item("ER-ENOUGH", group="Earrings", sales=2, stock=5, tt=1)

    assert build_order_recommendation(tt_one, order_set(tt_one), ORDER_MODE_PEARLS).quantity == 5
    assert build_order_recommendation(tt_two, order_set(tt_two), ORDER_MODE_PEARLS).quantity == 5
    assert build_order_recommendation(enough, order_set(enough), ORDER_MODE_PEARLS).quantity == 0


def test_zero_real_set_creates_minimum_assortment():
    earrings = item("ER-ZERO", group="Earrings")
    ring = item("RG-ZERO", group="Ring", row=11)
    pendant = item("PD-ZERO", group="Pendant", row=12)
    current_set = order_set(earrings, ring, pendant, category=CATEGORY_ZERO)

    assert build_order_recommendation(earrings, current_set, ORDER_MODE_PEARLS).quantity == 5
    assert build_order_recommendation(ring, current_set, ORDER_MODE_PEARLS).quantity == 3
    assert build_order_recommendation(pendant, current_set, ORDER_MODE_PEARLS).quantity == 2


def test_standalone_zero_stock_uses_independent_type_targets_without_set_wording():
    earrings = item("ER-SOLO", group="Earrings", ungrouped=True, set_id="Без комплекта")
    ring = item("RG-SOLO", group="Ring", row=11, ungrouped=True, set_id="Без комплекта")
    pendant = item("PD-SOLO", group="Pendant", row=12, ungrouped=True, set_id="Без комплекта")

    for current, expected in ((earrings, 5), (ring, 3), (pendant, 2)):
        rec = build_order_recommendation(current, order_set(current, category=CATEGORY_ZERO, ungrouped=True), ORDER_MODE_PEARLS)
        assert rec.quantity == expected
        assert all("комплект" not in reason.lower() for reason in rec.reasons)


def test_standalone_small_stock_without_tt_uses_compact_tt_top_up():
    earrings = item("ER-SMALL", group="Earrings", sales=2, stock=3, tt=0, ungrouped=True, set_id="Без комплекта")
    ring = item("RG-SMALL", group="Ring", sales=1, stock=2, tt=0, ungrouped=True, set_id="Без комплекта")
    pendant = item("PD-SMALL", group="Pendant", sales=0, stock=1, tt=0, ungrouped=True, set_id="Без комплекта")

    assert build_order_recommendation(earrings, order_set(earrings, ungrouped=True), ORDER_MODE_PEARLS).quantity == 3
    assert build_order_recommendation(ring, order_set(ring, ungrouped=True), ORDER_MODE_PEARLS).quantity == 2
    assert build_order_recommendation(pendant, order_set(pendant, category=CATEGORY_ZERO, ungrouped=True), ORDER_MODE_PEARLS).quantity == 2


def test_incoming_earrings_block_self_order_but_force_ring_balance_inside_real_set():
    earrings = item("ER-TVP5", group="Earrings", tvp=5, stock=0, sales=0)
    ring = item("RG-ONE", group="Ring", stock=1, sales=0, tt=1, row=11)
    current_set = order_set(earrings, ring, category=CATEGORY_ZERO)

    earrings_rec = build_order_recommendation(earrings, current_set, ORDER_MODE_PEARLS)
    ring_rec = build_order_recommendation(ring, current_set, ORDER_MODE_PEARLS)

    assert earrings_rec.quantity == 0
    assert earrings_rec.blocked_by_tvp is True
    assert ring_rec.quantity == 3
    assert any("серьги находятся в пути" in reason.lower() for reason in ring_rec.reasons)


def test_ten_incoming_earrings_require_six_rings_only_when_missing():
    earrings = item("ER-TVP10", group="Earrings", tvp=10)
    no_rings = item("RG-ZERO", group="Ring", stock=0, row=11)
    six_rings = item("RG-SIX", group="Ring", stock=6, tt=1, row=11)

    assert build_order_recommendation(no_rings, order_set(earrings, no_rings, category=CATEGORY_ZERO), ORDER_MODE_PEARLS).quantity == 6
    assert build_order_recommendation(six_rings, order_set(earrings, six_rings, category=CATEGORY_ZERO), ORDER_MODE_PEARLS).quantity == 0


def test_incoming_rings_can_force_earrings_only_inside_real_set():
    rings = item("RG-TVP4", group="Ring", tvp=4)
    earrings = item("ER-MISSING", group="Earrings", stock=0, tt=1, row=11)
    current_set = order_set(rings, earrings, category=CATEGORY_ZERO)
    assert build_order_recommendation(earrings, current_set, ORDER_MODE_PEARLS).quantity == 5

    standalone_rings = item("RG-SOLO-TVP", group="Ring", tvp=4, ungrouped=True, set_id="Без комплекта")
    standalone_earrings = item(
        "ER-SOLO-ENOUGH", group="Earrings", sales=0, stock=5, tt=1,
        row=11, ungrouped=True, set_id="Без комплекта",
    )
    standalone_set = order_set(standalone_rings, standalone_earrings, category=CATEGORY_ZERO, ungrouped=True)
    assert build_order_recommendation(standalone_earrings, standalone_set, ORDER_MODE_PEARLS).quantity == 0
