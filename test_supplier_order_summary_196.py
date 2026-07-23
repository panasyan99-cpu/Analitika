from __future__ import annotations

from src.order_workflow import OrderDraft, OrderItem, OrderSet, order_quantity_summary


def make_item(row: int, sku: str, group: str) -> OrderItem:
    return OrderItem(
        row=row,
        set_id="Set# summary",
        sku=sku,
        stone="Blue Sapphire",
        group=group,
        sales=0,
        stock_63=0,
        stock_20=0,
        stores={"TT": 0, "63": 0},
        total_stock=0,
        working_stock=0,
        ntr2_stock=0,
        ntr2_calculated=False,
        tvp_raw=0,
        stock_tt=0,
    )


def make_set(*items: OrderItem) -> OrderSet:
    return OrderSet(
        key="summary|set",
        set_id="Set# summary",
        stone="Blue Sapphire",
        items=tuple(items),
        category="Средние комплекты",
        driver_sku=items[0].sku,
        max_sales=0,
        has_positive_tvp=False,
        has_negative_tvp=False,
    )


def test_order_quantity_summary_counts_units_not_ring_skus():
    earrings = make_item(1, "ER-1", "Earrings")
    ring_a = make_item(2, "RG-1", "Ring")
    ring_b = make_item(3, "RG-2", "Ring")
    pendant = make_item(4, "PD-1", "Pendant")
    current_set = make_set(earrings, ring_a, ring_b, pendant)
    draft = OrderDraft(source_hash="hash", source_name="source.xlsx", mode="Камни")
    draft.orders = {
        earrings.key: 5,
        ring_a.key: 3,
        ring_b.key: 4,
        pendant.key: 3,
    }

    summary = order_quantity_summary((current_set,), draft)

    assert summary == {
        "earrings_qty": 5,
        "rings_qty": 7,
        "pendants_qty": 3,
        "other_qty": 0,
        "total_qty": 15,
        "sku_count": 4,
    }


def test_order_quantity_summary_excludes_limited_order_and_zero_rows():
    earrings = make_item(1, "ER-1", "Earrings")
    ring = make_item(2, "RG-1", "Ring")
    pendant = make_item(3, "PD-1", "Pendant")
    current_set = make_set(earrings, ring, pendant)
    draft = OrderDraft(source_hash="hash", source_name="source.xlsx", mode="Жемчуг")
    draft.orders = {earrings.key: 5, ring.key: 3, pendant.key: 0}
    draft.limited_orders = {ring.key: True}

    summary = order_quantity_summary((current_set,), draft)

    assert summary["earrings_qty"] == 5
    assert summary["rings_qty"] == 0
    assert summary["pendants_qty"] == 0
    assert summary["total_qty"] == 5
    assert summary["sku_count"] == 1
