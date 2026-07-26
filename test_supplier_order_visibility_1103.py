from __future__ import annotations

from src.order_workflow import (
    ORDER_MODE_PEARLS,
    ORDER_MODE_STONES,
    OTHER_STONES_GROUP,
    OrderItem,
    item_in_mode,
    order_stone_bucket,
)


def _item(stone: str, *, sku: str = "SKU", group: str = "Pendant") -> OrderItem:
    return OrderItem(
        row=11,
        set_id="Set# 1",
        sku=sku,
        stone=stone,
        group=group,
        sales=1,
        stock_63=0,
        stock_20=0,
        stores={},
        total_stock=0,
        working_stock=0,
        ntr2_stock=0,
        ntr2_calculated=False,
        tvp_raw=0,
    )


def test_every_round_pearl_is_hidden_from_supplier_order() -> None:
    values = (
        "Freshwater Pearl Round White",
        "Freshwater Pearl Round Pink",
        "Freshwater Pearl Round Rose",
        "Freshwater Pearl Round Grey",
        "Freshwater Pearl Round Gray",
        "Freshwater Pearl Round Black",
        "Round Freshwater Pearl",
    )
    for value in values:
        item = _item(value, sku=f"PEARL-{value}")
        assert item_in_mode(item, ORDER_MODE_PEARLS) is False
        assert item_in_mode(item, ORDER_MODE_STONES) is False


def test_round_marker_in_sku_also_hides_pearl() -> None:
    item = _item("Colored Freshwater Pearl", sku="SKP-ROUND-PEARL-01")
    assert item_in_mode(item, ORDER_MODE_PEARLS) is False


def test_non_round_white_and_colored_pearls_remain_available() -> None:
    assert item_in_mode(_item("Freshwater Pearl White"), ORDER_MODE_PEARLS) is True
    assert item_in_mode(_item("Freshwater Pearl Pink"), ORDER_MODE_PEARLS) is True
    assert item_in_mode(_item("Freshwater Pearl Grey"), ORDER_MODE_PEARLS) is True
    assert item_in_mode(_item("Freshwater Pearl Black"), ORDER_MODE_PEARLS) is True


def test_rare_abbreviated_and_unknown_stones_use_other_stones_bucket() -> None:
    assert order_stone_bucket("MOP") == OTHER_STONES_GROUP
    assert order_stone_bucket("MOR") == OTHER_STONES_GROUP
    assert order_stone_bucket("AMA") == OTHER_STONES_GROUP
    assert order_stone_bucket("Unknown supplier material") == OTHER_STONES_GROUP
    assert order_stone_bucket("") == OTHER_STONES_GROUP


def test_established_top_stones_use_business_navigation_entries() -> None:
    assert order_stone_bucket("Blue Sapphire") == "Sapphire"
    assert order_stone_bucket("Ruby") == "Ruby"
    assert order_stone_bucket("Moissanite") == "Moissanite"
    assert order_stone_bucket("London Topaz") == "Topaz"
    assert order_stone_bucket("Swiss Topaz") == "Topaz"
