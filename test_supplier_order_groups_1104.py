from __future__ import annotations

from src.order_workflow import (
    ORDER_MODE_PEARLS,
    ORDER_MODE_STONES,
    GREEN_STONES_GROUP,
    OTHER_STONES_GROUP,
    PEARL_ORDER_BUCKET_ORDER,
    STONE_ORDER_BUCKET_ORDER,
    OrderItem,
    build_order_sets,
    canonical_stone,
    item_in_mode,
    order_navigation_options,
    order_set_navigation_bucket,
    order_stone_bucket,
    pearl_order_bucket,
)


def _item(stone: str, *, sku: str = "SKU", group: str = "Pendant", set_id: str = "Set# 1") -> OrderItem:
    return OrderItem(
        row=11,
        set_id=set_id,
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


def test_stone_order_has_exactly_six_top_level_sections() -> None:
    assert order_navigation_options(ORDER_MODE_STONES) == (
        "Sapphire",
        "Ruby",
        "Moissanite",
        "Topaz",
        "Green Stones",
        "Other Stones",
    )
    assert order_navigation_options(ORDER_MODE_STONES) == STONE_ORDER_BUCKET_ORDER


def test_sapphire_ruby_and_moissanite_have_single_business_buckets() -> None:
    assert order_stone_bucket("Blue Sapphire") == "Sapphire"
    assert order_stone_bucket("Blue Sapphire High Quality") == "Sapphire"
    assert order_stone_bucket("Blue Sapphire Medium Quality") == "Sapphire"
    assert order_stone_bucket("Ruby") == "Ruby"
    assert order_stone_bucket("Moissanite") == "Moissanite"


def test_all_supported_topaz_variants_share_topaz_bucket() -> None:
    values = (
        "London Topaz",
        "Swiss Topaz",
        "Azure Topaz",
        "White Topaz",
        "Blue Topaz",
        "Sky Blue Topaz",
        "Multi Blue Topaz",
    )
    for value in values:
        assert order_stone_bucket(value) == "Topaz"
    assert canonical_stone("Azure Topaz") == "Azure Topaz"


def test_business_green_stones_share_one_bucket() -> None:
    values = (
        "Emerald",
        "Created Emerald",
        "Red Emerald",
        "Rhombium",
        "Chrome Diopside",
        "Garnet",
        "Peridot",
    )
    for value in values:
        assert order_stone_bucket(value) == GREEN_STONES_GROUP


def test_every_other_material_is_kept_in_other_stones_pool() -> None:
    values = ("Amethyst", "MOP", "MOR", "AMA", "Unknown supplier material", "")
    for value in values:
        assert order_stone_bucket(value) == OTHER_STONES_GROUP


def test_pearl_order_has_exactly_five_sections() -> None:
    assert order_navigation_options(ORDER_MODE_PEARLS) == ("White", "Grey", "Pink", "Black", "Baroque")
    assert order_navigation_options(ORDER_MODE_PEARLS) == PEARL_ORDER_BUCKET_ORDER


def test_pearl_colours_and_baroque_are_kept_separate() -> None:
    assert pearl_order_bucket("Freshwater Pearl White") == "White"
    assert pearl_order_bucket("Freshwater Pearl Grey") == "Grey"
    assert pearl_order_bucket("Freshwater Pearl Gray") == "Grey"
    assert pearl_order_bucket("Freshwater Pearl Pink") == "Pink"
    assert pearl_order_bucket("Freshwater Pearl Rose") == "Pink"
    assert pearl_order_bucket("Freshwater Pearl Black") == "Black"
    assert pearl_order_bucket("Freshwater Baroque Pearl Pink") == "Baroque"


def test_sea_round_and_unresolved_pearls_do_not_enter_order() -> None:
    values = (
        "South Sea Pearl",
        "Akoya Pearl",
        "Freshwater Pearl Round White",
        "Freshwater Pearl Round Pink",
        "Colored Freshwater Pearl",
        "Unknown Pearl",
    )
    for value in values:
        item = _item(value)
        assert pearl_order_bucket(value) is None
        assert item_in_mode(item, ORDER_MODE_PEARLS) is False


def test_pearl_mode_uses_pearl_bucket_not_secondary_stone_name() -> None:
    item = _item("Freshwater Pearl Pink", sku="PINK-PEARL-BS")
    order_set = build_order_sets((item,), ORDER_MODE_PEARLS)[0]
    assert order_set_navigation_bucket(order_set, ORDER_MODE_PEARLS) == "Pink"


def test_secondary_stone_code_does_not_replace_pearl_set_identity() -> None:
    item = _item("Freshwater Pearl White", sku="SKE21A001-BS")
    order_set = build_order_sets((item,), ORDER_MODE_PEARLS)[0]
    assert order_set.stone == "White"
    assert order_set_navigation_bucket(order_set, ORDER_MODE_PEARLS) == "White"
