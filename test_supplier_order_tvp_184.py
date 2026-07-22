from src.order_workflow import (
    CATEGORY_TOP,
    OrderItem,
    OrderSet,
    filter_order_sets_by_tvp,
)


def _item(row: int, tvp: int, *, group: str = "Ring", sales: int = 5) -> OrderItem:
    return OrderItem(
        row=row,
        set_id="Короны BS",
        sku=f"SKU-{row}",
        stone="Blue Sapphire",
        group=group,
        sales=sales,
        stock_63=0,
        stock_20=0,
        stores={},
        total_stock=0,
        working_stock=0,
        ntr2_stock=0,
        ntr2_calculated=True,
        tvp_raw=tvp,
    )


def _set() -> OrderSet:
    items = (
        _item(1, 5, group="Earrings", sales=0),
        _item(2, 0, group="Ring", sales=5),
        _item(3, -1, group="Pendant", sales=1),
    )
    return OrderSet(
        key="Камни|Blue Sapphire|Короны BS",
        set_id="Короны BS",
        stone="Blue Sapphire",
        items=items,
        category=CATEGORY_TOP,
        driver_sku="SKU-2",
        max_sales=5,
        has_positive_tvp=True,
        has_negative_tvp=True,
    )


def test_positive_tvp_filter_keeps_full_matching_set_context():
    result = filter_order_sets_by_tvp((_set(),), True)
    assert len(result) == 1
    assert [item.tvp_raw for item in result[0].items] == [5, 0, -1]
    assert [item.group for item in result[0].items] == ["Earrings", "Ring", "Pendant"]
    assert result[0].set_id == "Короны BS"
    assert result[0].category == CATEGORY_TOP


def test_nonpositive_tvp_filter_also_keeps_full_matching_set_context():
    result = filter_order_sets_by_tvp((_set(),), False)
    assert len(result) == 1
    assert [item.tvp_raw for item in result[0].items] == [5, 0, -1]
    assert result[0].has_positive_tvp is True
    assert result[0].has_negative_tvp is True


def test_tvp_filter_excludes_sets_without_matching_rows():
    positive_only_set = OrderSet(
        **{**_set().__dict__, "items": (_item(10, 3),), "has_negative_tvp": False}
    )
    zero_only_set = OrderSet(
        **{**_set().__dict__, "key": "zero", "items": (_item(11, 0),), "has_positive_tvp": False}
    )
    assert filter_order_sets_by_tvp((positive_only_set,), False) == ()
    assert filter_order_sets_by_tvp((zero_only_set,), True) == ()
