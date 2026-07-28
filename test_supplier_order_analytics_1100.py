from __future__ import annotations

from pathlib import Path

import src.order_workflow as workflow
from src.order_workflow import (
    ORDER_MODE_PEARLS,
    ORDER_MODE_STONES,
    OrderDraft,
    OrderItem,
    ParsedOrderWorkbook,
    _workspace_status_label,
    build_order_analytics,
)


def make_item(row: int, sku: str, stone: str, group: str) -> OrderItem:
    return OrderItem(
        row=row,
        set_id=f"Set# {row}",
        sku=sku,
        stone=stone,
        group=group,
        sales=0,
        stock_63=0,
        stock_20=0,
        stores={},
        total_stock=0,
        working_stock=0,
        ntr2_stock=0,
        ntr2_calculated=False,
        tvp_raw=0,
        stock_tt=0,
    )


def make_parsed(items: tuple[OrderItem, ...]) -> ParsedOrderWorkbook:
    return ParsedOrderWorkbook(
        source_name="source.xlsx",
        source_hash="hash",
        upload_path="source.xlsx",
        period="",
        supplier="Y&J",
        store_columns=(),
        has_actual_ntr2=True,
        items=items,
    )


def families_by_name(section: dict) -> dict[str, dict]:
    return {row["name"]: row for row in section["families"]}


def test_stone_order_analytics_counts_pieces_and_hierarchies() -> None:
    sapphire_earrings = make_item(1, "ER-BS-1", "Blue Sapphire", "Earrings")
    sapphire_ring = make_item(2, "RG-BS", "Blue Sapphire", "Ring")
    ruby_pendant = make_item(3, "PD-RUBY", "Ruby", "Pendant")
    citrine_ring = make_item(4, "RG-CIT", "Citrine", "Ring")
    amethyst_earrings = make_item(5, "ER-AMST", "Amethyst", "Earrings")
    unknown_pendant = make_item(6, "PD-AMA", "AMA", "Pendant")
    limited_moissanite = make_item(7, "ER-MOIS", "Moissanite", "Earrings")
    parsed = make_parsed(
        (
            sapphire_earrings,
            sapphire_ring,
            ruby_pendant,
            citrine_ring,
            amethyst_earrings,
            unknown_pendant,
            limited_moissanite,
        )
    )
    draft = OrderDraft(source_hash="hash", source_name="source.xlsx", mode=ORDER_MODE_STONES)
    draft.orders = {
        sapphire_earrings.key: 10,
        sapphire_ring.key: 4,
        ruby_pendant.key: 3,
        citrine_ring.key: 2,
        amethyst_earrings.key: 5,
        unknown_pendant.key: 1,
        limited_moissanite.key: 8,
    }
    draft.limited_orders = {limited_moissanite.key: True}

    analytics = build_order_analytics(parsed, draft, ORDER_MODE_STONES)

    assert analytics["total_quantity"] == 25
    assert analytics["sku_count"] == 6
    assert analytics["limited_positions"] == 1
    assert analytics["group_totals"] == {"Earrings": 15, "Ring": 6, "Pendant": 4}

    sections = {section["name"]: section for section in analytics["sections"]}
    assert sections["Топовые камни"]["total_quantity"] == 17
    assert sections["Цветные камни"]["total_quantity"] == 8

    top = families_by_name(sections["Топовые камни"])
    assert top["Blue Sapphire — все вариации"]["total_quantity"] == 14
    assert top["Ruby"]["total_quantity"] == 3

    colored = families_by_name(sections["Цветные камни"])
    assert colored["Quartz Group"]["total_quantity"] == 2
    assert colored["Other Colored Stones"]["total_quantity"] == 5
    assert colored["Unrecognized"]["total_quantity"] == 1


def test_pearl_order_analytics_splits_white_and_colored_families() -> None:
    white = make_item(1, "ER-FPW", "Freshwater Pearl White", "Earrings")
    colored = make_item(2, "RG-FPC", "Freshwater Pearl Colored", "Ring")
    baroque = make_item(3, "PD-BAROQUE", "Freshwater Baroque Pearl", "Pendant")
    parsed = make_parsed((white, colored, baroque))
    draft = OrderDraft(source_hash="hash", source_name="source.xlsx", mode=ORDER_MODE_PEARLS)
    draft.orders = {white.key: 6, colored.key: 4, baroque.key: 2}

    analytics = build_order_analytics(parsed, draft, ORDER_MODE_PEARLS)

    assert analytics["total_quantity"] == 12
    assert analytics["group_totals"] == {"Earrings": 6, "Ring": 4, "Pendant": 2}
    families = families_by_name(analytics["sections"][0])
    assert families["White Freshwater"]["total_quantity"] == 6
    assert "Round White Freshwater" not in families
    assert families["Colored Freshwater"]["total_quantity"] == 4
    assert families["Baroque Pearls"]["total_quantity"] == 2


def test_pearl_order_analytics_keeps_pink_grey_and_black_in_colored() -> None:
    rose = make_item(1, "ER-FPROSE", "Freshwater Pearl Rose", "Earrings")
    round_pink = make_item(2, "ER-FPRP", "Freshwater Pearl Round Pink", "Earrings")
    grey = make_item(3, "RG-FPGR", "Freshwater Pearl Grey", "Ring")
    round_gray = make_item(4, "RG-FPRG", "Freshwater Pearl Round Gray", "Ring")
    black = make_item(5, "PD-FPB", "Freshwater Pearl Black", "Pendant")
    round_black = make_item(6, "PD-FPRB", "Freshwater Pearl Round Black", "Pendant")
    parsed = make_parsed((rose, round_pink, grey, round_gray, black, round_black))
    draft = OrderDraft(source_hash="hash", source_name="source.xlsx", mode=ORDER_MODE_PEARLS)
    draft.orders = {item.key: 1 for item in parsed.items}

    analytics = build_order_analytics(parsed, draft, ORDER_MODE_PEARLS)

    families = families_by_name(analytics["sections"][0])
    assert families["Colored Freshwater"]["total_quantity"] == 3
    assert "White Freshwater" not in families
    assert "Round White Freshwater" not in families

    # Round pearls are normally excluded from this supplier order, but the
    # classifier still must keep coloured round names out of the white family.
    assert workflow._pearl_analytics_family("Freshwater Pearl Round White") == "White Freshwater"
    assert workflow._pearl_analytics_family("Freshwater Pearl Round Pink") == "Colored Freshwater"
    assert workflow._pearl_analytics_family("Freshwater Pearl Round Rose") == "Colored Freshwater"
    assert workflow._pearl_analytics_family("Freshwater Pearl Round Gray") == "Colored Freshwater"
    assert workflow._pearl_analytics_family("Freshwater Pearl Round Black") == "Colored Freshwater"


def test_order_library_has_separate_completion_and_information_controls() -> None:
    source = Path(workflow.__file__).read_text(encoding="utf-8")
    assert '"Информация по заказу"' in source
    assert '"Завершить заказ по камням"' in source
    assert '"Завершить заказ по жемчугу"' in source
    assert '"Начать заказ"' in source
    assert "build_order_analytics(parsed, draft, mode)" in source
    assert "include_completed=True" in source
    assert "Показать завершённые" not in source


def test_workspace_status_keeps_two_order_states_visible() -> None:
    label = _workspace_status_label({ORDER_MODE_STONES: {"status": "completed"}})
    assert label == "Камни: завершён · Жемчуг: не начат"
