from datetime import datetime

import pandas as pd

from src.warehouse import category_chart, movement_chart, stone_sku_chart


def _inventory() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Категория": ["Ожерелья", "Браслеты"],
            "Остаток": [3591, 1650],
            "Камень": ["Agate; Labradorite", "Agate"],
        }
    )


def test_horizontal_bar_labels_have_reserved_space() -> None:
    category = category_chart(_inventory(), "Категории")
    maximum = max(category.data[0].x)
    assert category.data[0].cliponaxis is False
    assert category.layout.xaxis.range[1] > maximum
    assert category.layout.margin.r >= 70

    stone = stone_sku_chart(_inventory())
    maximum_stone = max(stone.data[0].x)
    assert stone.data[0].cliponaxis is False
    assert stone.layout.xaxis.range[1] > maximum_stone


def test_single_day_movement_uses_narrow_date_bars() -> None:
    operations = pd.DataFrame(
        {
            "Дата": [pd.Timestamp(datetime.now()).normalize()],
            "Изменение": [6400],
        }
    )
    figure = movement_chart(operations, 30)
    assert figure.layout.height == 330
    assert figure.layout.xaxis.range is not None
    assert figure.data[0].width == 9 * 60 * 60 * 1000
    assert figure.data[0].cliponaxis is False
