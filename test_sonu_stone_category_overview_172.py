from pathlib import Path
import pandas as pd
from src.sonu import cached_parse_sonu, sonu_stone_category_overview, _period_days


def test_stone_category_overview_has_requested_columns():
    path = Path('/mnt/data/sonu(2).xlsx')
    if not path.exists():
        return
    report = cached_parse_sonu(path.read_bytes())
    table = sonu_stone_category_overview(report.data, 26300, _period_days(report.period))
    expected = {
        'Группа камней', 'Камень', 'Номенклатурная группа',
        'Продано уникальных SKU', 'Продано изделий', 'Продажи, USD',
        'Средняя цена изделия, USD', 'SKU на остатке', 'Остаток, шт.'
    }
    assert expected.issubset(table.columns)
    assert not table.empty
    assert table['Номенклатурная группа'].nunique() >= 3
