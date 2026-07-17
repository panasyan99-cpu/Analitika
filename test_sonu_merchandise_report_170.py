import pandas as pd
from src.sonu import sonu_merchandise_tables, sonu_order_recommendations, SONU_MAIN_COLUMNS


def sample_frame():
    return pd.DataFrame([
        {"SKU":"E1-BS","Категория":"Earrings","Категория RU":"Серьги","Камень":"SAPPHIRE","Проба":"925","Магазин":"63","Остаток сети":2,"Скорость продаж":8,"Продажи VND":8000000},
        {"SKU":"E1-BS","Категория":"Earrings","Категория RU":"Серьги","Камень":"SAPPHIRE","Проба":"925","Магазин":"SCR","Остаток сети":2,"Скорость продаж":2,"Продажи VND":2000000},
        {"SKU":"R1-RUBY","Категория":"Ring","Категория RU":"Кольца","Камень":"RUBY","Проба":"925","Магазин":"63","Остаток сети":10,"Скорость продаж":1,"Продажи VND":1000000},
    ])


def test_five_tables_and_network_stock_once():
    tables = sonu_merchandise_tables(sample_frame(), 1000000, 30)
    assert list(tables) == ["Серьги", "Кольца", "Подвески", "Браслеты не полный круг", "Браслеты полный круг"]
    earrings = tables["Серьги"]
    assert list(earrings.columns) == SONU_MAIN_COLUMNS
    assert earrings.iloc[0]["Продано изделий"] == 10
    assert earrings.iloc[0]["Остаток, шт."] == 2


def test_recommendations_have_no_day_horizons():
    table = pd.DataFrame([{"Камень":"Sapphire","Кол-во уникальных SKU":5,"Продано изделий":20,"Общий Total продаж, USD":1000,"SKU на остатке":1,"Остаток, шт.":2}])
    rec = sonu_order_recommendations(table)
    assert rec.iloc[0]["Приоритет"] == "Очень нужно заказать"
    assert not any("дней" in col.lower() for col in rec.columns)
