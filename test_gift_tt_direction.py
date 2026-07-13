from src.report import StoreData

from streamlit_app import outlet_direction_frame, stores_fact_dataframe


def test_gift_tt_is_not_in_comparison_fact_segments():
    store = StoreData("OUTLET")
    store.add("PEARLS", "White Freshwater Pearl", "Necklace", 2, 1000.0, "White FW", "test")
    store.extras["GIFT TT"]["qty"] = 7
    store.extras["GIFT TT"]["amount"] = 3500.0

    facts = stores_fact_dataframe([store])

    assert "GIFT TT" not in set(facts["Сегмент"].astype(str))
    assert "GIFT TT" not in set(facts["Камень"].astype(str))
    assert set(facts["Сегмент"]) == {"Pearls"}


def test_gift_tt_has_one_separate_outlet_comparison_row():
    store = StoreData("OUTLET")
    store.extras["GIFT TT"]["qty"] = 7
    store.extras["GIFT TT"]["amount"] = 3500.0

    row = outlet_direction_frame(store)

    assert len(row) == 1
    assert row.iloc[0]["Направление"] == "GIFT TT"
    assert row.iloc[0]["Количество"] == 7
    assert row.iloc[0]["Выручка"] == 3500.0
    assert row.iloc[0]["Средняя стоимость"] == 500.0


def test_missing_period_still_keeps_gift_tt_row_for_outer_comparison():
    row = outlet_direction_frame(None)

    assert len(row) == 1
    assert row.iloc[0]["Направление"] == "GIFT TT"
    assert row.iloc[0]["Количество"] == 0
    assert row.iloc[0]["Выручка"] == 0
