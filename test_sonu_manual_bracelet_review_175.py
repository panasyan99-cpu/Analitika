from pathlib import Path

import pandas as pd

import src.sonu as sonu


def _frame() -> pd.DataFrame:
    return pd.DataFrame([
        {
            "SKU": "UNKNOWN-FAMILY-RUBY",
            "Категория": "Bracelet",
            "Категория RU": "Браслеты",
            "Камень": "Ruby",
            "Магазин": "TT",
            "Скорость продаж": 8,
            "Продажи VND": 8_000_000,
            "Остаток сети": 1,
            "Проба": "B 925",
            "Фото": b"image-a",
        },
        {
            "SKU": "UNKNOWN-FAMILY-BS",
            "Категория": "Bracelet",
            "Категория RU": "Браслеты",
            "Камень": "Blue Sapphire",
            "Магазин": "SCR",
            "Скорость продаж": 3,
            "Продажи VND": 3_000_000,
            "Остаток сети": 7,
            "Проба": "B 925",
            "Фото": b"image-b",
        },
    ])


def test_family_decision_has_priority_and_applies_to_all_related_skus(tmp_path, monkeypatch):
    override_file = tmp_path / "bracelet_classification_overrides.json"
    monkeypatch.setattr(sonu, "BRACELET_OVERRIDE_FILE", override_file)
    monkeypatch.setattr(sonu, "BRACELET_CATALOG_FILE", tmp_path / "missing_catalog.json")

    automatic = sonu.classify_bracelets(_frame(), 25_000, 30)
    assert set(automatic["Источник классификации"]) == {"Требует классификации"}

    family = sonu._bracelet_model_key("UNKNOWN-FAMILY-RUBY")
    key = sonu._bracelet_family_override_key(family)
    saved, persisted, _ = sonu.save_bracelet_overrides(
        {key: sonu.FULL_CIRCLE_BRACELET_LABEL}
    )
    assert persisted is True
    assert saved[key] == sonu.FULL_CIRCLE_BRACELET_LABEL

    classified = sonu.classify_bracelets(_frame(), 25_000, 30)
    family_rows = classified[classified["Модельная семья"] == family]
    assert set(family_rows["Тип браслета"]) == {sonu.FULL_CIRCLE_BRACELET_LABEL}
    assert set(family_rows["Источник классификации"]) == {"Ручной выбор семейства"}


def test_review_rows_group_skus_by_model_family_and_keep_photo(tmp_path, monkeypatch):
    monkeypatch.setattr(sonu, "BRACELET_OVERRIDE_FILE", tmp_path / "overrides.json")
    monkeypatch.setattr(sonu, "BRACELET_CATALOG_FILE", tmp_path / "missing_catalog.json")
    rows = sonu._bracelet_review_rows(_frame(), 25_000, 30, mode="pending")
    assert len(rows) == 1
    assert rows.iloc[0]["Моделей в семье"] == 2
    assert rows["Фото"].notna().all()


def test_override_backup_can_store_and_import_family_decisions(tmp_path, monkeypatch):
    first = tmp_path / "first.json"
    monkeypatch.setattr(sonu, "BRACELET_OVERRIDE_FILE", first)
    key = sonu._bracelet_family_override_key("UNKNOWN-FAMILY")
    sonu.save_bracelet_overrides({key: sonu.CENTERED_BRACELET_LABEL})
    backup = sonu.bracelet_overrides_json()

    second = tmp_path / "second.json"
    monkeypatch.setattr(sonu, "BRACELET_OVERRIDE_FILE", second)
    imported, persisted, _ = sonu.import_bracelet_overrides(backup)
    assert persisted is True
    assert imported[key] == sonu.CENTERED_BRACELET_LABEL
    assert second.exists()


def test_ui_uses_family_review_and_pending_terminology():
    source = Path(sonu.__file__).read_text(encoding="utf-8")
    assert 'Разобрать модели ·' in source
    assert '@st.dialog("Разобрать модели, требующие классификации"' in source
    assert 'Разобрать спорные модели' not in source
    assert 'Сохранить решения (' in source


def test_guided_dialog_renders_without_exception(tmp_path, monkeypatch):
    monkeypatch.setattr(sonu, "BRACELET_CATALOG_FILE", tmp_path / "missing_catalog.json")
    app = tmp_path / "dialog_app.py"
    app.write_text(
        '''
import io
import pandas as pd
import streamlit as st
from PIL import Image
import src.sonu as sonu
from src.sonu import (
    BRACELET_REVIEW_MODE_KEY,
    BRACELET_REVIEW_OPEN_KEY,
    _render_bracelet_classification_audit,
)
sonu.BRACELET_CATALOG_FILE = __import__('pathlib').Path("/definitely/missing/catalog.json")
buffer = io.BytesIO()
Image.new("RGB", (64, 64), "white").save(buffer, format="PNG")
image = buffer.getvalue()
frame = pd.DataFrame([
    {"SKU":"UNKNOWN-FAMILY-RUBY","Категория":"Bracelet","Категория RU":"Браслеты","Камень":"Ruby","Магазин":"TT","Скорость продаж":8,"Продажи VND":8000000,"Остаток сети":1,"Проба":"B 925","Фото":image},
    {"SKU":"UNKNOWN-FAMILY-BS","Категория":"Bracelet","Категория RU":"Браслеты","Камень":"Blue Sapphire","Магазин":"TT","Скорость продаж":3,"Продажи VND":3000000,"Остаток сети":6,"Проба":"B 925","Фото":image},
])
st.session_state[BRACELET_REVIEW_OPEN_KEY] = True
st.session_state[BRACELET_REVIEW_MODE_KEY] = "pending"
_render_bracelet_classification_audit(frame, 25000, 30)
''',
        encoding="utf-8",
    )
    from streamlit.testing.v1 import AppTest

    app_test = AppTest.from_file(str(app), default_timeout=20).run()
    assert not app_test.exception
    labels = {button.label for button in app_test.button}
    assert "С затяжкой / центральная композиция" in labels
    assert "Без затяжки / полный круг" in labels
    assert "Сохранить решения (0/1)" in labels
