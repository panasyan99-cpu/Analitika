from pathlib import Path

import pandas as pd

import src.sonu as sonu


def _frame() -> pd.DataFrame:
    return pd.DataFrame([
        {
            "SKU": "UNKNOWN-A",
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
            "SKU": "UNKNOWN-B",
            "Категория": "Bracelet",
            "Категория RU": "Браслеты",
            "Камень": "Moissanite",
            "Магазин": "SCR",
            "Скорость продаж": 3,
            "Продажи VND": 3_000_000,
            "Остаток сети": 7,
            "Проба": "B 925",
            "Фото": b"image-b",
        },
    ])


def test_manual_decision_has_priority_over_automatic_5050(tmp_path, monkeypatch):
    override_file = tmp_path / "bracelet_classification_overrides.json"
    monkeypatch.setattr(sonu, "BRACELET_OVERRIDE_FILE", override_file)

    automatic = sonu.classify_bracelets(_frame(), 25_000, 30)
    assert set(automatic["Источник классификации"]) == {"Правило 50/50"}

    saved, persisted, _ = sonu.save_bracelet_overrides(
        {"UNKNOWN-A": sonu.FULL_CIRCLE_BRACELET_LABEL}
    )
    assert persisted is True
    assert saved["UNKNOWN-A"] == sonu.FULL_CIRCLE_BRACELET_LABEL

    classified = sonu.classify_bracelets(_frame(), 25_000, 30)
    selected = classified.set_index("SKU").loc["UNKNOWN-A"]
    assert selected["Тип браслета"] == sonu.FULL_CIRCLE_BRACELET_LABEL
    assert selected["Источник классификации"] == "Ручной выбор"

    _, audit = sonu.bracelet_classification_audit(_frame(), 25_000, 30)
    audited = audit.set_index("SKU").loc["UNKNOWN-A"]
    assert audited["Статус классификации"] == "Разобрано вручную"


def test_review_rows_keep_photos_for_large_manual_card(tmp_path, monkeypatch):
    monkeypatch.setattr(
        sonu,
        "BRACELET_OVERRIDE_FILE",
        tmp_path / "bracelet_classification_overrides.json",
    )
    rows = sonu._bracelet_review_rows(_frame(), 25_000, 30, mode="pending")
    assert len(rows) == 2
    assert rows["Фото"].notna().all()


def test_override_backup_can_be_imported(tmp_path, monkeypatch):
    first = tmp_path / "first.json"
    monkeypatch.setattr(sonu, "BRACELET_OVERRIDE_FILE", first)
    sonu.save_bracelet_overrides({"UNKNOWN-A": sonu.CENTERED_BRACELET_LABEL})
    backup = sonu.bracelet_overrides_json()

    second = tmp_path / "second.json"
    monkeypatch.setattr(sonu, "BRACELET_OVERRIDE_FILE", second)
    imported, persisted, _ = sonu.import_bracelet_overrides(backup)
    assert persisted is True
    assert imported["UNKNOWN-A"] == sonu.CENTERED_BRACELET_LABEL
    assert second.exists()


def test_ui_uses_guided_review_instead_of_ambiguous_table():
    source = Path(sonu.__file__).read_text(encoding="utf-8")
    assert 'Разобрать спорные модели ·' in source
    assert '@st.dialog("Разобрать спорные модели"' in source
    assert 'sonu_ambiguous_bracelets' not in source
    assert 'Сохранить решения (' in source


def test_guided_dialog_renders_without_exception(tmp_path):
    app = tmp_path / "dialog_app.py"
    app.write_text(
        '''
import io
import pandas as pd
import streamlit as st
from PIL import Image
from src.sonu import (
    BRACELET_REVIEW_MODE_KEY,
    BRACELET_REVIEW_OPEN_KEY,
    _render_bracelet_classification_audit,
)
buffer = io.BytesIO()
Image.new("RGB", (64, 64), "white").save(buffer, format="PNG")
image = buffer.getvalue()
frame = pd.DataFrame([
    {"SKU":"UNKNOWN-A","Категория":"Bracelet","Категория RU":"Браслеты","Камень":"Ruby","Магазин":"TT","Скорость продаж":8,"Продажи VND":8000000,"Остаток сети":1,"Проба":"B 925","Фото":image},
    {"SKU":"UNKNOWN-B","Категория":"Bracelet","Категория RU":"Браслеты","Камень":"Moissanite","Магазин":"TT","Скорость продаж":3,"Продажи VND":3000000,"Остаток сети":6,"Проба":"B 925","Фото":image},
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
    assert "Сохранить решения (0/2)" in labels
