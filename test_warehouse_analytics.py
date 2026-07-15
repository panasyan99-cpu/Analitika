from pathlib import Path

from src.warehouse import (
    DEFAULT_MINIMUM_STOCK,
    EARLY_WARNING_STOCK,
    file_url,
    normalize_inventory_rows,
    normalize_operations,
    stock_status,
    text_value,
)


def test_stock_status_thresholds() -> None:
    assert DEFAULT_MINIMUM_STOCK == 10
    assert EARLY_WARNING_STOCK == 15
    assert stock_status(0, 10) == "Нет в наличии"
    assert stock_status(8, 10) == "Ниже минимума"
    assert stock_status(12, 10) == "Заканчивается"
    assert stock_status(16, 10) == "В наличии"


def test_baserow_values_and_photo_url() -> None:
    assert text_value([{"value": "Agate"}, {"value": "Labradorite"}]) == "Agate; Labradorite"
    assert file_url([{"url": "/media/test.jpg"}], "https://example.com") == "https://example.com/media/test.jpg"


def test_inventory_normalization() -> None:
    rows = [
        {
            "id": 1,
            "Артикул": "FXN-1",
            "Фото": [{"url": "https://example.com/photo.jpg"}],
            "Категория": {"value": "Necklace"},
            "Материал": [{"value": "Brass"}],
            "Камень": [{"value": "Agate"}, {"value": "Labradorite"}],
            "Остаток": 14,
            "Минимальный остаток": 10,
        }
    ]
    frame = normalize_inventory_rows(rows, section="Сувенирка", base_url="https://example.com")
    assert frame.loc[0, "Артикул"] == "FXN-1"
    assert frame.loc[0, "Камень"] == "Agate; Labradorite"
    assert frame.loc[0, "Статус"] == "Заканчивается"


def test_operation_direction() -> None:
    rows = [
        {"Тип операции": "Приход", "Количество": 12},
        {"Тип операции": "Передача в бухгалтерию", "Количество": 5},
    ]
    frame = normalize_operations(rows)
    assert set(frame["Изменение"].tolist()) == {12, -5}


def test_complete_rebuild_contains_warehouse_mode() -> None:
    root = Path(__file__).resolve().parent
    app = (root / "streamlit_app.py").read_text(encoding="utf-8")
    assert 'APP_VERSION = "1.2.1"' in app
    assert '"Сувениры и касты на складе"' in app
    assert "render_warehouse_dashboard" in app
    assert (root / "src" / "warehouse.py").exists()
    assert (root / ".streamlit" / "secrets.example.toml").exists()
