from pathlib import Path


def test_private_baserow_photos_are_converted_to_data_uri():
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    assert "def _remote_thumbnail_data_uri" in ui
    assert "fetch_image_bytes(url, token)" in ui
    assert "data:image/jpeg;base64" in ui
    assert "_item_photo_data_uri" in ui
    assert "_row_photo_data_uri" in ui


def test_warehouse_ux_groups_daily_tasks():
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    assert 'SUPPLY_WORKSPACES = ("Реестр", "Новая поставка", "Приёмка")' in ui
    assert 'HISTORY_WORKSPACES = ("Операции",)' in ui
    assert "def _workflow" in ui
    assert '"Добавить поставку"' in ui
    assert '"Перейти к приёмке"' in ui
    assert '"Передать товар"' in ui


def test_photos_exist_in_all_operational_tables():
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    for function in ("render_catalog", "render_supplies", "render_new_supply", "render_receiving", "render_transfer"):
        body = ui[ui.index(f"def {function}"): ]
        assert '"Фото"' in body
    assert 'st.column_config.ImageColumn("Фото"' in ui
    assert "row_height=92" in ui
