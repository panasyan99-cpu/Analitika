from pathlib import Path


def test_transfer_modes_are_lazy_and_do_not_use_streamlit_tabs() -> None:
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    start = ui.index("def render_transfer")
    end = ui.index("def render_operations", start)
    body = ui[start:end]
    assert "st.tabs(" not in body
    assert 'key="warehouse_transfer_mode"' in body
    assert 'if mode == "По поставке — рекомендуемый"' in body
    assert 'if mode == "По отдельным SKU"' in body


def test_quantity_editor_has_per_row_hard_maximum_and_visible_edit_hint() -> None:
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    start = ui.index("def _render_quantity_editor")
    end = ui.index("def _render_catalog_cards", start)
    body = ui[start:end]
    assert 'kwargs["max_value"] = int(maximum)' in body
    assert "Золотые поля справа редактируются" in body
    assert 'quantity_label="Передать, шт."' in ui
    assert 'quantity_label="Принять сейчас, шт."' in ui


def test_operational_photos_are_large_and_paginated() -> None:
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    assert "WAREHOUSE_PHOTO_SIZE = 320" in ui
    assert "WAREHOUSE_TABLE_ROW_HEIGHT = 138" in ui
    assert 'st.image(data_uri, width=190)' in ui
    assert "def _page_slice" in ui
    assert "Фото загружаются только для этой страницы" in ui


def test_safe_schema_transfer_skips_redundant_operation_scan() -> None:
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    assert "already = {} if service.has_supply_lines else service.transferred_by_supply" in ui


def test_baserow_client_caches_rows_within_one_render() -> None:
    client = Path("src/warehouse_management/client.py").read_text(encoding="utf-8")
    assert "self._row_cache" in client
    assert "if not refresh and cache_key in self._row_cache" in client
