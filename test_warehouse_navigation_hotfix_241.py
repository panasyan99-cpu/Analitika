from pathlib import Path

from src.warehouse_management.ui import (
    _apply_pending_widget_state,
    _queue_widget_state,
)


def test_pending_navigation_is_applied_before_widget_creation() -> None:
    state: dict[str, object] = {}
    _queue_widget_state(
        state=state,
        warehouse_workspace="Поставки",
        warehouse_supply_workspace="Приёмка",
    )
    assert "warehouse_workspace" not in state
    _apply_pending_widget_state(state=state)
    assert state["warehouse_workspace"] == "Поставки"
    assert state["warehouse_supply_workspace"] == "Приёмка"
    assert "_warehouse_pending_widget_state" not in state


def test_invalid_pending_widget_value_is_ignored() -> None:
    state: dict[str, object] = {}
    _queue_widget_state(state=state, warehouse_workspace="Несуществующий раздел")
    _apply_pending_widget_state(state=state)
    assert "warehouse_workspace" not in state


def test_safe_schema_redirect_does_not_mutate_live_widget_key_directly() -> None:
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    assert 'st.session_state["warehouse_workspace"] = "История"' not in ui
    assert "_auto_prepare_safe_schema" in ui
    assert "_apply_pending_widget_state()" in ui


def test_catalog_and_supply_redirects_use_pending_state() -> None:
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    assert '_queue_widget_state(' in ui
    assert 'warehouse_catalog_mode="Управление"' in ui
