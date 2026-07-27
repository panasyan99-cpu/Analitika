from pathlib import Path

ROOT = Path(__file__).parent
SOURCE = (ROOT / "src" / "order_workflow.py").read_text(encoding="utf-8")


def test_completed_historical_order_opens_on_overview_once():
    assert "supplier_order_open_completed_overview::" in SOURCE
    assert 'draft.stage = "order"' in SOURCE
    assert 'draft.status == "completed"' in SOURCE


def test_open_order_has_explicit_save_and_close_action():
    assert '"Сохранить и закрыть"' in SOURCE
    assert 'supplier_order_save_close_inline::' in SOURCE
    assert 'st.session_state["supplier_order_library_open"] = True' in SOURCE


def test_ring_stage_has_top_and_bottom_exit_controls():
    assert 'supplier_order_close_from_sizes::' in SOURCE
    assert 'supplier_order_close_from_sizes_bottom::' in SOURCE
    assert 'supplier_order_back_from_sizes::' in SOURCE
    assert 'supplier_order_back_from_sizes_bottom::' in SOURCE


def test_active_workspace_close_button_is_not_mislabeled_as_unfinished_orders():
    active_block = SOURCE[SOURCE.index("if active is not None:"):SOURCE.index("library_open = bool", SOURCE.index("if active is not None:"))]
    assert "← Закрыть заказ" in active_block
    assert "Незавершённые заказы" not in active_block
