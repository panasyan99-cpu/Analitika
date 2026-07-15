from pathlib import Path


def test_warehouse_sidebar_navigation_switches_lazy_sections():
    text = Path("src/warehouse.py").read_text(encoding="utf-8")
    assert "WAREHOUSE_SECTIONS" in text
    assert "def warehouse_navigation_items" in text
    assert "def render_navigation(" in text
    assert 'kind="button"' in text
    assert 'items=warehouse_navigation_items(current_section)' in text
    assert 'st.session_state["warehouse_section"] = selected' in text
    assert 'st.segmented_control(' in text
    assert 'key="warehouse_section"' in text


def test_warehouse_has_phone_and_tablet_breakpoints():
    text = Path("src/warehouse.py").read_text(encoding="utf-8")
    assert '@media (max-width:900px)' in text
    assert '@media (max-width:640px)' in text
