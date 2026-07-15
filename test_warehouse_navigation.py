from pathlib import Path


def test_warehouse_sidebar_navigation_switches_lazy_sections():
    text = Path("src/warehouse.py").read_text(encoding="utf-8")
    assert "WAREHOUSE_SECTIONS" in text
    assert "def render_navigation(current_section: str)" in text
    assert 'st.session_state["warehouse_section"] = section' in text
    assert 'type="primary" if section == current_section else "secondary"' in text
    assert 'st.segmented_control(' in text
    assert 'key="warehouse_section"' in text


def test_warehouse_has_phone_and_tablet_breakpoints():
    text = Path("src/warehouse.py").read_text(encoding="utf-8")
    assert '@media (max-width:900px)' in text
    assert '@media (max-width:640px)' in text
