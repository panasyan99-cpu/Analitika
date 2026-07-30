from pathlib import Path


def test_warehouse_sidebar_navigation_scrolls_full_page_sections():
    text = Path("src/warehouse.py").read_text(encoding="utf-8")
    assert "WAREHOUSE_SECTIONS" in text
    assert "def warehouse_navigation_items" in text
    assert "def render_navigation(" in text
    assert 'kind="anchor"' in text
    assert 'href=f"#{anchor_name}"' in text
    assert "render_mobile_navigation(items)" in text
    assert 'key="warehouse_section"' not in text
    assert '"Раздел склада"' not in text


def test_warehouse_renders_every_section_in_one_page():
    text = Path("src/warehouse.py").read_text(encoding="utf-8")
    expected = [
        '"warehouse-overview"',
        "render_overview(bundle)",
        '"warehouse-souvenirs"',
        "render_inventory_section(bundle.souvenirs",
        '"warehouse-components"',
        "render_inventory_section(bundle.components",
        '"warehouse-attention"',
        "render_attention(bundle)",
        '"warehouse-movement"',
        "render_movement(bundle.operations)",
        '"warehouse-supplies"',
        "render_supplies(bundle.supplies)",
    ]
    positions = [text.index(token, text.index("def render_warehouse_dashboard")) for token in expected]
    assert positions == sorted(positions)


def test_warehouse_has_phone_and_tablet_breakpoints():
    text = Path("src/warehouse.py").read_text(encoding="utf-8")
    assert '@media (max-width:900px)' in text
    assert '@media (max-width:640px)' in text
