from pathlib import Path

from src.sonu import sonu_navigation_items
from src.warehouse import WAREHOUSE_SECTIONS, warehouse_navigation_items


ROOT = Path(__file__).resolve().parent


def test_warehouse_navigation_is_anchor_based():
    items = warehouse_navigation_items()
    warehouse_items = items[1: 1 + len(WAREHOUSE_SECTIONS)]
    assert all(item.kind == "anchor" for item in warehouse_items)
    assert [item.href for item in warehouse_items] == [f"#{anchor}" for _, anchor, _ in WAREHOUSE_SECTIONS]


def test_warehouse_has_no_central_section_switcher():
    source = (ROOT / "src" / "warehouse.py").read_text(encoding="utf-8")
    assert 'st.segmented_control(\n        "Раздел склада"' not in source
    assert "render_overview(bundle)" in source
    assert "render_inventory_section(bundle.souvenirs" in source
    assert "render_inventory_section(bundle.components" in source
    assert "render_attention(bundle)" in source
    assert "render_movement(bundle.operations)" in source
    assert "render_supplies(bundle.supplies)" in source


def test_mobile_fx_stacks_without_overlap():
    source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    assert "flex-direction:column !important" in source
    assert ".st-key-global_fx_compact [data-testid=\"stNumberInput\"]" in source
    assert "width:100% !important" in source


def test_sonu_uses_stone_type_terminology():
    source = (ROOT / "src" / "sonu.py").read_text(encoding="utf-8")
    assert "Виды камней" in source
    assert "участниками камней" not in source
    assert "Участники группы" not in source
    assert sonu_navigation_items(False)[0].label == "Загрузка отчета"


def test_release_version_is_140():
    version = (ROOT / "version.json").read_text(encoding="utf-8")
    changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    assert '"version": "1.7.1"' in version
    assert "## 1.6.0 — Metal filters and Sonu AI order report" in changelog
