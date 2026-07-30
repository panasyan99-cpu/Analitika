from pathlib import Path

from src.sonu import sonu_navigation_items
from src.warehouse import WAREHOUSE_SECTIONS, warehouse_navigation_items


ROOT = Path(__file__).resolve().parent


def test_warehouse_uses_lazy_internal_navigation():
    core = (ROOT / "src" / "warehouse.py").read_text(encoding="utf-8")
    ui = (ROOT / "src" / "warehouse_management" / "ui.py").read_text(encoding="utf-8")
    assert "render_warehouse_workspace(config, selected_metal_groups)" in core
    assert '"Раздел склада"' in ui
    assert '"Главная"' in ui
    assert '"Каталог"' in ui
    assert '"Поставки"' in ui
    assert '"Новая поставка"' in ui
    assert '"Приёмка"' in ui
    assert '"Передача в бухгалтерию"' in ui
    assert '"Операции"' in ui
    assert 'HISTORY_WORKSPACES = ("Операции",)' in ui


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
    assert '"version": "2.5.6"' in version
    assert "## 1.6.0 — Metal filters and Sonu AI order report" in changelog
