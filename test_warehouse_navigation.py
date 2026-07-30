from pathlib import Path


def test_warehouse_uses_task_oriented_lazy_workspace():
    core = Path("src/warehouse.py").read_text(encoding="utf-8")
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    assert "render_warehouse_workspace(config, selected_metal_groups)" in core
    assert 'WORKSPACES = (' in ui
    assert '"Главная"' in ui
    assert '"Товары"' in ui
    assert '"Поставки"' in ui
    assert '"Передача"' in ui
    assert '"История"' in ui
    assert "def render_supply_hub" in ui
    assert "def render_history_hub" in ui


def test_only_selected_primary_workspace_is_rendered():
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    body = ui[ui.index("def render_warehouse_workspace"): ]
    order = [
        'if current == "Главная"',
        'elif current == "Товары"',
        'elif current == "Поставки"',
        'elif current == "Передача"',
    ]
    positions = [body.index(token) for token in order]
    assert positions == sorted(positions)
    assert "render_history_hub(config)" in body
    assert "Загружается только выбранный раздел" in body


def test_warehouse_has_phone_and_tablet_breakpoints():
    text = Path("src/warehouse.py").read_text(encoding="utf-8")
    management = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    assert '@media (max-width:900px)' in text
    assert '@media (max-width:640px)' in text
    assert '@media (max-width:900px)' in management
    assert '@media (max-width:640px)' in management
