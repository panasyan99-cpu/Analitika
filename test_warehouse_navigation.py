from pathlib import Path


def test_warehouse_uses_one_lazy_internal_workspace():
    core = Path("src/warehouse.py").read_text(encoding="utf-8")
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    assert "render_warehouse_workspace(config, selected_metal_groups)" in core
    assert "WORKSPACES" in ui
    assert '"Раздел склада"' in ui
    assert "if current == \"Обзор\"" in ui
    assert "elif current == \"Каталог\"" in ui
    assert "elif current == \"Новая поставка\"" in ui
    assert "elif current == \"Приёмка\"" in ui
    assert "elif current == \"Передача в бухгалтерию\"" in ui
    assert "elif current == \"Операции\"" in ui


def test_only_selected_warehouse_workspace_is_rendered():
    ui = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    body = ui[ui.index("def render_warehouse_workspace"):]
    order = [
        'if current == "Обзор"',
        'elif current == "Каталог"',
        'elif current == "Поставки"',
        'elif current == "Новая поставка"',
        'elif current == "Приёмка"',
        'elif current == "Передача в бухгалтерию"',
        'elif current == "Операции"',
    ]
    positions = [body.index(token) for token in order]
    assert positions == sorted(positions)
    assert "Кэш чтения — 60 секунд" in body


def test_warehouse_has_phone_and_tablet_breakpoints():
    text = Path("src/warehouse.py").read_text(encoding="utf-8")
    management = Path("src/warehouse_management/ui.py").read_text(encoding="utf-8")
    assert '@media (max-width:900px)' in text
    assert '@media (max-width:640px)' in text
    assert '@media (max-width:640px)' in management
