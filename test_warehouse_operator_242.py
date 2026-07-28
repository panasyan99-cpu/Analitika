from pathlib import Path

from src.product_info import release_history_html


ROOT = Path(__file__).resolve().parent


def test_visible_history_contains_all_24x_releases() -> None:
    html = release_history_html(ROOT / "CHANGELOG.md")
    for version in ("2.4.0", "2.4.1", "2.4.2"):
        assert version in html


def test_maintenance_workspace_is_removed() -> None:
    ui = (ROOT / "src" / "warehouse_management" / "ui.py").read_text(encoding="utf-8")
    assert 'HISTORY_WORKSPACES = ("Операции",)' in ui
    assert "def render_maintenance" not in ui
    assert '"Создать и мигрировать"' not in ui


def test_operator_and_viewer_roles_are_server_side() -> None:
    auth = (ROOT / "src" / "auth.py").read_text(encoding="utf-8")
    assert 'ROLE_VIEWER = "viewer"' in auth
    assert 'ROLE_OPERATOR = "operator"' in auth
    assert "def can_write" in auth
    assert "operator_email" in auth
    assert "operator_password" in auth


def test_warehouse_uses_automatic_jwt_and_schema_setup() -> None:
    client = (ROOT / "src" / "warehouse_management" / "client.py").read_text(encoding="utf-8")
    ui = (ROOT / "src" / "warehouse_management" / "ui.py").read_text(encoding="utf-8")
    assert "/api/user/token-auth/" in client
    assert "_set_jwt_auth" in client
    assert "_auto_prepare_safe_schema" in ui
    assert "Повторить автоматическую настройку" in ui
