from pathlib import Path

ROOT = Path(__file__).resolve().parent


def test_login_is_password_only() -> None:
    auth = (ROOT / "src" / "auth.py").read_text(encoding="utf-8")
    assert 'st.text_input("Пароль", type="password")' in auth
    assert '"Email"' not in auth
    assert "Введите общий пароль" in auth


def test_shared_login_grants_private_warehouse_access() -> None:
    auth = (ROOT / "src" / "auth.py").read_text(encoding="utf-8")
    assert 'st.session_state[_SESSION_ROLE] = ROLE_OPERATOR' in auth
    assert "server-side" in auth


def test_baserow_credentials_are_automatic() -> None:
    warehouse = (ROOT / "src" / "warehouse.py").read_text(encoding="utf-8")
    client = (ROOT / "src" / "warehouse_management" / "client.py").read_text(encoding="utf-8")
    assert "BASEROW_EMAIL" in warehouse
    assert "BASEROW_PASSWORD" in warehouse
    assert "/api/user/token-auth/" in client
    assert "_set_jwt_auth" in client
