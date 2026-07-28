from __future__ import annotations

import hmac
import os
import time
from typing import Any, Mapping

import streamlit as st

_SESSION_AUTH = "analitika_authenticated"
_SESSION_ACTIVITY = "analitika_last_activity"
_SESSION_ATTEMPTS = "analitika_login_attempts"
_SESSION_LOCK_UNTIL = "analitika_login_lock_until"
_SESSION_ROLE = "analitika_role"
_SESSION_EMAIL = "analitika_user_email"
_IDLE_TIMEOUT_SECONDS = 8 * 60 * 60
_MAX_ATTEMPTS = 5
_LOCK_SECONDS = 60
ROLE_VIEWER = "viewer"
ROLE_OPERATOR = "operator"


def _mapping_get(mapping: object, key: str, default: Any = None) -> Any:
    if isinstance(mapping, Mapping):
        return mapping.get(key, default)
    try:
        return mapping[key]  # type: ignore[index]
    except (KeyError, TypeError):
        return default


def _auth_section() -> object:
    secrets = getattr(st, "secrets", {})
    return _mapping_get(secrets, "auth", {})


def configured_password() -> str:
    """Read the shared view-only application password."""
    auth = _auth_section()
    value = _mapping_get(auth, "password", "")
    if not value:
        value = _mapping_get(getattr(st, "secrets", {}), "APP_PASSWORD", "")
    if not value:
        value = os.getenv("ANALITIKA_APP_PASSWORD", "")
    return str(value or "")


def configured_operator() -> tuple[str, str]:
    """Return the single warehouse operator account from private configuration."""
    try:
        from src.private_operator import OPERATOR_EMAIL, OPERATOR_PASSWORD
    except ImportError:
        OPERATOR_EMAIL = ""
        OPERATOR_PASSWORD = ""
    auth = _auth_section()
    email = str(
        _mapping_get(auth, "operator_email", "")
        or os.getenv("ANALITIKA_OPERATOR_EMAIL", "")
        or OPERATOR_EMAIL
        or ""
    ).strip()
    password = str(
        _mapping_get(auth, "operator_password", "")
        or os.getenv("ANALITIKA_OPERATOR_PASSWORD", "")
        or OPERATOR_PASSWORD
        or ""
    )
    return email, password


def current_role() -> str:
    role = str(st.session_state.get(_SESSION_ROLE) or ROLE_VIEWER)
    return ROLE_OPERATOR if role == ROLE_OPERATOR else ROLE_VIEWER


def current_user_email() -> str:
    return str(st.session_state.get(_SESSION_EMAIL) or "")


def can_write() -> bool:
    """Only the configured operator may create or change warehouse data."""
    return bool(st.session_state.get(_SESSION_AUTH)) and current_role() == ROLE_OPERATOR


def _logout() -> None:
    for key in (
        _SESSION_AUTH,
        _SESSION_ACTIVITY,
        _SESSION_ATTEMPTS,
        _SESSION_LOCK_UNTIL,
        _SESSION_ROLE,
        _SESSION_EMAIL,
    ):
        st.session_state.pop(key, None)
    st.rerun()


def render_logout_control() -> None:
    """Render the active access level and one unobtrusive exit action."""
    if not st.session_state.get(_SESSION_AUTH):
        return
    left, role_column, exit_column = st.columns([8, 2, 1])
    with role_column:
        if can_write():
            label = "Управление"
            email = current_user_email()
            st.caption(f"{label} · {email}" if email else label)
        else:
            st.caption("Только просмотр")
    with exit_column:
        if st.button("Выйти", key="analitika_logout", width="stretch"):
            _logout()


def _authenticate_success(*, role: str, email: str, now: float) -> None:
    st.session_state[_SESSION_AUTH] = True
    st.session_state[_SESSION_ACTIVITY] = now
    st.session_state[_SESSION_ATTEMPTS] = 0
    st.session_state[_SESSION_ROLE] = role
    st.session_state[_SESSION_EMAIL] = email
    st.session_state.pop(_SESSION_LOCK_UNTIL, None)
    st.rerun()


def require_password() -> bool:
    """Authenticate a viewer with the shared password or the single operator."""
    now = time.time()
    authenticated = bool(st.session_state.get(_SESSION_AUTH, False))
    last_activity = float(st.session_state.get(_SESSION_ACTIVITY, 0.0) or 0.0)
    if authenticated and (not last_activity or now - last_activity <= _IDLE_TIMEOUT_SECONDS):
        st.session_state[_SESSION_ACTIVITY] = now
        return True
    if authenticated:
        for key in (_SESSION_AUTH, _SESSION_ROLE, _SESSION_EMAIL):
            st.session_state.pop(key, None)

    viewer_password = configured_password()
    operator_email, operator_password = configured_operator()
    st.markdown(
        """
        <style>
        .analitika-login-shell {
            max-width: 540px; margin: 8vh auto 1.5rem; padding: 2rem 2rem 1.6rem;
            border: 1px solid rgba(159,112,42,.28); border-radius: 22px;
            background: linear-gradient(180deg, rgba(255,253,248,.98), rgba(250,244,233,.98));
            box-shadow: 0 18px 55px rgba(65,45,20,.10); text-align: center;
        }
        .analitika-login-kicker {color:#a16b20; font-size:.78rem; font-weight:800; letter-spacing:.16em;}
        .analitika-login-title {font-family:Georgia,serif; font-size:2rem; color:#24180f; margin:.45rem 0 .45rem;}
        .analitika-login-copy {color:#6d6258; margin:0;}
        @media (max-width: 640px) {
            .analitika-login-shell {margin-top:4vh; padding:1.45rem 1.15rem 1.2rem; border-radius:18px;}
            .analitika-login-title {font-size:1.65rem;}
        }
        </style>
        <div class="analitika-login-shell">
          <div class="analitika-login-kicker">PRINCESS JEWELRY</div>
          <div class="analitika-login-title">Analitika</div>
          <p class="analitika-login-copy">Руководители входят в режим просмотра. Рабочий аккаунт открывает операции.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not viewer_password and not (operator_email and operator_password):
        st.error("Доступ не настроен. Добавьте [auth] в Streamlit Secrets.")
        return False

    lock_until = float(st.session_state.get(_SESSION_LOCK_UNTIL, 0.0) or 0.0)
    if lock_until > now:
        remaining = max(1, int(lock_until - now))
        st.warning(f"Слишком много попыток. Повторите через {remaining} сек.")
        return False

    with st.form("analitika_login_form", clear_on_submit=True):
        supplied_email = st.text_input(
            "Email",
            placeholder="Оставьте пустым для режима просмотра",
        ).strip()
        supplied_password = st.text_input("Пароль", type="password")
        submitted = st.form_submit_button("Войти", type="primary", width="stretch")
    if not submitted:
        return False

    operator_match = bool(
        operator_email
        and operator_password
        and hmac.compare_digest(supplied_email.casefold(), operator_email.casefold())
        and hmac.compare_digest(str(supplied_password), operator_password)
    )
    if operator_match:
        _authenticate_success(role=ROLE_OPERATOR, email=operator_email, now=now)
        return True

    viewer_match = bool(
        not supplied_email
        and viewer_password
        and hmac.compare_digest(str(supplied_password), viewer_password)
    )
    if viewer_match:
        # Backward-compatible deployments without a separate operator account
        # keep their previous full-access behavior until operator credentials
        # are added to Secrets.
        role = ROLE_VIEWER if operator_email and operator_password else ROLE_OPERATOR
        _authenticate_success(role=role, email="", now=now)
        return True

    attempts = int(st.session_state.get(_SESSION_ATTEMPTS, 0) or 0) + 1
    st.session_state[_SESSION_ATTEMPTS] = attempts
    if attempts >= _MAX_ATTEMPTS:
        st.session_state[_SESSION_LOCK_UNTIL] = now + _LOCK_SECONDS
        st.session_state[_SESSION_ATTEMPTS] = 0
        st.error("Неверные данные. Вход временно заблокирован на 60 секунд.")
    else:
        st.error("Неверный email или пароль.")
    return False
