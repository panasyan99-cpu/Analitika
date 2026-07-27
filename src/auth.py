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
_IDLE_TIMEOUT_SECONDS = 8 * 60 * 60
_MAX_ATTEMPTS = 5
_LOCK_SECONDS = 60


def _mapping_get(mapping: object, key: str, default: Any = None) -> Any:
    if isinstance(mapping, Mapping):
        return mapping.get(key, default)
    try:
        return mapping[key]  # type: ignore[index]
    except (KeyError, TypeError):
        return default


def configured_password() -> str:
    """Read the shared application password without placing it in source code."""
    secrets = getattr(st, "secrets", {})
    auth = _mapping_get(secrets, "auth", {})
    value = _mapping_get(auth, "password", "")
    if not value:
        value = _mapping_get(secrets, "APP_PASSWORD", "")
    if not value:
        value = os.getenv("ANALITIKA_APP_PASSWORD", "")
    return str(value or "")


def _logout() -> None:
    for key in (_SESSION_AUTH, _SESSION_ACTIVITY, _SESSION_ATTEMPTS, _SESSION_LOCK_UNTIL):
        st.session_state.pop(key, None)
    st.rerun()


def render_logout_control() -> None:
    """Render one unobtrusive session exit action above the workspace."""
    if not st.session_state.get(_SESSION_AUTH):
        return
    left, right = st.columns([10, 1])
    with right:
        if st.button("Выйти", key="analitika_logout", width="stretch"):
            _logout()


def require_password() -> bool:
    """Return True only after the shared password has been verified for this session."""
    now = time.time()
    authenticated = bool(st.session_state.get(_SESSION_AUTH, False))
    last_activity = float(st.session_state.get(_SESSION_ACTIVITY, 0.0) or 0.0)
    if authenticated and (not last_activity or now - last_activity <= _IDLE_TIMEOUT_SECONDS):
        st.session_state[_SESSION_ACTIVITY] = now
        return True
    if authenticated:
        st.session_state.pop(_SESSION_AUTH, None)

    expected = configured_password()
    st.markdown(
        """
        <style>
        .analitika-login-shell {
            max-width: 520px; margin: 8vh auto 1.5rem; padding: 2rem 2rem 1.6rem;
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
          <p class="analitika-login-copy">Введите пароль для доступа к внутренней системе.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not expected:
        st.error("Пароль доступа не настроен. Добавьте [auth] password в Streamlit Secrets.")
        return False

    lock_until = float(st.session_state.get(_SESSION_LOCK_UNTIL, 0.0) or 0.0)
    if lock_until > now:
        remaining = max(1, int(lock_until - now))
        st.warning(f"Слишком много попыток. Повторите через {remaining} сек.")
        return False

    with st.form("analitika_login_form", clear_on_submit=True):
        supplied = st.text_input("Пароль", type="password")
        submitted = st.form_submit_button("Войти", type="primary", width="stretch")
    if not submitted:
        return False

    if hmac.compare_digest(str(supplied), expected):
        st.session_state[_SESSION_AUTH] = True
        st.session_state[_SESSION_ACTIVITY] = now
        st.session_state[_SESSION_ATTEMPTS] = 0
        st.session_state.pop(_SESSION_LOCK_UNTIL, None)
        st.rerun()
        return True

    attempts = int(st.session_state.get(_SESSION_ATTEMPTS, 0) or 0) + 1
    st.session_state[_SESSION_ATTEMPTS] = attempts
    if attempts >= _MAX_ATTEMPTS:
        st.session_state[_SESSION_LOCK_UNTIL] = now + _LOCK_SECONDS
        st.session_state[_SESSION_ATTEMPTS] = 0
        st.error("Неверный пароль. Вход временно заблокирован на 60 секунд.")
    else:
        st.error("Неверный пароль.")
    return False
