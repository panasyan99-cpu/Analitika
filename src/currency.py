from __future__ import annotations

from typing import Any

import streamlit as st


DEFAULT_VND_PER_USD = 26_300
GLOBAL_FX_SESSION_KEY = "site_vnd_per_usd"
LEGACY_FX_SESSION_KEYS = ("analytics_fx_rate", "sonu_fx_rate")


def _valid_rate(value: Any) -> float | None:
    try:
        rate = float(value)
    except (TypeError, ValueError):
        return None
    return rate if rate > 0 else None


def get_vnd_per_usd() -> float:
    """Return the single VND-per-USD rate shared by every site workspace."""
    current = _valid_rate(st.session_state.get(GLOBAL_FX_SESSION_KEY))
    if current is not None:
        return current

    # Preserve a user's current value when upgrading from versions that kept
    # separate rates for the standard analytics and Sonu workspaces.
    for legacy_key in LEGACY_FX_SESSION_KEYS:
        legacy_rate = _valid_rate(st.session_state.get(legacy_key))
        if legacy_rate is not None:
            st.session_state[GLOBAL_FX_SESSION_KEY] = int(round(legacy_rate))
            return legacy_rate

    st.session_state[GLOBAL_FX_SESSION_KEY] = DEFAULT_VND_PER_USD
    return float(DEFAULT_VND_PER_USD)


def vnd_to_usd(value: float | int, rate: float | None = None) -> float:
    active_rate = _valid_rate(rate) or get_vnd_per_usd()
    return float(value or 0) / active_rate


def format_vnd(value: float | int) -> str:
    return f"{float(value):,.0f}".replace(",", " ")


def render_global_fx_control() -> float:
    """Render one compact site-wide VND/USD editor."""
    get_vnd_per_usd()
    with st.container(key="global_fx_compact"):
        label_col, input_col = st.columns([1.25, 1], gap="medium", vertical_alignment="center")
        with label_col:
            current = get_vnd_per_usd()
            st.markdown(
                f'<div class="fx-compact-title">Курс VND / USD</div>'
                f'<div class="fx-compact-value">1 USD = {format_vnd(current)} VND · единый для всех разделов</div>',
                unsafe_allow_html=True,
            )
        with input_col:
            rate = float(
                st.number_input(
                    "Курс VND за 1 USD",
                    min_value=1_000,
                    max_value=100_000,
                    step=100,
                    format="%d",
                    key=GLOBAL_FX_SESSION_KEY,
                    label_visibility="collapsed",
                    help="Используется во всех денежных показателях сайта.",
                )
            )
    return rate
