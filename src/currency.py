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
    """Render the single site-wide FX editor on the main page.

    The widget is intentionally shown before any report uploader so a user can
    set the rate first. Streamlit reruns the page immediately after the value is
    confirmed (Enter, +/- control or leaving the field), so every workspace
    receives the updated rate from the same session-state key.
    """
    get_vnd_per_usd()
    with st.container(border=True):
        left, right = st.columns([1, 2], vertical_alignment="center")
        with left:
            st.markdown("### Единый курс Analitika")
            rate = float(
                st.number_input(
                    "Курс VND за 1 USD",
                    min_value=1_000,
                    max_value=100_000,
                    step=100,
                    format="%d",
                    key=GLOBAL_FX_SESSION_KEY,
                    help=(
                        "Один курс используется в обычном отчете, сравнении периодов, "
                        "складской аналитике и разделе «Заказ Sonu»."
                    ),
                )
            )
        with right:
            st.markdown(
                f"**1 USD = {format_vnd(rate)} VND**  \n"
                "Изменение применяется ко всем денежным показателям сайта сразу после подтверждения значения."
            )
    return rate
