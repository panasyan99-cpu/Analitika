from __future__ import annotations

from dataclasses import dataclass
from html import escape
from pathlib import Path
from typing import Literal

import streamlit as st

from src.app_meta import APP_VERSION


StatusTone = Literal["neutral", "success", "warning", "error"]
NavKind = Literal["anchor", "button"]


@dataclass(frozen=True)
class NavigationItem:
    """One item in the shared sidebar navigation."""

    item_id: str
    label: str
    href: str | None = None
    enabled: bool = True
    current: bool = False
    kind: NavKind = "anchor"


@dataclass
class SidebarResult:
    clicked_item: str | None
    action_clicked: bool
    status_slot: object


def _status_html(text: str, tone: StatusTone) -> str:
    safe_tone = tone if tone in {"neutral", "success", "warning", "error"} else "neutral"
    return (
        f'<div class="sidebar-status sidebar-status-{safe_tone}">'
        f'<span class="sidebar-status-dot" aria-hidden="true"></span>'
        f'<span>{escape(text)}</span>'
        "</div>"
    )


def update_sidebar_status(slot: object, text: str, tone: StatusTone = "neutral") -> None:
    """Update a status placeholder returned by render_sidebar."""
    markdown = getattr(slot, "markdown", None)
    if callable(markdown):
        markdown(_status_html(text, tone), unsafe_allow_html=True)


def _anchor_html(item: NavigationItem) -> str:
    classes = ["sidebar-nav-item"]
    if item.current:
        classes.append("is-current")
    class_name = " ".join(classes)
    if not item.enabled or not item.href:
        return (
            f'<span class="{class_name} is-disabled" aria-disabled="true">'
            f'{escape(item.label)}</span>'
        )
    return (
        f'<a class="{class_name}" href="{escape(item.href, quote=True)}">'
        f'{escape(item.label)}</a>'
    )


def render_sidebar(
    *,
    module_title: str,
    navigation_title: str,
    items: list[NavigationItem],
    status_text: str,
    status_tone: StatusTone = "neutral",
    source_text: str,
    action_label: str | None = None,
    action_key: str | None = None,
) -> SidebarResult:
    """Render the single sidebar shell used by every Analitika workspace.

    Anchor items scroll within a report. Button items are used by lazy workspaces
    such as Baserow, while sharing the exact same visual language.
    """
    clicked_item: str | None = None
    action_clicked = False

    with st.sidebar:
        logo = Path(__file__).resolve().parents[1] / "assets" / "logo.png"
        if logo.exists():
            st.image(str(logo), width="stretch")

        st.markdown(
            f"""
            <div class="sidebar-product-header">
              <div class="sidebar-suite-title">Princess Jewelry Analytics</div>
              <div class="sidebar-module-title">{escape(module_title)}</div>
              <div class="sidebar-version">Analitika Web {escape(APP_VERSION)}</div>
            </div>
            <div class="sidebar-navigation-title">{escape(navigation_title)}</div>
            """,
            unsafe_allow_html=True,
        )

        # Keep true navigation controls in a dedicated keyed container.
        # Global action buttons use the light gold theme, while this container
        # is intentionally overridden by CSS to retain the dark sidebar menu.
        with st.container(key="sidebar_navigation_controls"):
            for item in items:
                if item.kind == "button":
                    if st.button(
                        item.label,
                        key=f"sidebar_nav_{item.item_id}",
                        width="stretch",
                        type="primary" if item.current else "secondary",
                        disabled=not item.enabled,
                    ):
                        clicked_item = item.item_id
                else:
                    st.markdown(_anchor_html(item), unsafe_allow_html=True)

        status_slot = st.empty()
        update_sidebar_status(status_slot, status_text, status_tone)
        st.markdown(
            f'<div class="sidebar-source">{escape(source_text)}</div>',
            unsafe_allow_html=True,
        )

        if action_label and action_key:
            st.markdown('<div class="sidebar-action-separator"></div>', unsafe_allow_html=True)
            action_clicked = st.button(
                action_label,
                key=action_key,
                width="stretch",
            )

        st.markdown(
            '<div class="sidebar-footer">Разработка: Vladimir Panasyan</div>',
            unsafe_allow_html=True,
        )

    return SidebarResult(clicked_item, action_clicked, status_slot)


def render_mobile_navigation(items: list[NavigationItem]) -> None:
    """Render the same navigation as a sticky, touch-friendly strip."""
    controls: list[str] = []
    for item in items:
        classes = ["mobile-nav-item"]
        if item.current:
            classes.append("is-current")
        class_name = " ".join(classes)
        if item.kind != "anchor" or not item.enabled or not item.href:
            controls.append(
                f'<span class="{class_name} is-disabled" aria-disabled="true">'
                f'{escape(item.label)}</span>'
            )
        else:
            controls.append(
                f'<a class="{class_name}" href="{escape(item.href, quote=True)}">'
                f'{escape(item.label)}</a>'
            )
    st.markdown(
        '<div class="mobile-nav-shell"><nav class="mobile-nav">'
        + "".join(controls)
        + "</nav></div>",
        unsafe_allow_html=True,
    )
