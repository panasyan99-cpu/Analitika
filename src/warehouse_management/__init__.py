from __future__ import annotations

from typing import Any


def render_warehouse_workspace(*args: Any, **kwargs: Any) -> Any:
    """Import Streamlit UI lazily so parser and service tests stay lightweight."""
    from .ui import render_warehouse_workspace as _render

    return _render(*args, **kwargs)


__all__ = ["render_warehouse_workspace"]
