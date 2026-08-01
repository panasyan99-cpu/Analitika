from __future__ import annotations

import re


def _compact(value: object) -> str:
    return re.sub(r"[^A-ZА-Я0-9]", "", str(value or "").upper().replace("Ё", "Е"))


def analytics_store_name(value: object) -> str:
    """Return the user-facing store name used by analytical workspaces.

    The two 63 locations are intentionally separate in analytics. Legacy labels
    that do not say whether the location is Timing or Retail remain ``63``
    instead of being guessed.
    """
    text = " ".join(str(value or "").strip().split())
    compact = _compact(text)

    if compact.startswith("63"):
        if "RETAIL" in compact:
            return "63 Retail"
        if "TIMING" in compact or "TIMINGS" in compact:
            return "63 Timing"
        return "63"

    return text


def supplier_order_store_name(value: object) -> str:
    """Return the stock-planning store name.

    Supplier order keeps the historical rule: both 63 locations form one stock
    pool and are therefore returned as a single ``63`` store.
    """
    analytics_name = analytics_store_name(value)
    return "63" if analytics_name.startswith("63") else analytics_name
