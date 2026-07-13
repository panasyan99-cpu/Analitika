from __future__ import annotations

import gc
import hashlib
import tempfile
from html import escape
from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook

from src.report import (
    COLORED_ORDER,
    PEARL_ORDER,
    PRODUCT_ORDER,
    SEG_ORDER,
    TOP_ORDER,
    StoreData,
    build_report_units,
    classify,
    extract_period,
    norm_product,
    normalize_store_from_report,
    totals_for,
)

APP_VERSION = "1.1.9"
SEGMENT_LABELS = {
    "TOP STONES": "Top Stones",
    "PEARLS": "Pearls",
    "COLORED STONES": "Colored Stones",
}
SEGMENT_COLORS = {
    "TOP STONES": "#7030A0",
    "PEARLS": "#D3A338",
    "COLORED STONES": "#548235",
}
LIGHT_COLORS = {
    "TOP STONES": "#E9DDF1",
    "PEARLS": "#F5E7B8",
    "COLORED STONES": "#DDE8D4",
}
STONE_ORDERS = {
    "TOP STONES": TOP_ORDER,
    "PEARLS": PEARL_ORDER,
    "COLORED STONES": COLORED_ORDER,
}
PRODUCT_LABELS = {
    "Earrings": "Серьги",
    "Ring": "Кольца",
    "Pendant": "Подвески",
    "Bracelet": "Браслеты",
    "Necklace": "Ожерелья",
    "Brooch": "Броши",
    "Pearl Necklace": "Жемчужные нити",
    "Pearl Bracelet": "Жемчужные браслеты",
    "Pearl Chain": "Жемчуг на цепочке",
    "Stone": "Камни",
    "Other": "Другое",
}


# Plotly remains informative but cannot be accidentally changed on touch devices.
# Hover/tap tooltips stay enabled; zooming, panning, selection, editing and export
# controls are disabled. Streamlit dataframes intentionally remain interactive.
LOCKED_CHART_CONFIG = {
    "displayModeBar": False,
    "displaylogo": False,
    "scrollZoom": False,
    "doubleClick": False,
    "showTips": False,
    "editable": False,
    "staticPlot": False,
    "responsive": True,
    "showAxisDragHandles": False,
    "showAxisRangeEntryBoxes": False,
}


def lock_chart_interactions(fig: go.Figure) -> go.Figure:
    """Return a view-only Plotly figure while preserving hover/tap tooltips."""
    fig.update_layout(
        dragmode=False,
        clickmode="event",
        hovermode="closest",
        legend_itemclick=False,
        legend_itemdoubleclick=False,
    )

    cartesian_types = {
        "bar", "scatter", "scattergl", "box", "violin", "histogram",
        "histogram2d", "heatmap", "contour", "waterfall", "funnel",
        "candlestick", "ohlc",
    }
    if any(getattr(trace, "type", "") in cartesian_types for trace in fig.data):
        fig.update_xaxes(fixedrange=True)
        fig.update_yaxes(fixedrange=True)
    return fig


def locked_plotly_chart(fig: go.Figure, *, width: str = "stretch", key: str | None = None) -> None:
    """Render a locked chart without changing dataframe/table behaviour."""
    st.plotly_chart(
        lock_chart_interactions(fig),
        width=width,
        key=key,
        config=LOCKED_CHART_CONFIG,
    )


class StoredUpload:
    """Persistent in-session representation of an uploaded file.

    Streamlit removes widget-owned values when a file uploader is no longer
    rendered. Keeping immutable bytes under a separate session key lets users
    navigate across pages without uploading the file again.
    """

    def __init__(self, name: str, data: bytes) -> None:
        self.name = name
        self._data = data

    def getvalue(self) -> bytes:
        return self._data


def persist_uploads(uploaded_files) -> None:
    if uploaded_files:
        payloads = [
            {"name": item.name, "data": bytes(item.getvalue())}
            for item in uploaded_files
        ]
        previous = st.session_state.get("uploaded_payloads", [])
        previous_signature = [(x.get("name"), len(x.get("data", b""))) for x in previous]
        new_signature = [(x["name"], len(x["data"])) for x in payloads]
        if previous_signature != new_signature:
            st.session_state["uploaded_payloads"] = payloads


def saved_uploads() -> list[StoredUpload]:
    return [
        StoredUpload(item["name"], item["data"])
        for item in st.session_state.get("uploaded_payloads", [])
    ]


def clear_saved_uploads() -> None:
    st.session_state.pop("uploaded_payloads", None)
    st.session_state.pop("upload_widget", None)
    st.session_state.pop("report_cache_signature", None)
    st.session_state.pop("report_cache_stores", None)
    st.session_state.pop("report_cache_errors", None)
    st.session_state.pop("report_cache_suppliers", None)


def uploads_signature(uploaded_files: list[StoredUpload]) -> str:
    """Stable content signature used to reuse parsed report data across reruns."""
    digest = hashlib.sha256()
    for uploaded in uploaded_files:
        data = uploaded.getvalue()
        digest.update(uploaded.name.encode("utf-8", errors="replace"))
        digest.update(len(data).to_bytes(8, "big"))
        digest.update(data)
    return digest.hexdigest()


st.set_page_config(
    page_title="Analitika — Princess Jewelry",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="auto",
)


def _css() -> str:
    return """
<style>
:root {
  --gold: #b7893f;
  --gold-soft: #ead8b8;
  --ink: #111111;
  --muted: #6c6c6c;
  --line: #e9e4dc;
  --paper: #fbfaf8;
}
html, body, [class*="css"] { font-family: Inter, Arial, sans-serif; }
.stApp {
  background:
    radial-gradient(circle at 72% 18%, rgba(230,212,183,.20), transparent 24%),
    linear-gradient(135deg, #ffffff 0%, #fbfaf8 72%, #f6f1e9 100%);
  color: var(--ink);
}
[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #090806 0%, #15110b 100%);
  border-right: 1px solid #3a2b16;
  color: #f5ead8;
}
[data-testid="stSidebar"] * { color: #f5ead8; }
[data-testid="stSidebar"] > div:first-child { padding-top: 1.2rem; }
.block-container { padding-top: 1.2rem; padding-bottom: 3rem; max-width: 1500px; }
.brand-card {
  border: 1px solid var(--line); border-radius: 18px; background: rgba(255,255,255,.92);
  padding: 22px 24px; box-shadow: 0 10px 35px rgba(34,24,9,.05); margin-bottom: 18px;
}
.brand-kicker { color: var(--gold); font-size: 12px; letter-spacing: .12em; text-transform: uppercase; font-weight: 700; }
.brand-title { font-family: Georgia, serif; font-size: 44px; margin: 4px 0 4px; color: #171411; }
.brand-subtitle { color: var(--muted); font-size: 15px; }
.upload-panel {
  border: 1px dashed #c9aa72; border-radius: 18px; background: rgba(255,255,255,.78);
  padding: 18px 22px; margin: 6px 0 20px;
}
.kpi-card {
  border: 1px solid var(--line); border-radius: 14px; background: rgba(255,255,255,.95);
  padding: 18px 18px 16px; min-height: 118px; box-shadow: 0 8px 25px rgba(34,24,9,.045);
  overflow: visible;
}
.kpi-label { color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .06em; }
.kpi-value {
  font-family: Georgia, serif; font-size: clamp(18px, 2vw, 29px); line-height: 1.15;
  font-weight: 700; color: #16120d; margin-top: 9px; white-space: normal;
  overflow-wrap: anywhere; word-break: normal;
}
.kpi-note { color: var(--gold); font-size: 12px; margin-top: 6px; }
.section-title { font-family: Georgia, serif; font-size: 30px; margin: 22px 0 10px; }
.section-divider {
  margin: 38px 0 18px; padding: 18px 22px; border-radius: 16px;
  background: linear-gradient(90deg, rgba(183,137,63,.14), rgba(255,255,255,.96) 45%, rgba(183,137,63,.08));
  border-top: 1px solid rgba(183,137,63,.55); border-bottom: 1px solid rgba(183,137,63,.28);
  box-shadow: 0 10px 28px rgba(34,24,9,.045);
}
.section-divider-kicker { color: var(--gold); font-size: 11px; font-weight: 800; letter-spacing: .16em; text-transform: uppercase; }
.section-divider-title { font-family: Georgia, serif; color: #17120c; font-size: 28px; margin-top: 4px; }
.section-divider-copy { color: var(--muted); font-size: 13px; margin-top: 5px; }
.analysis-panel {
  margin: 16px 0 26px; padding: 18px 20px; border-radius: 15px;
  background: rgba(255,255,255,.94); border: 1px solid var(--line);
  box-shadow: 0 9px 26px rgba(34,24,9,.04);
}
.analysis-panel-title { font-family: Georgia, serif; font-size: 20px; color: #6f4b16; margin-bottom: 8px; }
.analysis-line { padding: 8px 0; border-bottom: 1px solid #f0ece5; color: #28231d; }
.analysis-line:last-child { border-bottom: none; }
.insight {
  border-left: 4px solid var(--gold); background: rgba(255,255,255,.93); border-radius: 0 12px 12px 0;
  padding: 13px 15px; margin: 8px 0; border-top: 1px solid var(--line); border-right: 1px solid var(--line); border-bottom: 1px solid var(--line);
}
.filter-panel {
  border: 1px solid var(--line); border-radius: 15px; background: rgba(255,255,255,.92);
  padding: 14px 16px 4px; margin: 8px 0 14px; box-shadow: 0 8px 22px rgba(34,24,9,.035);
}
.small-muted { color: var(--muted); font-size: 12px; }
div[data-testid="stFileUploader"] section {
  border: 1px dashed #c9aa72; border-radius: 14px; background: #fffdf9;
}
div.stButton > button {
  border-radius: 9px; border: 1px solid #2a2114; background: linear-gradient(90deg, #111 0%, #2b241c 100%);
  color: #e8c98e; font-weight: 700; min-height: 44px;
}
div.stButton > button:hover {
  border-color: #b7893f; color: #fff; box-shadow: 0 5px 18px rgba(183,137,63,.25);
}
[data-testid="stMetric"] { border: 1px solid var(--line); padding: 12px; border-radius: 12px; background: #fff; }
hr { border-color: var(--line); }
[data-testid="stSidebar"] [role="radiogroup"] { gap: 0.35rem; }
[data-testid="stSidebar"] [role="radiogroup"] label {
  border-radius: 10px; padding: 0.62rem 0.72rem; border: 1px solid transparent;
  transition: all .15s ease; background: transparent;
}
[data-testid="stSidebar"] [role="radiogroup"] label:hover {
  background: rgba(183,137,63,.14); border-color: rgba(183,137,63,.35);
}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {
  background: linear-gradient(90deg, rgba(183,137,63,.32) 0%, rgba(183,137,63,.10) 100%);
  border-color: #b7893f; color: #f2cf8c; font-weight: 700;
}
.side-nav { display:flex; flex-direction:column; gap:7px; margin:.15rem 0 1rem; }
.side-nav a,
.side-nav a:visited,
[data-testid="stSidebar"] .side-nav a,
[data-testid="stSidebar"] .side-nav a:visited {
  display:block; color:#f5ead8 !important; text-decoration:none !important;
  border-left:2px solid transparent; border-radius:0 10px 10px 0;
  padding:.66rem .78rem; font-size:.94rem; line-height:1.25;
  transition:background-color .16s ease, color .16s ease, border-color .16s ease, transform .16s ease;
}
.side-nav a:hover,
[data-testid="stSidebar"] .side-nav a:hover {
  color:#f1cc85 !important; text-decoration:none !important;
  background:linear-gradient(90deg, rgba(183,137,63,.24), rgba(183,137,63,.07));
  border-left-color:#b7893f; transform:translateX(2px);
}
.side-nav a:focus,
.side-nav a:active,
[data-testid="stSidebar"] .side-nav a:focus,
[data-testid="stSidebar"] .side-nav a:active {
  color:#ffe2a8 !important; text-decoration:none !important; outline:none;
  background:linear-gradient(90deg, rgba(183,137,63,.32), rgba(183,137,63,.10));
  border-left-color:#d4a95c;
}
.nav-hint { color:#cdbb9b; font-size:12px; margin:.2rem 0 .8rem; }
.executive-banner {
  position:relative; overflow:hidden; margin:4px 0 18px; padding:24px 26px;
  border-radius:18px; border:1px solid rgba(183,137,63,.46);
  background:
    radial-gradient(circle at 88% 18%, rgba(207,166,92,.24), transparent 30%),
    linear-gradient(135deg, #12100c 0%, #21190f 62%, #342511 100%);
  box-shadow:0 18px 45px rgba(38,25,7,.16); color:#fff7e8;
}
.executive-banner:after {
  content:""; position:absolute; inset:0; pointer-events:none;
  background:linear-gradient(115deg, transparent 0%, rgba(255,255,255,.05) 47%, transparent 72%);
}
.executive-banner-content { position:relative; z-index:2; max-width:920px; }
.executive-eyebrow { color:#e7c98e; font-size:11px; font-weight:800; letter-spacing:.17em; text-transform:uppercase; }
.executive-title { font-family:Georgia,serif; font-size:34px; line-height:1.12; margin:7px 0 7px; color:#fffaf0; }
.executive-copy { color:#ddcfb7; font-size:14px; line-height:1.55; }
.executive-note {
  margin:8px 0 14px; padding:11px 14px; border-radius:11px;
  border:1px solid #eadfcd; background:rgba(255,255,255,.78); color:#5e5549; font-size:12px;
}
.about-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:14px; margin:12px 0 20px; }
.about-card { border:1px solid var(--line); border-radius:14px; background:rgba(255,255,255,.94); padding:18px 19px; box-shadow:0 8px 24px rgba(34,24,9,.035); }
.about-card h4 { font-family:Georgia,serif; color:#6f4b16; font-size:19px; margin:0 0 8px; }
.about-card p { color:#4f4941; font-size:14px; line-height:1.55; margin:0; }
.about-card ul { color:#4f4941; font-size:14px; line-height:1.58; margin:.25rem 0 0; padding-left:1.1rem; }
.about-step { border-left:3px solid #b7893f; padding-left:12px; margin:9px 0; color:#302a23; }
@media (max-width: 780px) { .about-grid { grid-template-columns:1fr; } }

.luxury-hero {
  position: relative; overflow: hidden; min-height: 310px; border-radius: 24px;
  border: 1px solid #eadfcd; margin-bottom: 22px; padding: 44px 46px;
  background:
    radial-gradient(circle at 84% 20%, rgba(183,137,63,.24), transparent 26%),
    radial-gradient(circle at 72% 76%, rgba(234,216,184,.42), transparent 32%),
    linear-gradient(135deg, #fffdf9 0%, #f7f0e4 58%, #efe0c5 100%);
  box-shadow: 0 24px 65px rgba(56,36,10,.12);
}
.luxury-hero:after {
  content:""; position:absolute; inset:0; pointer-events:none;
  background: linear-gradient(135deg, rgba(183,137,63,.08), transparent 45%);
}
.luxury-hero-content { position:relative; z-index:2; max-width:620px; }
.luxury-eyebrow { color:#9d6f29; font-size:12px; font-weight:800; letter-spacing:.17em; text-transform:uppercase; }
.luxury-title { font-family: Georgia, 'Times New Roman', serif; font-size: clamp(42px, 5vw, 66px); line-height:1.02; margin:10px 0 12px; color:#17120c; }
.luxury-title span { color:#a8742a; }
.luxury-copy { color:#5e5549; font-size:17px; line-height:1.65; max-width:560px; }
.luxury-badges { display:flex; flex-wrap:wrap; gap:10px; margin-top:22px; }
.luxury-badge { border:1px solid rgba(183,137,63,.32); background:rgba(255,255,255,.78); color:#6f4b16; border-radius:999px; padding:8px 12px; font-size:12px; font-weight:700; }
.luxury-divider { width:70px; height:2px; background:linear-gradient(90deg,#b7893f,transparent); margin:18px 0; }

[data-testid="stSidebar"]:before {
  content:""; display:block; height:6px; background:linear-gradient(90deg,#15120e,#b7893f,#15120e);
}
[data-testid="stSidebar"] { box-shadow: 12px 0 35px rgba(50,32,8,.06); }

@media (max-width: 900px) {
  .luxury-hero { padding:30px 26px; min-height:280px; background-position:68% center; }
  .luxury-hero:before { content:""; position:absolute; inset:0; background:rgba(255,255,255,.40); }
  .luxury-title { font-size:42px; }
}

/* Responsive shell: one codebase for desktop, iPad and phones. */
.mobile-nav-shell { display:none; }
[id] { scroll-margin-top: 86px; }
[data-testid="stPlotlyChart"],
[data-testid="stPlotlyChart"] > div,
.js-plotly-plot,
.plot-container,
.svg-container { max-width:100% !important; }
[data-testid="stDataFrame"] { max-width:100%; }

@media (max-width: 900px) {
  .block-container {
    max-width:100%; padding:0.85rem 1rem 2.25rem; overflow-x:hidden;
  }
  [data-testid="stSidebar"] { width:min(88vw, 350px) !important; }
  .mobile-nav-shell {
    display:block; position:sticky; top:0.35rem; z-index:999;
    margin:0 0 0.9rem; padding:0.48rem;
    border:1px solid rgba(183,137,63,.30); border-radius:14px;
    background:rgba(255,253,249,.94); backdrop-filter:blur(12px);
    -webkit-backdrop-filter:blur(12px); box-shadow:0 8px 28px rgba(35,24,10,.10);
  }
  .mobile-nav {
    display:flex; gap:0.45rem; overflow-x:auto; overscroll-behavior-x:contain;
    scrollbar-width:none; -webkit-overflow-scrolling:touch; white-space:nowrap;
  }
  .mobile-nav::-webkit-scrollbar { display:none; }
  .mobile-nav a,
  .mobile-nav a:visited {
    flex:0 0 auto; min-height:42px; display:inline-flex; align-items:center;
    color:#2b2114 !important; text-decoration:none !important;
    border:1px solid rgba(183,137,63,.28); border-radius:999px;
    background:#fff; padding:0.55rem 0.78rem; font-size:0.82rem; font-weight:700;
  }
  .mobile-nav a:active { background:#f3e5cd; border-color:#b7893f; }

  [data-testid="stHorizontalBlock"] {
    flex-wrap:wrap !important; gap:0.85rem !important; align-items:stretch !important;
  }
  [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
    min-width:0 !important;
  }
  /* KPI/filter rows with 3+ columns become a comfortable 2-column grid on iPad. */
  [data-testid="stHorizontalBlock"]:has(> [data-testid="stColumn"]:nth-child(3)) > [data-testid="stColumn"] {
    flex:1 1 calc(50% - 0.5rem) !important;
    width:calc(50% - 0.5rem) !important;
    min-width:260px !important;
  }
  .brand-card, .upload-panel, .analysis-panel, .section-divider { border-radius:14px; }
  .executive-banner { padding:21px 20px; border-radius:15px; }
  .executive-title { font-size:29px; }
  .section-divider { margin:28px 0 14px; padding:15px 17px; }
  .section-divider-title { font-size:25px; }
  .section-title { font-size:27px; }
  .kpi-card { min-height:108px; padding:15px; }
  .kpi-value { font-size:clamp(20px, 4vw, 28px); }
  [data-testid="stPlotlyChart"] { width:100% !important; overflow:visible !important; }
  [data-baseweb="tab-list"] {
    overflow-x:auto !important; flex-wrap:nowrap !important; scrollbar-width:none;
    -webkit-overflow-scrolling:touch;
  }
  [data-baseweb="tab-list"]::-webkit-scrollbar { display:none; }
  [data-baseweb="tab"] { flex:0 0 auto !important; min-width:max-content; }
  div[data-baseweb="select"] > div, input, textarea { min-height:44px; }
}

@media (max-width: 820px) {
  /* Two-column chart groups stack in iPad portrait so labels stay readable. */
  [data-testid="stHorizontalBlock"]:not(:has(> [data-testid="stColumn"]:nth-child(3))) {
    flex-direction:column !important;
  }
  [data-testid="stHorizontalBlock"]:not(:has(> [data-testid="stColumn"]:nth-child(3))) > [data-testid="stColumn"] {
    flex:1 1 100% !important; width:100% !important; min-width:0 !important;
  }
  .about-grid { grid-template-columns:1fr; }
  .luxury-hero { min-height:auto; padding:26px 22px; border-radius:18px; }
  .luxury-title { font-size:38px; }
  .luxury-copy { font-size:15px; line-height:1.55; }
}

@media (max-width: 600px) {
  .block-container { padding:0.65rem 0.72rem 1.8rem; }
  [data-testid="stHorizontalBlock"] { flex-direction:column !important; }
  [data-testid="stHorizontalBlock"] > [data-testid="stColumn"],
  [data-testid="stHorizontalBlock"]:has(> [data-testid="stColumn"]:nth-child(3)) > [data-testid="stColumn"] {
    flex:1 1 100% !important; width:100% !important; min-width:0 !important;
  }
  .luxury-hero { padding:22px 18px; margin-bottom:15px; }
  .luxury-title { font-size:32px; line-height:1.08; }
  .luxury-eyebrow { font-size:10px; letter-spacing:.13em; }
  .luxury-copy { font-size:14px; }
  .luxury-badges { gap:7px; margin-top:16px; }
  .luxury-badge { padding:7px 9px; font-size:11px; }
  .executive-banner { padding:18px 16px; }
  .executive-title { font-size:25px; }
  .executive-copy { font-size:13px; }
  .section-divider { padding:14px; margin:23px 0 12px; }
  .section-divider-title { font-size:22px; }
  .section-divider-copy { font-size:12px; }
  .section-title { font-size:24px; }
  .kpi-card { min-height:96px; }
  .kpi-value { font-size:23px; }
  .analysis-panel { padding:15px; }
  .analysis-panel-title { font-size:18px; }
  [data-testid="stDataFrame"] { overflow-x:auto !important; -webkit-overflow-scrolling:touch; }
}

</style>
"""


st.markdown(_css(), unsafe_allow_html=True)


def money(value: float) -> str:
    return f"{value:,.0f}".replace(",", " ")


def pct(value: float) -> str:
    return f"{value:.2%}".replace(".", ",")


def kpi_card(label: str, value: str, note: str = "") -> None:
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div><div class="kpi-note">{note}</div></div>',
        unsafe_allow_html=True,
    )


def base_store_name(name: str) -> str:
    return name.split(" — ")[0]


def is_tourist_flow_store(name: str) -> bool:
    """Stores excluded only from the retail-network leader ranking."""
    normalized = "".join(str(name).upper().split())
    return normalized == "OUTLET" or normalized.startswith("63")


def retail_leader_summary(store_summary: pd.DataFrame) -> pd.DataFrame:
    """Retail stores used for revenue/quantity leaders; keeps the full report intact."""
    if store_summary.empty or "Магазин" not in store_summary.columns:
        return store_summary.copy()
    retail = store_summary.loc[
        ~store_summary["Магазин"].astype(str).map(is_tourist_flow_store)
    ].copy()
    return retail if not retail.empty else store_summary.copy()


def segment_totals(store) -> dict[str, dict[str, float]]:
    result: dict[str, dict[str, float]] = {}
    for segment in SEG_ORDER:
        q, a = totals_for(store, seg=segment)
        result[segment] = {"qty": int(q), "amount": float(a)}
    return result


def network_summary(stores: Iterable) -> pd.DataFrame:
    rows = []
    for store in stores:
        segs = segment_totals(store)
        row = {
            "Магазин": base_store_name(store.name),
            "Период": store.period_text(),
            "Количество": store.total_qty,
            "Выручка": store.total_amount,
            "Средняя стоимость": store.total_amount / store.total_qty if store.total_qty else 0,
        }
        for seg in SEG_ORDER:
            row[f"{SEGMENT_LABELS[seg]} — шт. %"] = segs[seg]["qty"] / store.total_qty if store.total_qty else 0
            row[f"{SEGMENT_LABELS[seg]} — продажи %"] = segs[seg]["amount"] / store.total_amount if store.total_amount else 0
        rows.append(row)
    return pd.DataFrame(rows)


def network_segment_summary(stores: Iterable) -> pd.DataFrame:
    """Compact network-level segment mix for the executive brief."""
    rows: list[dict] = []
    stores = list(stores)
    total_qty = sum(int(store.total_qty) for store in stores)
    total_sales = sum(float(store.total_amount) for store in stores)
    for segment in SEG_ORDER:
        qty = 0
        sales = 0.0
        for store in stores:
            current_qty, current_sales = totals_for(store, seg=segment)
            qty += int(current_qty)
            sales += float(current_sales)
        rows.append({
            "Сегмент": SEGMENT_LABELS[segment],
            "Количество": qty,
            "Выручка": sales,
            "Средняя стоимость": sales / qty if qty else 0,
            "% количества": qty / total_qty if total_qty else 0,
            "% выручки": sales / total_sales if total_sales else 0,
        })
    return pd.DataFrame(rows)


def executive_store_summary(stores: Iterable) -> pd.DataFrame:
    """One-row-per-store management table used in the operational brief."""
    stores = list(stores)
    total_sales = sum(float(store.total_amount) for store in stores)
    rows: list[dict] = []
    for store in stores:
        segments = segment_totals(store)
        leader_segment = max(SEG_ORDER, key=lambda segment: segments[segment]["amount"])
        leader_sales = float(segments[leader_segment]["amount"])
        rows.append({
            "Магазин": base_store_name(store.name),
            "Выручка": float(store.total_amount),
            "% выручки сети": float(store.total_amount) / total_sales if total_sales else 0,
            "Количество": int(store.total_qty),
            "Средняя стоимость": float(store.total_amount) / int(store.total_qty) if store.total_qty else 0,
            "Главный сегмент": SEGMENT_LABELS[leader_segment],
            "% главного сегмента": leader_sales / float(store.total_amount) if store.total_amount else 0,
        })
    if not rows:
        return pd.DataFrame(columns=[
            "Магазин", "Выручка", "% выручки сети", "Количество",
            "Средняя стоимость", "Главный сегмент", "% главного сегмента",
        ])
    return pd.DataFrame(rows).sort_values("Выручка", ascending=False).reset_index(drop=True)


def executive_insights(
    stores: list[StoreData],
    store_summary: pd.DataFrame,
    segment_summary: pd.DataFrame,
    supplier_df: pd.DataFrame,
) -> list[str]:
    """Generate factual, decision-oriented observations without forecasting."""
    if store_summary.empty:
        return []

    lines: list[str] = []
    total_sales = float(store_summary["Выручка"].sum())

    retail_summary = retail_leader_summary(store_summary)
    revenue_leader = retail_summary.sort_values("Выручка", ascending=False).iloc[0]
    retail_total_sales = float(retail_summary["Выручка"].sum())
    retail_share = float(revenue_leader["Выручка"]) / retail_total_sales if retail_total_sales else 0
    lines.append(
        f"Лидер розничной сети по выручке — {revenue_leader['Магазин']}: "
        f"{money(float(revenue_leader['Выручка']))} VND, "
        f"или {pct(retail_share)} розничной сети. OUTLET и 63 в рейтинг не входят."
    )

    top_three_share = float(store_summary.head(3)["Выручка"].sum()) / total_sales if total_sales else 0
    lines.append(f"Три крупнейших магазина формируют {pct(top_three_share)} выручки сети.")

    if not segment_summary.empty:
        segment_leader = segment_summary.sort_values("Выручка", ascending=False).iloc[0]
        lines.append(
            f"Главный сегмент сети — {segment_leader['Сегмент']}: "
            f"{pct(float(segment_leader['% выручки']))} выручки."
        )

    avg_leader = store_summary.sort_values("Средняя стоимость", ascending=False).iloc[0]
    lines.append(
        f"Самая высокая средняя стоимость проданного изделия — в {avg_leader['Магазин']}: "
        f"{money(float(avg_leader['Средняя стоимость']))} VND."
    )

    concentration_leader = store_summary.sort_values("% главного сегмента", ascending=False).iloc[0]
    lines.append(
        f"Наибольшая концентрация на одном сегменте — в {concentration_leader['Магазин']}: "
        f"{concentration_leader['Главный сегмент']} дает "
        f"{pct(float(concentration_leader['% главного сегмента']))} выручки магазина."
    )

    if not supplier_df.empty:
        suppliers = supplier_summary(supplier_df)
        if not suppliers.empty:
            supplier_leader = suppliers.iloc[0]
            top_supplier_share = float(suppliers.head(3)["% выручки"].sum())
            lines.append(
                f"Лидер среди поставщиков — {supplier_leader['Поставщик']} "
                f"({pct(float(supplier_leader['% выручки']))}); топ-3 поставщика дают "
                f"{pct(top_supplier_share)} выручки."
            )

    return lines[:6]


def render_executive_brief(
    stores: list[StoreData],
    summary_df: pd.DataFrame,
    supplier_df: pd.DataFrame,
) -> None:
    """Compact iPad-friendly overview intended for company leadership."""
    store_summary = executive_store_summary(stores)
    segment_summary = network_segment_summary(stores)
    total_qty = int(summary_df["Количество"].sum())
    total_sales = float(summary_df["Выручка"].sum())
    average_item = total_sales / total_qty if total_qty else 0
    periods = sorted(set(summary_df["Период"].astype(str).tolist())) if "Период" in summary_df.columns else []
    period_label = periods[0] if len(periods) == 1 else f"{len(periods)} периода"

    st.markdown(
        '<div class="executive-banner"><div class="executive-banner-content">'
        '<div class="executive-eyebrow">ОПЕРАТИВНЫЙ РЕЖИМ · ДЛЯ РУКОВОДИТЕЛЯ</div>'
        '<div class="executive-title">Сеть в одном экране</div>'
        '<div class="executive-copy">Ключевые цифры, лидеры, структура продаж и точки концентрации. '
        'Подробные таблицы и разрезы остаются в разделах ниже.</div>'
        '</div></div>',
        unsafe_allow_html=True,
    )

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        kpi_card("Период", period_label)
    with k2:
        kpi_card("Выручка сети", f"{money(total_sales)} VND")
    with k3:
        kpi_card("Продано", f"{money(total_qty)} шт.")
    with k4:
        kpi_card("Средняя стоимость", f"{money(average_item)} VND", "не средний чек")
    with k5:
        kpi_card("Магазинов", str(len(stores)))

    if not store_summary.empty:
        retail_summary = retail_leader_summary(store_summary)
        revenue_leader = retail_summary.sort_values("Выручка", ascending=False).iloc[0]
        qty_leader = retail_summary.sort_values("Количество", ascending=False).iloc[0]
        avg_leader = store_summary.sort_values("Средняя стоимость", ascending=False).iloc[0]
        segment_leader = segment_summary.sort_values("Выручка", ascending=False).iloc[0]
        l1, l2, l3, l4 = st.columns(4)
        with l1:
            kpi_card(
                "Лидер по выручке розничной сети",
                escape(str(revenue_leader["Магазин"])),
                f"{money(float(revenue_leader['Выручка']))} VND · без OUTLET и 63",
            )
        with l2:
            kpi_card(
                "Лидер по количеству розничной сети",
                escape(str(qty_leader["Магазин"])),
                f"{money(float(qty_leader['Количество']))} шт. · без OUTLET и 63",
            )
        with l3:
            kpi_card(
                "Самая высокая средняя стоимость",
                escape(str(avg_leader["Магазин"])),
                f"{money(float(avg_leader['Средняя стоимость']))} VND",
            )
        with l4:
            kpi_card(
                "Главный сегмент",
                escape(str(segment_leader["Сегмент"])),
                pct(float(segment_leader["% выручки"])),
            )

    insight_panel(
        "Что важно сейчас",
        executive_insights(stores, store_summary, segment_summary, supplier_df),
    )

    left, right = st.columns(2)
    with left:
        locked_plotly_chart(
            horizontal_bar(
                store_summary,
                "Магазин",
                "Выручка",
                "Выручка по магазинам",
            ),
            width="stretch",
            key="executive_store_revenue",
        )
    with right:
        locked_plotly_chart(
            donut(
                segment_summary["Сегмент"].tolist(),
                segment_summary["Выручка"].tolist(),
                "Структура выручки по сегментам",
                [SEGMENT_COLORS[segment] for segment in SEG_ORDER],
            ),
            width="stretch",
            key="executive_segment_mix",
        )

    st.markdown("### Магазины одним взглядом")
    st.dataframe(formatted_table(store_summary), width="stretch", hide_index=True)

    if not supplier_df.empty:
        suppliers = supplier_summary(supplier_df)
        if not suppliers.empty:
            leader = suppliers.iloc[0]
            top_three_share = float(suppliers.head(3)["% выручки"].sum())
            other_rows = suppliers[suppliers["Поставщик"].astype(str).str.casefold() == "other"]
            other_share = float(other_rows["% выручки"].sum()) if not other_rows.empty else 0
            s1, s2, s3, s4 = st.columns(4)
            with s1:
                kpi_card("Поставщик №1", escape(str(leader["Поставщик"])), pct(float(leader["% выручки"])))
            with s2:
                kpi_card("Доля топ-3 поставщиков", pct(top_three_share))
            with s3:
                kpi_card("Поставщиков", str(int(suppliers["Поставщик"].nunique())))
            with s4:
                kpi_card("Other", pct(other_share), "доля выручки")
            st.markdown("### Крупнейшие поставщики")
            st.dataframe(
                formatted_table(suppliers.head(7)),
                width="stretch",
                hide_index=True,
            )

    st.markdown(
        '<div class="executive-note">Показатель «Средняя стоимость» рассчитывается как '
        'выручка ÷ количество проданных изделий. Это не средний чек.</div>',
        unsafe_allow_html=True,
    )


def segment_bar(df: pd.DataFrame, segment: str) -> go.Figure:
    qty_key = f"{SEGMENT_LABELS[segment]} — шт. %"
    sales_key = f"{SEGMENT_LABELS[segment]} — продажи %"
    fig = go.Figure()
    fig.add_bar(
        x=df["Магазин"], y=df[qty_key] * 100, name="Шт. %",
        marker_color=SEGMENT_COLORS[segment], text=[pct(v) for v in df[qty_key]], textposition="outside",
        hovertemplate="%{x}<br>Количество: %{y:.2f}%<extra></extra>",
    )
    fig.add_bar(
        x=df["Магазин"], y=df[sales_key] * 100, name="Продажи %",
        marker_color=LIGHT_COLORS[segment], text=[pct(v) for v in df[sales_key]], textposition="outside",
        hovertemplate="%{x}<br>Выручка: %{y:.2f}%<extra></extra>",
    )
    fig.update_layout(
        title=SEGMENT_LABELS[segment].upper(), barmode="group", height=380,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=30, r=20, t=55, b=35), legend=dict(orientation="h", y=1.06),
        yaxis=dict(title="%", range=[0, 105], gridcolor="#ece8e1"),
        xaxis=dict(title=""), font=dict(family="Arial", color="#1c1813"),
    )
    return fig


def donut(labels: list[str], values: list[float], title: str, colors: list[str] | None = None) -> go.Figure:
    # Outside labels need real breathing room in Streamlit columns.
    # `automargin` lets Plotly expand the drawable area instead of clipping callouts.
    pie_kwargs = {
        "labels": labels,
        "values": values,
        "hole": .58,
        "textinfo": "label+percent",
        "textposition": "auto",
        "automargin": True,
        "sort": False,
        "insidetextorientation": "horizontal",
        "hovertemplate": "%{label}<br>%{value:,.2f}<br>%{percent}<extra></extra>",
    }
    if colors:
        pie_kwargs["marker"] = dict(colors=colors)
    fig = go.Figure(go.Pie(**pie_kwargs))
    fig.update_traces(textfont=dict(size=11), outsidetextfont=dict(size=11))
    fig.update_layout(
        title=title, height=430, showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=85, r=85, t=60, b=55),
        font=dict(family="Arial", color="#1c1813"),
    )
    return fig


def horizontal_bar(df: pd.DataFrame, label_col: str, value_col: str, title: str, suffix: str = "") -> go.Figure:
    clean = df[df[value_col] > 0].copy().sort_values(value_col, ascending=True)
    labels = [f"{money(v)}{suffix}" for v in clean[value_col]]
    max_value = float(clean[value_col].max()) if not clean.empty else 0.0

    # Reserve extra x-axis space for labels printed outside the bars.
    # Longer numbers receive a little more headroom.
    longest_label = max((len(label) for label in labels), default=0)
    headroom = 1.30 if longest_label >= 12 else 1.22
    x_range = [0, max_value * headroom] if max_value > 0 else None

    fig = go.Figure(go.Bar(
        x=clean[value_col], y=clean[label_col], orientation="h",
        marker_color="#b7893f", text=labels, textposition="outside",
        cliponaxis=False, textfont=dict(size=11),
        hovertemplate="%{y}<br>%{text}<extra></extra>",
    ))
    fig.update_layout(
        title=title, height=max(330, 42 * len(clean) + 100),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=135, t=55, b=35),
        xaxis=dict(gridcolor="#ece8e1", range=x_range, automargin=True),
        yaxis=dict(title="", automargin=True),
    )
    return fig


def stone_dataframe(store) -> pd.DataFrame:
    rows = []
    for seg in SEG_ORDER:
        _, seg_amount = totals_for(store, seg=seg)
        for stone in STONE_ORDERS[seg]:
            q, a = totals_for(store, seg, stone)
            rows.append({
                "Сегмент": SEGMENT_LABELS[seg], "Камень": stone,
                "Количество": q, "% количества магазина": q / store.total_qty if store.total_qty else 0,
                "Выручка": a, "% выручки магазина": a / store.total_amount if store.total_amount else 0,
                "Средняя стоимость": a / q if q else 0,
                "% выручки сегмента": a / seg_amount if seg_amount else 0,
            })
    return pd.DataFrame(rows)


def product_dataframe(store, segment: str | None = None, stone: str | None = None) -> pd.DataFrame:
    rows: list[dict] = []
    stone_qty, stone_amount = totals_for(store, segment, stone) if segment and stone else (0, 0)
    for (seg, stone_name), products in store.data.items():
        if segment and seg != segment:
            continue
        if stone and stone_name != stone:
            continue
        for product, vals in products.items():
            qty = int(vals.get("qty", 0))
            amount = float(vals.get("amount", 0))
            if qty == 0 and amount == 0:
                continue
            rows.append({
                "Сегмент": SEGMENT_LABELS.get(seg, seg),
                "Камень": stone_name,
                "Номенклатурная группа": PRODUCT_LABELS.get(product, product),
                "Код группы": product,
                "Количество": qty,
                "Выручка": amount,
                "% количества магазина": qty / store.total_qty if store.total_qty else 0,
                "% выручки магазина": amount / store.total_amount if store.total_amount else 0,
                "% количества камня": qty / stone_qty if stone_qty else 0,
                "% выручки камня": amount / stone_amount if stone_amount else 0,
                "Средняя стоимость": amount / qty if qty else 0,
            })
    if not rows:
        return pd.DataFrame(columns=[
            "Сегмент", "Камень", "Номенклатурная группа", "Код группы", "Количество", "Выручка",
            "% количества магазина", "% выручки магазина", "% количества камня", "% выручки камня",
            "Средняя стоимость",
        ])
    order_map = {PRODUCT_LABELS.get(p, p): idx for idx, p in enumerate(PRODUCT_ORDER)}
    df = pd.DataFrame(rows)
    df["_order"] = df["Номенклатурная группа"].map(order_map).fillna(999)
    return df.sort_values(["_order", "Номенклатурная группа"]).drop(columns="_order")


def cross_store_product_dataframe(stores: list, segment: str, stone: str, product_label: str) -> pd.DataFrame:
    rows = []
    for store in stores:
        df = product_dataframe(store, segment, stone)
        selected = df[df["Номенклатурная группа"] == product_label]
        qty = int(selected["Количество"].sum()) if not selected.empty else 0
        amount = float(selected["Выручка"].sum()) if not selected.empty else 0
        rows.append({
            "Магазин": base_store_name(store.name),
            "Количество": qty,
            "Выручка": amount,
            "Средняя стоимость": amount / qty if qty else 0,
            "% количества магазина": qty / store.total_qty if store.total_qty else 0,
            "% выручки магазина": amount / store.total_amount if store.total_amount else 0,
        })
    return pd.DataFrame(rows)


def formatted_table(df: pd.DataFrame) -> pd.DataFrame:
    display = df.copy()
    for col in [c for c in display.columns if c.startswith("%")]:
        display[col] = display[col].map(pct)
    for col in [c for c in ["Количество", "Выручка", "Средняя стоимость"] if c in display.columns]:
        display[col] = display[col].map(money)
    if "Код группы" in display.columns:
        display = display.drop(columns="Код группы")
    return display


def section_divider(title: str, subtitle: str = "", kicker: str = "ANALITIKA") -> None:
    st.markdown(
        f'<div class="section-divider">'
        f'<div class="section-divider-kicker">{kicker}</div>'
        f'<div class="section-divider-title">{title}</div>'
        f'<div class="section-divider-copy">{subtitle}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def insight_panel(title: str, lines: list[str]) -> None:
    clean = [line for line in lines if line]
    if not clean:
        return
    body = "".join(f'<div class="analysis-line">• {line}</div>' for line in clean)
    st.markdown(
        f'<div class="analysis-panel"><div class="analysis-panel-title">{title}</div>{body}</div>',
        unsafe_allow_html=True,
    )


def network_conclusions(summary_df: pd.DataFrame) -> list[str]:
    if summary_df.empty:
        return []
    lines: list[str] = []
    leader = summary_df.sort_values("Выручка", ascending=False).iloc[0]
    lines.append(f"Лидер сети по выручке — {leader['Магазин']}: {money(leader['Выручка'])} VND.")
    qty_leader = summary_df.sort_values("Количество", ascending=False).iloc[0]
    lines.append(f"Больше всего изделий продано в {qty_leader['Магазин']} — {money(qty_leader['Количество'])} шт.")
    segment_sales = {}
    for seg in SEG_ORDER:
        col = f"{SEGMENT_LABELS[seg]} — продажи %"
        if col in summary_df.columns:
            segment_sales[SEGMENT_LABELS[seg]] = float(summary_df[col].mean())
    if segment_sales:
        seg_name, seg_share = max(segment_sales.items(), key=lambda item: item[1])
        lines.append(f"Доминирующий сегмент сети — {seg_name}: в среднем {pct(seg_share)} выручки магазинов.")
    return lines[:4]


def interactive_conclusions(store, segment: str, stone: str, product_df: pd.DataFrame, selected_product: str, qty: int, sales: float) -> list[str]:
    lines: list[str] = []
    lines.append(f"Текущий фильтр: {base_store_name(store.name)} → {SEGMENT_LABELS[segment]} → {stone}.")
    if selected_product != "Все номенклатурные группы":
        lines.append(f"Группа «{selected_product}» формирует {pct(sales / store.total_amount if store.total_amount else 0)} выручки магазина и {pct(qty / store.total_qty if store.total_qty else 0)} количества.")
    elif not product_df.empty:
        top = product_df.sort_values("Выручка", ascending=False).iloc[0]
        lines.append(f"Лидер по выручке внутри {stone} — «{top['Номенклатурная группа']}»: {pct(float(top['% выручки камня']))}.")
        if len(product_df) > 1:
            low = product_df[product_df["Количество"] > 0].sort_values("Выручка", ascending=True)
            if not low.empty:
                row = low.iloc[0]
                lines.append(f"Минимальная представленность — «{row['Номенклатурная группа']}»: {money(row['Количество'])} шт.")
    avg = sales / qty if qty else 0
    if avg:
        lines.append(f"Средняя стоимость в выбранном срезе — {money(avg)} VND.")
    return lines[:4]


def supplier_conclusions(df: pd.DataFrame, summary: pd.DataFrame) -> list[str]:
    if df.empty or summary.empty:
        return []
    lines: list[str] = []
    leader = summary.iloc[0]
    lines.append(f"Лидер среди поставщиков — {leader['Поставщик']}: {pct(float(leader['% выручки']))} общей выручки.")
    qty_leader = summary.sort_values("Количество", ascending=False).iloc[0]
    lines.append(f"По количеству лидирует {qty_leader['Поставщик']} — {money(qty_leader['Количество'])} шт.")
    if "Магазин" in df.columns:
        coverage = df.groupby("Поставщик")["Магазин"].nunique().sort_values(ascending=False)
        if not coverage.empty:
            lines.append(f"Самое широкое покрытие у {coverage.index[0]} — {int(coverage.iloc[0])} магазинов.")
    return lines[:4]


def conclusions(store, all_stores: list) -> list[str]:
    lines: list[str] = []
    seg = segment_totals(store)
    if store.total_amount:
        leader = max(SEG_ORDER, key=lambda x: seg[x]["amount"])
        share = seg[leader]["amount"] / store.total_amount
        lines.append(f"Основную выручку формирует {SEGMENT_LABELS[leader]} — {pct(share)}.")
    network_avg = sum(s.total_amount for s in all_stores) / max(1, sum(s.total_qty for s in all_stores))
    store_avg = store.total_amount / store.total_qty if store.total_qty else 0
    if network_avg:
        delta = store_avg / network_avg - 1
        direction = "выше" if delta >= 0 else "ниже"
        lines.append(f"Средняя стоимость изделия {direction} средней по сети на {pct(abs(delta))}.")
    top_stones = [(stone, totals_for(store, "TOP STONES", stone)[1]) for stone in TOP_ORDER]
    top_stones = [x for x in top_stones if x[1] > 0]
    if top_stones:
        name, amount = max(top_stones, key=lambda x: x[1])
        top_total = seg["TOP STONES"]["amount"]
        lines.append(f"Лидер внутри Top Stones — {name}: {pct(amount / top_total if top_total else 0)} выручки сегмента.")
        products = product_dataframe(store, "TOP STONES", name)
        if not products.empty:
            product = products.sort_values("Выручка", ascending=False).iloc[0]
            lines.append(
                f"В {name} основную выручку дает группа «{product['Номенклатурная группа']}» — "
                f"{pct(float(product['% выручки камня']))}."
            )
    return lines[:4]


def interactive_explorer(store, all_stores: list, namespace: str = "interactive") -> None:
    st.caption("Выберите сегмент → камень → номенклатурную группу. Данные и диаграммы перестроятся сразу.")

    f1, f2, f3 = st.columns(3)
    with f1:
        selected_segment = st.selectbox(
            "Сегмент",
            SEG_ORDER,
            format_func=lambda s: SEGMENT_LABELS[s],
            key=f"{namespace}_segment_{base_store_name(store.name)}",
        )

    available_stones = [
        stone for stone in STONE_ORDERS[selected_segment]
        if totals_for(store, selected_segment, stone)[0] or totals_for(store, selected_segment, stone)[1]
    ]
    if not available_stones:
        available_stones = STONE_ORDERS[selected_segment]
    with f2:
        selected_stone = st.selectbox(
            "Камень / группа камней",
            available_stones,
            key=f"{namespace}_stone_{base_store_name(store.name)}",
        )

    product_df = product_dataframe(store, selected_segment, selected_stone)
    product_options = ["Все номенклатурные группы"] + product_df["Номенклатурная группа"].drop_duplicates().tolist()
    with f3:
        selected_product = st.selectbox(
            "Номенклатурная группа",
            product_options,
            key=f"{namespace}_product_{base_store_name(store.name)}",
        )

    if product_df.empty:
        st.info("В выбранной группе нет продаж за этот период.")
        return

    stone_qty, stone_sales = totals_for(store, selected_segment, selected_stone)
    if selected_product == "Все номенклатурные группы":
        selected_qty = stone_qty
        selected_sales = stone_sales
        context_note = f"Итого по {selected_stone}"
    else:
        selected_rows = product_df[product_df["Номенклатурная группа"] == selected_product]
        selected_qty = int(selected_rows["Количество"].sum())
        selected_sales = float(selected_rows["Выручка"].sum())
        context_note = f"{selected_stone} → {selected_product}"

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1: kpi_card("Количество", f"{money(selected_qty)} шт.", context_note)
    with k2: kpi_card("Выручка", f"{money(selected_sales)} VND", context_note)
    with k3: kpi_card("Средняя стоимость", f"{money(selected_sales / selected_qty if selected_qty else 0)} VND")
    with k4: kpi_card("% количества магазина", pct(selected_qty / store.total_qty if store.total_qty else 0))
    with k5: kpi_card("% выручки магазина", pct(selected_sales / store.total_amount if store.total_amount else 0))

    if selected_product == "Все номенклатурные группы":
        left, right = st.columns(2)
        with left:
            locked_plotly_chart(
                horizontal_bar(product_df, "Номенклатурная группа", "Количество", f"{selected_stone}: количество по группам", " шт."),
                width="stretch",
            )
        with right:
            locked_plotly_chart(
                horizontal_bar(product_df, "Номенклатурная группа", "Выручка", f"{selected_stone}: выручка по группам"),
                width="stretch",
            )
        st.dataframe(formatted_table(product_df), width="stretch", hide_index=True)
    else:
        comparison = cross_store_product_dataframe(all_stores, selected_segment, selected_stone, selected_product)
        left, right = st.columns(2)
        with left:
            locked_plotly_chart(
                horizontal_bar(comparison, "Магазин", "Количество", f"{selected_product}: количество по магазинам", " шт."),
                width="stretch",
            )
        with right:
            locked_plotly_chart(
                horizontal_bar(comparison, "Магазин", "Выручка", f"{selected_product}: выручка по магазинам"),
                width="stretch",
            )
        st.markdown("#### Сравнение выбранной группы по сети")
        st.dataframe(formatted_table(comparison), width="stretch", hide_index=True)

    insight_panel(
        "Аналитика по выбранным параметрам",
        interactive_conclusions(store, selected_segment, selected_stone, product_df, selected_product, selected_qty, selected_sales),
    )


def store_view(store, all_stores: list) -> None:
    st.markdown(f'<div class="section-title">Магазин {base_store_name(store.name)}</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi_card("Выручка", f"{money(store.total_amount)} VND")
    with c2: kpi_card("Продано изделий", money(store.total_qty) + " шт.")
    with c3: kpi_card("Средняя стоимость", f"{money(store.total_amount / store.total_qty if store.total_qty else 0)} VND")
    with c4:
        network_sales = sum(s.total_amount for s in all_stores)
        kpi_card("Доля в выручке сети", pct(store.total_amount / network_sales if network_sales else 0))

    seg = segment_totals(store)
    labels = [SEGMENT_LABELS[s] for s in SEG_ORDER]
    colors = [SEGMENT_COLORS[s] for s in SEG_ORDER]
    a, b = st.columns(2)
    with a:
        locked_plotly_chart(
            donut(labels, [seg[s]["amount"] for s in SEG_ORDER], "Структура продаж", colors),
            width="stretch",
            key=f"store_sales_structure_{base_store_name(store.name)}",
        )
    with b:
        locked_plotly_chart(
            donut(labels, [seg[s]["qty"] for s in SEG_ORDER], "Структура количества", colors),
            width="stretch",
            key=f"store_qty_structure_{base_store_name(store.name)}",
        )

    detail_options = ["Все камни", "Все номенклатурные группы", "Top Stones", "Pearls", "Colored Stones"]
    detail_mode = st.segmented_control(
        "Детализация магазина",
        detail_options,
        default="Все камни",
        key="store_detail_mode",
    ) or "Все камни"

    data = stone_dataframe(store)
    if detail_mode == "Все камни":
        st.dataframe(formatted_table(data), width="stretch", hide_index=True)
    elif detail_mode == "Все номенклатурные группы":
        st.dataframe(formatted_table(product_dataframe(store)), width="stretch", hide_index=True)
    else:
        segment_lookup = {
            "Top Stones": "TOP STONES",
            "Pearls": "PEARLS",
            "Colored Stones": "COLORED STONES",
        }
        seg_code = segment_lookup[detail_mode]
        subset = data[data["Сегмент"] == detail_mode]
        x1, x2 = st.columns(2)
        with x1:
            locked_plotly_chart(
                donut(subset["Камень"].tolist(), subset["Количество"].tolist(), f"{detail_mode}: количество"),
                width="stretch",
                key=f"store_detail_qty_{base_store_name(store.name)}_{seg_code}",
            )
        with x2:
            locked_plotly_chart(
                donut(subset["Камень"].tolist(), subset["Выручка"].tolist(), f"{detail_mode}: выручка"),
                width="stretch",
                key=f"store_detail_sales_{base_store_name(store.name)}_{seg_code}",
            )
        st.markdown("#### Номенклатурные группы сегмента")
        st.dataframe(formatted_table(product_dataframe(store, seg_code)), width="stretch", hide_index=True)

    if base_store_name(store.name) == "OUTLET" and store.extras:
        st.markdown("### Дополнительные подразделения OUTLET")
        cols = st.columns(2)
        for idx, name in enumerate(["GIFT TT", "CAFE"]):
            values = store.extras.get(name, {"qty": 0, "amount": 0})
            avg = values["amount"] / values["qty"] if values["qty"] else 0
            with cols[idx]:
                st.markdown(f"**{name}**")
                a, b, c = st.columns(3)
                with a: kpi_card("Выручка", f"{money(values['amount'])} VND")
                with b: kpi_card("Количество", f"{money(values['qty'])} шт.")
                with c: kpi_card("Средняя стоимость", f"{money(avg)} VND")

    insight_panel("Аналитика по магазину", conclusions(store, all_stores))


def is_supplier_report(path: Path) -> bool:
    """Detect the supplier hierarchy export by its first header rows."""
    wb = load_workbook(path, data_only=True, read_only=False)
    try:
        ws = wb.active
        header = " ".join(str(ws.cell(r, 1).value or "") for r in range(1, 7)).upper()
        return "ПОСТАВЩИК" in header and "НОМЕНКЛАТУРНАЯ ГРУППА" in header
    finally:
        wb.close()


def parse_supplier_report_with_period(path: Path) -> tuple[pd.DataFrame, tuple | None]:
    """Parse supplier detail and period in one workbook pass.

    The previous implementation reopened the same workbook several times during
    one upload. On Community Cloud that created avoidable memory spikes. This
    function loads the workbook once, extracts both the hierarchy and period,
    and closes it deterministically before returning.
    """
    wb = load_workbook(path, data_only=True, read_only=False)
    rows: list[dict] = []
    try:
        ws = wb.active
        period = extract_period(ws)
        current_store: str | None = None
        current_stone: str | None = None
        current_product: str | None = None
        skip_store_section = False
        has_store_dimension = "МАГАЗИН" in str(ws.cell(4, 1).value or "").upper()

        for row in range(7, ws.max_row + 1):
            cell = ws.cell(row, 1)
            value = cell.value
            if value is None:
                continue
            text = str(value).strip()
            if not text:
                continue
            indent = float(cell.alignment.indent or 0)
            upper = text.upper()

            if upper in {"ИТОГО", "ИТОГО:", "ПОСТАВЩИКИ"} or upper.startswith("ОТЧЕТ"):
                continue

            if has_store_dimension and indent == 0 and cell.font.bold:
                normalized = normalize_store_from_report(text)
                current_store = normalized
                skip_store_section = normalized is None
                current_stone = None
                current_product = None
                continue

            if has_store_dimension and skip_store_section:
                continue

            if (has_store_dimension and indent == 2) or (not has_store_dimension and indent == 0 and not cell.font.bold):
                current_stone = text
                current_product = None
                continue

            if current_stone and ((has_store_dimension and indent == 4) or (not has_store_dimension and indent == 2)):
                current_product = norm_product(text)
                continue

            supplier_indent = 7 if has_store_dimension else 5
            is_supplier = current_stone and current_product and indent >= supplier_indent and not cell.font.bold
            if not is_supplier:
                continue

            qty = int(round(float(ws.cell(row, 8).value or 0)))
            amount = float(ws.cell(row, 9).value or 0)
            if qty == 0 and amount == 0:
                continue
            segment, stone, rule = classify(current_stone)
            supplier_name = text.strip()
            if supplier_name.upper() in {"", "СЕТЬ", "NETWORK", "NONE", "NAN", "UNKNOWN", "НЕ УКАЗАН", "БЕЗ ПОСТАВЩИКА"}:
                supplier_name = "Other"
            rows.append({
                "Магазин": current_store if has_store_dimension else "Сеть",
                "Поставщик": supplier_name,
                "Сегмент": SEGMENT_LABELS.get(segment, segment),
                "Код сегмента": segment,
                "Камень": stone,
                "Исходный камень": current_stone,
                "Номенклатурная группа": PRODUCT_LABELS.get(current_product, current_product),
                "Код группы": current_product,
                "Количество": qty,
                "Выручка": amount,
                "Правило": rule,
            })
    finally:
        wb.close()

    columns = [
        "Магазин", "Поставщик", "Сегмент", "Код сегмента", "Камень",
        "Исходный камень", "Номенклатурная группа", "Код группы",
        "Количество", "Выручка", "Правило",
    ]
    return pd.DataFrame(rows, columns=columns), period


def parse_supplier_report(path: Path) -> pd.DataFrame:
    detail, _ = parse_supplier_report_with_period(path)
    return detail


def supplier_units_from_detail(
    detail: pd.DataFrame,
    period: tuple | None,
    file_name: str,
) -> dict[str, StoreData]:
    """Convert already-parsed supplier rows into StoreData without reopening Excel."""
    if detail.empty:
        return {}

    stores: dict[str, StoreData] = {}
    touched: set[str] = set()
    for row in detail.to_dict("records"):
        store_name = str(row["Магазин"])
        if store_name in {"GIFT TT", "CAFE"}:
            outlet = stores.setdefault("OUTLET", StoreData("OUTLET"))
            outlet.extras[store_name]["qty"] += int(row["Количество"])
            outlet.extras[store_name]["amount"] += float(row["Выручка"])
            touched.add("OUTLET")
            continue
        store = stores.setdefault(store_name, StoreData(store_name))
        touched.add(store_name)
        store.add(
            row["Код сегмента"], row["Камень"], row["Код группы"],
            int(row["Количество"]), float(row["Выручка"]),
            str(row["Исходный камень"]), str(row["Правило"]),
        )
    for name in touched:
        stores[name].add_period(period, file_name)
    return stores


def supplier_report_units(path: Path) -> dict[str, StoreData]:
    detail, period = parse_supplier_report_with_period(path)
    return supplier_units_from_detail(detail, period, path.name)

def supplier_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Поставщик", "Количество", "Выручка", "Средняя стоимость", "% количества", "% выручки"])
    df = df.copy()
    df["Поставщик"] = df["Поставщик"].fillna("Other").astype(str).str.strip()
    df.loc[df["Поставщик"].str.upper().isin({"", "СЕТЬ", "NETWORK", "NONE", "NAN", "UNKNOWN", "НЕ УКАЗАН", "БЕЗ ПОСТАВЩИКА"}), "Поставщик"] = "Other"
    result = df.groupby("Поставщик", as_index=False).agg(
        Количество=("Количество", "sum"),
        Выручка=("Выручка", "sum"),
    )
    total_qty = float(result["Количество"].sum())
    total_sales = float(result["Выручка"].sum())
    result["Средняя стоимость"] = result["Выручка"] / result["Количество"].replace(0, pd.NA)
    result["Средняя стоимость"] = result["Средняя стоимость"].fillna(0)
    result["% количества"] = result["Количество"] / total_qty if total_qty else 0
    result["% выручки"] = result["Выручка"] / total_sales if total_sales else 0
    return result.sort_values("Выручка", ascending=False)


SUPPLIER_PIE_MIN_SHARE = 0.045


def supplier_pie_data(summary: pd.DataFrame, share_col: str) -> tuple[list[str], list[float]]:
    """Collapse suppliers below 4.5% into Other for supplier pie charts only.

    The detailed table and horizontal charts keep the original suppliers unchanged.
    Each pie is grouped independently by its own metric (revenue or quantity).
    """
    if summary.empty or share_col not in summary.columns:
        return [], []

    labels: list[str] = []
    values: list[float] = []
    other_value = 0.0

    for _, row in summary.iterrows():
        label = str(row["Поставщик"]).strip() or "Other"
        value = float(row[share_col])
        if label.casefold() == "other" or value < SUPPLIER_PIE_MIN_SHARE:
            other_value += value
        else:
            labels.append(label)
            values.append(value)

    if other_value > 0:
        labels.append("Other")
        values.append(other_value)

    return labels, values


def supplier_view(df: pd.DataFrame) -> None:
    st.caption("Общая аналитика по сети из выгрузки «Камень → Номенклатурная группа → Поставщик».")
    if df.empty:
        st.info("Загрузите выгрузку с поставщиками на странице «Главная».")
        return

    summary = supplier_summary(df)
    total_qty = int(df["Количество"].sum())
    total_sales = float(df["Выручка"].sum())
    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi_card("Поставщиков", str(summary["Поставщик"].nunique()))
    with c2: kpi_card("Продано изделий", f"{money(total_qty)} шт.")
    with c3: kpi_card("Выручка", f"{money(total_sales)} VND")
    with c4: kpi_card("Средняя стоимость", f"{money(total_sales / total_qty if total_qty else 0)} VND")

    revenue_labels, revenue_values = supplier_pie_data(summary, "% выручки")
    quantity_labels, quantity_values = supplier_pie_data(summary, "% количества")

    left, right = st.columns(2)
    with left:
        locked_plotly_chart(donut(revenue_labels, revenue_values, "Доля поставщиков по выручке"), width="stretch")
    with right:
        locked_plotly_chart(donut(quantity_labels, quantity_values, "Доля поставщиков по количеству"), width="stretch")

    left2, right2 = st.columns(2)
    with left2:
        locked_plotly_chart(horizontal_bar(summary.head(15), "Поставщик", "Выручка", "Топ поставщиков по выручке"), width="stretch")
    with right2:
        locked_plotly_chart(horizontal_bar(summary.head(15), "Поставщик", "Количество", "Топ поставщиков по количеству", " шт."), width="stretch")

    st.markdown("### Общая таблица поставщиков")
    st.dataframe(formatted_table(summary), width="stretch", hide_index=True)

    supplier_names = summary["Поставщик"].tolist()
    selected = st.selectbox("Выберите поставщика", supplier_names, key="supplier_selected")
    detail = df[df["Поставщик"] == selected].copy()
    selected_qty = int(detail["Количество"].sum())
    selected_sales = float(detail["Выручка"].sum())
    a, b, c, d = st.columns(4)
    with a: kpi_card("Поставщик", selected)
    with b: kpi_card("Количество", f"{money(selected_qty)} шт.")
    with c: kpi_card("Выручка", f"{money(selected_sales)} VND")
    with d: kpi_card("Средняя стоимость", f"{money(selected_sales / selected_qty if selected_qty else 0)} VND")

    by_segment = detail.groupby("Сегмент", as_index=False).agg(
        Количество=("Количество", "sum"), Выручка=("Выручка", "sum")
    ).sort_values("Выручка", ascending=False)
    by_product = detail.groupby("Номенклатурная группа", as_index=False).agg(
        Количество=("Количество", "sum"), Выручка=("Выручка", "sum")
    ).sort_values("Выручка", ascending=False)
    by_stone = detail.groupby(["Сегмент", "Камень"], as_index=False).agg(
        Количество=("Количество", "sum"), Выручка=("Выручка", "sum")
    ).sort_values("Выручка", ascending=False)
    by_store = detail.groupby("Магазин", as_index=False).agg(
        Количество=("Количество", "sum"), Выручка=("Выручка", "sum")
    ).sort_values("Выручка", ascending=False) if "Магазин" in detail.columns else pd.DataFrame()

    if not by_store.empty and by_store["Магазин"].nunique() > 1:
        st.markdown("#### По магазинам")
        locked_plotly_chart(horizontal_bar(by_store, "Магазин", "Выручка", f"{selected}: выручка по магазинам"), width="stretch")
        st.dataframe(formatted_table(by_store), width="stretch", hide_index=True)

    seg_l, seg_r = st.columns(2)
    with seg_l:
        locked_plotly_chart(donut(by_segment["Сегмент"].tolist(), by_segment["Выручка"].tolist(), f"{selected}: сегменты по выручке"), width="stretch")
    with seg_r:
        locked_plotly_chart(donut(by_segment["Сегмент"].tolist(), by_segment["Количество"].tolist(), f"{selected}: сегменты по количеству"), width="stretch")

    l, r = st.columns(2)
    with l:
        locked_plotly_chart(horizontal_bar(by_product, "Номенклатурная группа", "Выручка", f"{selected}: номенклатурные группы"), width="stretch")
    with r:
        locked_plotly_chart(horizontal_bar(by_stone.head(20), "Камень", "Выручка", f"{selected}: камни"), width="stretch")

    table_mode = st.segmented_control(
        "Таблица детализации поставщика",
        ["Сегменты", "Номенклатурные группы", "Камни", "Полная детализация"],
        default="Сегменты",
        key="supplier_table_mode",
    ) or "Сегменты"
    if table_mode == "Сегменты":
        table_df = by_segment
    elif table_mode == "Номенклатурные группы":
        table_df = by_product
    elif table_mode == "Камни":
        table_df = by_stone
    else:
        table_df = detail
    st.dataframe(formatted_table(table_df), width="stretch", hide_index=True)

    if "Магазин" in df.columns and df["Магазин"].nunique() > 1:
        st.caption("Доступен полный разрез: поставщик × магазин × камень × номенклатурная группа.")

    insight_panel("Аналитика по поставщикам", supplier_conclusions(df, summary))


def _merge_units(target: dict[str, StoreData], incoming: dict[str, StoreData]) -> None:
    for name, source in incoming.items():
        dest = target.setdefault(name, StoreData(name))
        dest.periods.extend(source.periods)
        dest.files.extend(source.files)
        for (segment, stone), products in source.data.items():
            for product, vals in products.items():
                dest.add(segment, stone, product, int(vals.get("qty", 0)), float(vals.get("amount", 0)), stone, "merged")
        for extra_name, vals in source.extras.items():
            dest.extras[extra_name]["qty"] += int(vals.get("qty", 0))
            dest.extras[extra_name]["amount"] += float(vals.get("amount", 0))


def parse_uploads(uploaded_files):
    """Parse all uploaded workbooks once and return stores, errors and supplier detail."""
    errors: list[tuple[str, str]] = []
    supplier_frames: list[pd.DataFrame] = []
    stores: dict[str, StoreData] = {}

    with tempfile.TemporaryDirectory(prefix="analitika_parse_") as temp_dir:
        normal_paths: list[Path] = []
        supplier_paths: list[Path] = []

        for uploaded in uploaded_files:
            path = Path(temp_dir) / uploaded.name
            path.write_bytes(uploaded.getvalue())
            try:
                if is_supplier_report(path):
                    supplier_paths.append(path)
                else:
                    normal_paths.append(path)
            except Exception as exc:
                errors.append((uploaded.name, str(exc)))

        if normal_paths:
            normal_stores, normal_errors = build_report_units(normal_paths)
            _merge_units(stores, normal_stores)
            errors.extend(normal_errors)

        for path in supplier_paths:
            try:
                detail, period = parse_supplier_report_with_period(path)
                if not detail.empty:
                    supplier_frames.append(detail)
                    _merge_units(stores, supplier_units_from_detail(detail, period, path.name))
            except Exception as exc:
                errors.append((path.name, str(exc)))

    if supplier_frames:
        supplier_df = pd.concat(supplier_frames, ignore_index=True, copy=False)
        supplier_df["Поставщик"] = supplier_df["Поставщик"].fillna("Other").astype(str).str.strip()
        service_values = {"", "СЕТЬ", "NETWORK", "NONE", "NAN", "UNKNOWN", "НЕ УКАЗАН", "БЕЗ ПОСТАВЩИКА"}
        supplier_df.loc[supplier_df["Поставщик"].str.upper().isin(service_values), "Поставщик"] = "Other"
    else:
        supplier_df = pd.DataFrame()

    return stores, errors, supplier_df


def cache_payloads(uploaded_files: list[StoredUpload]) -> tuple[tuple[str, bytes], ...]:
    """Immutable payload accepted by the shared report cache."""
    return tuple((item.name, item.getvalue()) for item in uploaded_files)


@st.cache_resource(ttl=1800, max_entries=2, show_spinner=False)
def parse_report_bundle(payloads: tuple[tuple[str, bytes], ...]):
    """Share one parsed report between reruns and open browser sessions.

    Parsed StoreData objects are treated as read-only by the UI. A resource cache
    avoids keeping separate copies of the same report in every iPad/PC session.
    TTL and max_entries bound memory use on the long-running Cloud process.
    """
    uploads = [StoredUpload(name, data) for name, data in payloads]
    return parse_uploads(uploads)


def sidebar_navigation(has_report: bool) -> None:
    """Render the compact, luxury-styled sidebar navigation."""
    items = [
        ("#upload", "⬆️ Загрузка"),
    ]
    if has_report:
        items.extend([
            ("#executive", "⚡ Оперативная сводка"),
            ("#summary", "📊 Сводка"),
            ("#stores", "🏪 Магазины"),
            ("#interactive", "🔎 Интерактивная аналитика"),
            ("#suppliers", "📦 Поставщики"),
        ])
    items.append(("#about", "ℹ️ О платформе"))
    links = "".join(f'<a href="{href}">{label}</a>' for href, label in items)

    with st.sidebar:
        logo = Path(__file__).parent / "assets" / "logo.png"
        if logo.exists():
            st.image(str(logo), width="stretch")
        st.markdown("---")
        st.markdown("**Princess Jewelry Analytics**")
        st.caption(f"Analitika Web {APP_VERSION}")
        st.markdown('<div class="nav-hint">Навигация по странице</div>', unsafe_allow_html=True)
        st.markdown(f'<nav class="side-nav">{links}</nav>', unsafe_allow_html=True)
        if has_report:
            st.markdown("---")
            st.success("Отчет загружен")
            if st.button("Загрузить другой отчет", width="stretch", key="replace_report"):
                clear_saved_uploads()
                st.rerun()
        st.markdown("---")
        st.caption("Разработка: Vladimir Panasyan")


def mobile_navigation(has_report: bool) -> None:
    """Compact sticky navigation shown only on iPad/phone via CSS media queries."""
    items = [("#upload", "⬆️ Загрузка")]
    if has_report:
        items.extend([
            ("#executive", "⚡ Для руководителя"),
            ("#summary", "📊 Сводка"),
            ("#stores", "🏪 Магазины"),
            ("#interactive", "🔎 Аналитика"),
            ("#suppliers", "📦 Поставщики"),
        ])
    items.append(("#about", "ℹ️ О платформе"))
    links = "".join(f'<a href="{href}">{label}</a>' for href, label in items)
    st.markdown(
        f'<div class="mobile-nav-shell"><nav class="mobile-nav">{links}</nav></div>',
        unsafe_allow_html=True,
    )


def render_about() -> None:
    st.markdown('<div id="about"></div>', unsafe_allow_html=True)
    section_divider(
        'О платформе',
        'Как подготовить выгрузку, что показывает отчет и что изменилось в последних версиях.',
        f'ANALITIKA WEB {APP_VERSION}',
    )
    st.markdown(
        """
        <div class="about-grid">
          <div class="about-card">
            <h4>Как подготовить отчет в 1С</h4>
            <div class="about-step"><b>1.</b> Откройте отчет <b>«Продажи товаров»</b>.</div>
            <div class="about-step"><b>2.</b> Выберите период для анализа.</div>
            <div class="about-step"><b>3.</b> Включите уровни: <b>Магазин</b>, <b>Номенклатурная группа</b>, <b>Камень / вставка</b>, <b>Поставщик</b>.</div>
            <div class="about-step"><b>4.</b> Сохраните результат в Excel и загрузите его в Analitika. Название файла может быть любым.</div>
          </div>
          <div class="about-card">
            <h4>Оперативная сводка</h4>
            <p>Компактный режим для руководителя: ключевые показатели сети, лидеры, структура продаж, концентрация по магазинам и поставщикам, а также короткие фактические выводы.</p>
          </div>
          <div class="about-card">
            <h4>Сводка</h4>
            <p>Общие показатели сети, доли магазинов, структура Top Stones / Pearls / Colored Stones, сравнительные диаграммы и основные выводы.</p>
          </div>
          <div class="about-card">
            <h4>Магазины</h4>
            <p>Выручка и количество по выбранному магазину, камни, номенклатурные группы, диаграммы сегментов и отдельный блок OUTLET с GIFT TT и CAFE.</p>
          </div>
          <div class="about-card">
            <h4>Интерактивная аналитика</h4>
            <p>Фильтры по магазину, сегменту, камню и товарной группе. Таблицы, диаграммы и выводы перестраиваются под выбранные параметры.</p>
          </div>
          <div class="about-card">
            <h4>Поставщики</h4>
            <p>Сравнение поставщиков по штукам и выручке, а также разрез выбранного поставщика по магазинам, сегментам, камням и номенклатурным группам.</p>
          </div>
          <div class="about-card">
            <h4>Обновления</h4>
            <div class="about-step"><b>Analitika Web 1.1.9 — Retail leader correction</b><br>В оперативной сводке лидеры по выручке и количеству теперь рассчитываются только по розничной сети: OUTLET и 63 исключены из рейтинга как магазины с отдельной туристической моделью трафика. Во всех диаграммах и таблицах их данные сохранены.</div>
            <div class="about-step"><b>Analitika Web 1.1.8 — Executive operational brief</b><br>Добавлен отдельный iPad-friendly блок для руководителя: сеть в одном экране, лидеры по магазинам, структура сегментов, ключевые концентрации и компактная сводка по поставщикам. Детальные разделы сохранены ниже без изменений.</div>
            <div class="about-step"><b>Analitika Web 1.1.7 — Stability and memory optimization</b><br>Обработка Excel выполняется один раз и переиспользуется между сессиями, фильтры обновляют только свой блок, а скрытые вкладки больше не создают лишние таблицы и диаграммы. Добавлены ограниченный кэш и принудительное освобождение временных объектов.</div>
            <div class="about-step"><b>Analitika Web 1.1.6 — Responsive mobile layout</b><br>Интерфейс адаптирован под iPad и смартфоны: добавлена мобильная навигация, KPI и фильтры перестраиваются под ширину экрана, парные диаграммы складываются в одну колонку в портретном режиме, а таблицы сохраняют сортировку и горизонтальную прокрутку.</div>
            <div class="about-step"><b>Analitika Web 1.1.5 — Locked chart interactions</b><br>Диаграммы переведены в режим просмотра: отключены масштабирование, перетаскивание, выделение, изменение легенды и панель сохранения. Подсказки по наведению на ПК и касанию на iPad сохранены; таблицы остаются интерактивными.</div>
            <div class="about-step"><b>Analitika Web 1.1.4 — Release history</b><br>Вместо изменяемых планов в разделе «О платформе» теперь отображается история фактических обновлений.</div>
            <div class="about-step"><b>Analitika Web 1.1.3 — Group small suppliers in pie charts</b><br>Поставщики с долей ниже 4,5% объединяются в Other только на круговых диаграммах. Полная детализация остается на линейных диаграммах ниже.</div>
            <div class="about-step"><b>Analitika Web 1.1.2 — Fix chart label clipping</b><br>Увеличены рабочие поля диаграмм, исправлено обрезание выносок и крупных значений.</div>
            <div class="about-step"><b>Analitika Web 1.1.1 — Cloud stability hotfix</b><br>Стабилизирован запуск в Streamlit Cloud, зафиксированы зависимости и оптимизировано повторное чтение отчета.</div>
            <div class="about-step"><b>Analitika Web 1.1.0 — Production release</b><br>Запущена производственная версия с одной загрузкой, навигацией по разделам и модулем поставщиков.</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption(f"Analitika Web {APP_VERSION} · Princess Jewelry · Developed by Vladimir Panasyan")



@st.fragment
def render_store_fragment(stores: list[StoreData]) -> None:
    """Rerun only the store block when its controls change."""
    try:
        store_names = [base_store_name(store.name) for store in stores]
        chosen = st.selectbox("Выберите магазин", store_names, index=0, key="store_page_select")
        chosen_store = next(store for store in stores if base_store_name(store.name) == chosen)
        store_view(chosen_store, stores)
    finally:
        gc.collect()


@st.fragment
def render_interactive_fragment(stores: list[StoreData]) -> None:
    """Isolate interactive filters from the rest of the dashboard."""
    try:
        store_names = [base_store_name(store.name) for store in stores]
        chosen_interactive = st.selectbox(
            "Магазин для интерактивного анализа",
            store_names,
            index=0,
            key="interactive_store_select",
        )
        interactive_store = next(
            store for store in stores if base_store_name(store.name) == chosen_interactive
        )
        interactive_explorer(interactive_store, stores, namespace="main_interactive")
    finally:
        gc.collect()


@st.fragment
def render_supplier_fragment(supplier_df: pd.DataFrame) -> None:
    """Rerun supplier controls without rebuilding summary and store charts."""
    try:
        if supplier_df.empty:
            st.info("В загруженном файле нет детализации по поставщикам.")
        else:
            supplier_view(supplier_df)
    finally:
        gc.collect()

def main() -> None:
    st.markdown('<div id="upload"></div>', unsafe_allow_html=True)
    st.markdown(
        '<section class="luxury-hero">'
        '<div class="luxury-hero-content">'
        '<div class="luxury-eyebrow">Princess Jewelry · Internal Analytics</div>'
        '<div class="luxury-title">Данные, которые<br><span>помогают решать</span></div>'
        '<div class="luxury-divider"></div>'
        '<div class="luxury-copy">Загрузите общую выгрузку продаж. Analitika автоматически определит магазины, сегменты, камни, товарные группы и поставщиков.</div>'
        '<div class="luxury-badges"><span class="luxury-badge">Одна загрузка</span><span class="luxury-badge">Интерактивный BI</span><span class="luxury-badge">PC · iPad · Mobile</span></div>'
        '</div></section>',
        unsafe_allow_html=True,
    )

    uploaded_files = st.file_uploader(
        "Загрузите общую выгрузку Excel",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        help="Название файла может быть любым. Магазины и период определяются по содержимому.",
        key="upload_widget",
    )
    persist_uploads(uploaded_files)
    active_files = saved_uploads()
    sidebar_navigation(bool(active_files))
    mobile_navigation(bool(active_files))

    if not active_files:
        st.markdown(
            '<div class="upload-panel"><b>Перетащите Excel-файл сюда</b><br>'
            '<span class="small-muted">После загрузки откроются сводка, магазины, интерактивная аналитика и поставщики.</span></div>',
            unsafe_allow_html=True,
        )
        render_about()
        st.stop()

    file_names = ", ".join(item.name for item in active_files)
    st.success(f"Загружено: {file_names}")

    with st.spinner("Обрабатываем отчет..."):
        stores_dict, errors, supplier_df = parse_report_bundle(cache_payloads(active_files))


    if errors:
        st.warning("Некоторые файлы не удалось обработать:\n" + "\n".join(f"• {name}: {error}" for name, error in errors))
    stores = list(stores_dict.values())
    summary_df = network_summary(stores)
    if summary_df.empty or "Количество" not in summary_df.columns:
        st.error("В файле не найдены строки продаж. Проверьте структуру выгрузки.")
        render_about()
        st.stop()

    # EXECUTIVE BRIEF — compact management view shown first after upload.
    st.markdown('<div id="executive"></div>', unsafe_allow_html=True)
    section_divider(
        'Оперативная сводка',
        'Ключевые показатели сети и фактические акценты для быстрого просмотра на iPad.',
        'ДЛЯ РУКОВОДИТЕЛЯ',
    )
    render_executive_brief(stores, summary_df, supplier_df)

    # SUMMARY
    st.markdown('<div id="summary"></div>', unsafe_allow_html=True)
    section_divider('Сводка по сети', 'Ключевые показатели и структура продаж по всем магазинам.', 'ОБЩИЙ ОБЗОР')
    total_qty = int(summary_df["Количество"].sum())
    total_sales = float(summary_df["Выручка"].sum())
    periods = sorted(set(summary_df["Период"].tolist())) if "Период" in summary_df.columns else []
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Период", periods[0] if len(periods) == 1 else f"{len(periods)} периода")
    with c2:
        kpi_card("Магазинов", str(len(stores)))
    with c3:
        kpi_card("Всего изделий", money(total_qty) + " шт.")
    with c4:
        kpi_card("Общая выручка", money(total_sales) + " VND")
    st.dataframe(formatted_table(summary_df), width="stretch", hide_index=True)
    chart_cols = st.columns(3)
    for col, segment in zip(chart_cols, SEG_ORDER):
        with col:
            locked_plotly_chart(
                segment_bar(summary_df, segment),
                width="stretch",
                key=f"summary_segment_{segment}",
            )
    insight_panel('Аналитика по сети', network_conclusions(summary_df))

    # STORES — fragment reruns only this block when its controls change.
    st.markdown('<div id="stores"></div>', unsafe_allow_html=True)
    section_divider('Магазины', 'Подробная аналитика выбранного магазина.', 'МАГАЗИНЫ')
    render_store_fragment(stores)

    # INTERACTIVE — independent fragment prevents a full-page rebuild per filter.
    st.markdown('<div id="interactive"></div>', unsafe_allow_html=True)
    section_divider(
        'Интерактивная аналитика',
        'Фильтруйте магазин, сегмент, камень и номенклатурную группу.',
        'ИССЛЕДОВАНИЕ ДАННЫХ',
    )
    render_interactive_fragment(stores)

    # SUPPLIERS — independent fragment keeps summary/store charts untouched.
    st.markdown('<div id="suppliers"></div>', unsafe_allow_html=True)
    section_divider(
        'Поставщики',
        'Сравнение поставщиков по выручке, количеству, магазинам, камням и группам.',
        'ПОСТАВЩИКИ',
    )
    render_supplier_fragment(supplier_df)

    render_about()


if __name__ == "__main__":
    main()
