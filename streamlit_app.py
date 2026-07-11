from __future__ import annotations

import io
import json
import tempfile
from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from src.report import (
    COLORED_ORDER,
    PEARL_ORDER,
    SEG_COLORS,
    SEG_ORDER,
    TOP_ORDER,
    build_report_units,
    run_files,
    totals_for,
)

APP_VERSION = "0.1.0-test"
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

st.set_page_config(
    page_title="Analitika — Princess Jewelry",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded",
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
  background: rgba(255,255,255,.95);
  border-right: 1px solid var(--line);
}
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
}
.kpi-label { color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .06em; }
.kpi-value { font-family: Georgia, serif; font-size: 29px; font-weight: 700; color: #16120d; margin-top: 9px; }
.kpi-note { color: var(--gold); font-size: 12px; margin-top: 6px; }
.section-title { font-family: Georgia, serif; font-size: 28px; margin: 20px 0 10px; }
.insight {
  border-left: 4px solid var(--gold); background: rgba(255,255,255,.93); border-radius: 0 12px 12px 0;
  padding: 13px 15px; margin: 8px 0; border-top: 1px solid var(--line); border-right: 1px solid var(--line); border-bottom: 1px solid var(--line);
}
.small-muted { color: var(--muted); font-size: 12px; }
div[data-testid="stFileUploader"] section {
  border: 1px dashed #c9aa72; border-radius: 14px; background: #fffdf9;
}
div.stButton > button, div.stDownloadButton > button {
  border-radius: 9px; border: 1px solid #2a2114; background: linear-gradient(90deg, #111 0%, #2b241c 100%);
  color: #e8c98e; font-weight: 700; min-height: 44px;
}
div.stButton > button:hover, div.stDownloadButton > button:hover {
  border-color: #b7893f; color: #fff; box-shadow: 0 5px 18px rgba(183,137,63,.25);
}
[data-testid="stMetric"] { border: 1px solid var(--line); padding: 12px; border-radius: 12px; background: #fff; }
hr { border-color: var(--line); }
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


def segment_bar(df: pd.DataFrame, segment: str) -> go.Figure:
    qty_key = f"{SEGMENT_LABELS[segment]} — шт. %"
    sales_key = f"{SEGMENT_LABELS[segment]} — продажи %"
    fig = go.Figure()
    fig.add_bar(
        x=df["Магазин"], y=df[qty_key] * 100, name="Шт. %",
        marker_color=SEGMENT_COLORS[segment], text=[pct(v) for v in df[qty_key]], textposition="outside",
    )
    fig.add_bar(
        x=df["Магазин"], y=df[sales_key] * 100, name="Продажи %",
        marker_color=LIGHT_COLORS[segment], text=[pct(v) for v in df[sales_key]], textposition="outside",
    )
    fig.update_layout(
        title=SEGMENT_LABELS[segment].upper(), barmode="group", height=380,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=30, r=20, t=55, b=35), legend=dict(orientation="h", y=1.06),
        yaxis=dict(title="%", range=[0, 105], gridcolor="#ece8e1"),
        xaxis=dict(title=""), font=dict(family="Arial", color="#1c1813"),
    )
    return fig


def donut(labels: list[str], values: list[float], title: str, colors: list[str]) -> go.Figure:
    fig = go.Figure(
        go.Pie(labels=labels, values=values, hole=.58, marker=dict(colors=colors), textinfo="label+percent")
    )
    fig.update_layout(
        title=title, height=360, showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)", margin=dict(l=10, r=10, t=55, b=10),
        font=dict(family="Arial", color="#1c1813"),
    )
    return fig


def stone_dataframe(store) -> pd.DataFrame:
    orders = {
        "TOP STONES": TOP_ORDER,
        "PEARLS": PEARL_ORDER,
        "COLORED STONES": COLORED_ORDER,
    }
    rows = []
    for seg in SEG_ORDER:
        sq, sa = totals_for(store, seg=seg)
        for stone in orders[seg]:
            q, a = totals_for(store, seg, stone)
            rows.append({
                "Сегмент": SEGMENT_LABELS[seg], "Камень": stone,
                "Количество": q, "% количества магазина": q / store.total_qty if store.total_qty else 0,
                "Выручка": a, "% выручки магазина": a / store.total_amount if store.total_amount else 0,
                "Средняя стоимость": a / q if q else 0,
                "% выручки сегмента": a / sa if sa else 0,
            })
    return pd.DataFrame(rows)


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
    colored_share = seg["COLORED STONES"]["amount"] / store.total_amount if store.total_amount else 0
    lines.append(f"Colored Stones занимают {pct(colored_share)} выручки магазина.")
    return lines[:4]


def store_view(store, all_stores: list) -> None:
    st.markdown(f'<div class="section-title">Магазин {base_store_name(store.name)}</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi_card("Выручка", f"{money(store.total_amount)} VND")
    with c2: kpi_card("Продано изделий", money(store.total_qty) + " шт.")
    with c3: kpi_card("Средняя стоимость", f"{money(store.total_amount / store.total_qty if store.total_qty else 0)} VND")
    with c4:
        network_sales = sum(s.total_amount for s in all_stores)
        kpi_card("Доля в выручке сети", pct(store.total_amount / network_sales if network_sales else 0))

    left, right = st.columns([2.3, 1])
    with left:
        seg = segment_totals(store)
        labels = [SEGMENT_LABELS[s] for s in SEG_ORDER]
        colors = [SEGMENT_COLORS[s] for s in SEG_ORDER]
        a, b = st.columns(2)
        with a:
            st.plotly_chart(donut(labels, [seg[s]["amount"] for s in SEG_ORDER], "Структура продаж", colors), use_container_width=True)
        with b:
            st.plotly_chart(donut(labels, [seg[s]["qty"] for s in SEG_ORDER], "Структура количества", colors), use_container_width=True)
    with right:
        st.markdown("### Выводы")
        for line in conclusions(store, all_stores):
            st.markdown(f'<div class="insight">{line}</div>', unsafe_allow_html=True)

    tab1, tab2, tab3, tab4 = st.tabs(["Обзор камней", "Top Stones", "Pearls", "Colored Stones"])
    data = stone_dataframe(store)
    with tab1:
        display = data.copy()
        for col in ["% количества магазина", "% выручки магазина", "% выручки сегмента"]:
            display[col] = display[col].map(pct)
        for col in ["Количество", "Выручка", "Средняя стоимость"]:
            display[col] = display[col].map(money)
        st.dataframe(display, use_container_width=True, hide_index=True)
    for tab, seg_name in zip([tab2, tab3, tab4], ["Top Stones", "Pearls", "Colored Stones"]):
        with tab:
            subset = data[data["Сегмент"] == seg_name]
            x1, x2 = st.columns(2)
            with x1:
                st.plotly_chart(donut(subset["Камень"].tolist(), subset["Количество"].tolist(), f"{seg_name}: количество", None), use_container_width=True)
            with x2:
                st.plotly_chart(donut(subset["Камень"].tolist(), subset["Выручка"].tolist(), f"{seg_name}: выручка", None), use_container_width=True)

    if base_store_name(store.name) == "OUTLET" and store.extras:
        st.markdown("### Дополнительные подразделения OUTLET")
        cols = st.columns(2)
        for idx, name in enumerate(["GIFT TT", "CAFE"]):
            values = store.extras.get(name, {"qty": 0, "amount": 0})
            avg = values["amount"] / values["qty"] if values["qty"] else 0
            with cols[idx]:
                st.markdown(f"**{name}**")
                a, b, c = st.columns(3)
                a.metric("Выручка", money(values["amount"]))
                b.metric("Количество", money(values["qty"]))
                c.metric("Средняя стоимость", money(avg))


def build_excel(uploaded_files) -> bytes:
    with tempfile.TemporaryDirectory(prefix="analitika_web_") as td:
        paths = []
        for uploaded in uploaded_files:
            p = Path(td) / uploaded.name
            p.write_bytes(uploaded.getvalue())
            paths.append(p)
        output = Path(td) / "Analitika_Report.xlsx"
        run_files(paths, output)
        return output.read_bytes()


def parse_uploads(uploaded_files):
    tmp = tempfile.TemporaryDirectory(prefix="analitika_preview_")
    paths = []
    for uploaded in uploaded_files:
        p = Path(tmp.name) / uploaded.name
        p.write_bytes(uploaded.getvalue())
        paths.append(p)
    stores, errors = build_report_units(paths)
    return tmp, stores, errors


with st.sidebar:
    logo = Path(__file__).parent / "assets" / "logo.png"
    if logo.exists():
        st.image(str(logo), use_container_width=True)
    st.markdown("---")
    st.markdown("**Princess Jewelry Analytics**")
    st.caption(f"Analitika Web {APP_VERSION}")
    st.markdown("Главная")
    st.markdown("Сводка")
    st.markdown("Магазины")
    st.markdown("Аналитика")
    st.markdown("Поставщики · скоро")
    st.markdown("Продавцы · скоро")
    st.markdown("Сравнение · скоро")
    st.markdown("---")
    st.caption("Разработка: Vladimir Panasyan")

st.markdown(
    '<div class="brand-card"><div class="brand-kicker">Princess Jewelry Analytics</div>'
    '<div class="brand-title">Аналитика продаж</div>'
    '<div class="brand-subtitle">Загрузите общую Excel-выгрузку, чтобы получить сводку по сети и готовый отчет.</div></div>',
    unsafe_allow_html=True,
)

uploaded_files = st.file_uploader(
    "Загрузите общую выгрузку Excel",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    help="Название файла может быть любым. Магазины и период определяются по содержимому.",
)

if not uploaded_files:
    st.markdown('<div class="upload-panel"><b>Перетащите Excel-файл сюда</b><br><span class="small-muted">Поддерживается одна общая выгрузка или несколько файлов за разные периоды.</span></div>', unsafe_allow_html=True)
    st.info("После загрузки появятся сводка, страницы магазинов и кнопка скачивания готового отчета.")
    st.stop()

preview_tmp, stores_dict, errors = parse_uploads(uploaded_files)
try:
    if errors:
        st.warning("Некоторые файлы не удалось обработать:\n" + "\n".join(f"• {n}: {e}" for n, e in errors))
    if not stores_dict:
        st.error("Не удалось найти магазины в загруженной выгрузке.")
        st.stop()

    stores = list(stores_dict.values())
    summary_df = network_summary(stores)
    total_qty = int(summary_df["Количество"].sum())
    total_sales = float(summary_df["Выручка"].sum())
    avg_price = total_sales / total_qty if total_qty else 0
    periods = sorted(set(summary_df["Период"].tolist()))

    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi_card("Период", periods[0] if len(periods) == 1 else f"{len(periods)} периода")
    with c2: kpi_card("Магазинов найдено", str(len(stores)))
    with c3: kpi_card("Продано изделий", money(total_qty) + " шт.")
    with c4: kpi_card("Общая выручка", money(total_sales) + " VND")

    st.markdown('<div class="section-title">Сводка по сети</div>', unsafe_allow_html=True)
    visible = summary_df.copy()
    visible["Количество"] = visible["Количество"].map(money)
    visible["Выручка"] = visible["Выручка"].map(money)
    visible["Средняя стоимость"] = visible["Средняя стоимость"].map(money)
    for col in [c for c in visible.columns if "%" in c]:
        visible[col] = visible[col].map(pct)
    st.dataframe(visible, use_container_width=True, hide_index=True)

    chart_cols = st.columns(3)
    for col, seg in zip(chart_cols, SEG_ORDER):
        with col:
            st.plotly_chart(segment_bar(summary_df, seg), use_container_width=True)

    st.markdown('<div class="section-title">Магазины</div>', unsafe_allow_html=True)
    store_names = [base_store_name(s.name) for s in stores]
    chosen = st.selectbox("Выберите магазин", store_names, index=0)
    chosen_store = next(s for s in stores if base_store_name(s.name) == chosen)
    store_view(chosen_store, stores)

    st.markdown("---")
    st.markdown("### Экспорт")
    try:
        excel_bytes = build_excel(uploaded_files)
        st.download_button(
            "Скачать отчет Excel / открыть в Google Sheets",
            data=excel_bytes,
            file_name="Analitika_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption("Google Sheets: скачайте файл и откройте его через Google Drive → Открыть с помощью Google Таблиц. Прямое создание таблицы будет подключено после настройки Google OAuth.")
    except Exception as exc:
        st.error(f"Не удалось сформировать Excel: {exc}")
finally:
    preview_tmp.cleanup()
