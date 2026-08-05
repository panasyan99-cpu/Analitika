from __future__ import annotations

from html import escape
from typing import Iterable
import gc
import math

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from src.management_report_analytics import (
    build_management_snapshot,
    new_and_lost_sku,
    significant_rows,
)
from src.management_report_parser import ParsedReport, parse_report
from src.currency import get_vnd_per_usd, round_usd_to_tens, vnd_to_usd_display
from src.management_report_suppliers import (
    UNKNOWN_SUPPLIER,
    load_supplier_catalog,
    normalize_sku,
    save_cloud_overrides,
)
from src.management_block_reports import (
    SALES,
    CONSULTANTS,
    SUPPLIERS,
    KIND_LABELS,
    BlockReport,
    compare_metric_maps,
    compare_totals,
    parse_block_report,
    validate_period_bundle,
)


REPORT_STATE_KEY = "management_report_parsed"
OVERRIDES_STATE_KEY = "management_supplier_overrides"


TECHNICAL_TABLE_COLUMNS = [
    "Позиция",
    "Количество · Период 1", "Количество · Период 2", "Δ количества", "Δ количества, %",
    "Выручка · Период 1", "Выручка · Период 2", "Δ выручки", "Δ выручки, %",
    "Доля · Период 1, %", "Доля · Период 2, %", "Δ доли, п.п.",
    "Средняя цена · Период 1", "Средняя цена · Период 2", "Δ средней цены, %",
]


def _css() -> None:
    st.markdown(
        """
<style>
.management-summary {
  border:1px solid rgba(183,137,63,.35); border-radius:18px; padding:22px 24px;
  background:linear-gradient(135deg,rgba(255,253,248,.98),rgba(246,237,221,.92));
  box-shadow:0 14px 34px rgba(75,48,14,.08); margin:.7rem 0 1.1rem;
}
.management-summary h3 { margin:0 0 10px; font-family:Georgia,serif; color:#251b10; }
.management-summary p { margin:0 0 9px; color:#44392d; line-height:1.62; }
.management-kpi-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin:10px 0 18px; }
.management-kpi { border:1px solid #e9dfd0; border-radius:15px; background:#fffdfa; padding:15px 16px; min-height:112px; }
.management-kpi span { display:block; color:#7a6f62; font-size:11px; font-weight:800; text-transform:uppercase; letter-spacing:.05em; }
.management-kpi strong { display:block; margin-top:9px; color:#1c160f; font-family:Georgia,serif; font-size:20px; line-height:1.2; }
.management-kpi em { display:block; margin-top:8px; font-style:normal; font-size:12px; font-weight:800; }
.management-positive { color:#287342 !important; }
.management-negative { color:#a14337 !important; }
.management-neutral { color:#87621f !important; }
.management-highlight-grid { display:grid; grid-template-columns:1fr 1fr; gap:14px; margin:.6rem 0 1.25rem; }
.management-highlight { border-radius:15px; padding:16px 18px; border:1px solid #e9dfd0; background:#fffdfa; }
.management-highlight.positive { border-left:4px solid #4f9365; }
.management-highlight.negative { border-left:4px solid #b35b4f; }
.management-highlight b { display:block; margin-bottom:8px; color:#251b10; }
.management-highlight div { color:#5e5347; line-height:1.55; margin:5px 0; }
.management-section-note { border-left:3px solid #b7893f; padding:8px 0 8px 14px; color:#5f5346; line-height:1.6; margin:.15rem 0 .9rem; }
.management-quality { border:1px solid #eadfcf; border-radius:13px; padding:12px 15px; background:#fffaf1; color:#5b4b38; margin:.4rem 0 .8rem; }
@media(max-width:900px) {
  .management-kpi-grid { grid-template-columns:repeat(2,minmax(0,1fr)); }
  .management-highlight-grid { grid-template-columns:1fr; }
}
@media(max-width:520px) {
  .management-kpi-grid { grid-template-columns:1fr; }
  .management-summary { padding:18px; }
}
</style>
        """,
        unsafe_allow_html=True,
    )


def _money(value: object) -> str:
    """Format a VND source value as whole USD using the site-wide rate."""
    try:
        number = vnd_to_usd_display(float(value or 0), get_vnd_per_usd())
    except (TypeError, ValueError):
        number = 0.0
    sign = "−" if number < 0 else ""
    return f"{sign}${abs(number):,.0f}".replace(",", " ")


def _money_usd_number(value: object) -> float:
    try:
        return vnd_to_usd_display(float(value or 0), get_vnd_per_usd())
    except (TypeError, ValueError):
        return 0.0


def _is_percent_column(name: object) -> bool:
    text = str(name)
    return text.endswith(", %") or text.endswith(" %") or text.startswith("Δ ") and text.endswith("%") or "Доля" in text and "%" in text or "Скидка" in text and "%" in text


def _is_money_column(name: object) -> bool:
    text = str(name)
    folded = text.casefold()
    if _is_percent_column(text):
        return False
    return (
        "выруч" in folded
        or "средняя цена" in folded
        or ("возврат" in folded and "колич" not in folded and "шт" not in folded)
        or folded in {"выручка", "возвраты", "δ возвратов", "Δ возвратов".casefold()}
    )


def _quantity(value: object) -> str:
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        number = 0.0
    if abs(number - round(number)) < 0.0005:
        return f"{int(round(number)):,}".replace(",", " ")
    return f"{number:,.3f}".replace(",", " ").rstrip("0").rstrip(".")


def _pct(value: object, *, signed: bool = True) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "—"
    if math.isinf(number):
        return "новая позиция"
    prefix = "+" if signed and number > 0 else ""
    return f"{prefix}{number:.1f}%".replace("-", "−")


def _pp(value: object) -> str:
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        number = 0.0
    prefix = "+" if number > 0 else ""
    return f"{prefix}{number:.1f} п.п.".replace("-", "−")


def _tone(value: float) -> str:
    if value > 0.05:
        return "management-positive"
    if value < -0.05:
        return "management-negative"
    return "management-neutral"


def _verb(value: float, feminine: bool = True) -> str:
    if value > 0.05:
        return "увеличилась" if feminine else "увеличилось"
    if value < -0.05:
        return "снизилась" if feminine else "снизилось"
    return "практически не изменилась" if feminine else "практически не изменилось"


def _metric_card(label: str, value: str, delta: str, numeric_delta: float) -> str:
    return (
        '<div class="management-kpi">'
        f'<span>{escape(label)}</span><strong>{escape(value)}</strong>'
        f'<em class="{_tone(numeric_delta)}">{escape(delta)}</em></div>'
    )


def _overall_narrative(snapshot: dict[str, object]) -> str:
    overall = snapshot["overall"]
    old = overall["old"]
    new = overall["new"]
    revenue_pct = float(overall["revenue_pct"] or 0)
    quantity_pct = float(overall["quantity_pct"] or 0)
    average_pct = float(overall["average_price_pct"] or 0)
    discount_delta = float(overall["discount_delta_pp"] or 0)
    net_pct = float(overall["net_revenue_pct"] or 0)

    first = (
        f"Выручка {_verb(revenue_pct)} с {_money(old['revenue'])} до {_money(new['revenue'])} "
        f"({_pct(revenue_pct)}). Количество реализованных изделий {_verb(quantity_pct, feminine=False)} "
        f"с {_quantity(old['quantity'])} до {_quantity(new['quantity'])} единиц ({_pct(quantity_pct)}), "
        f"а средняя стоимость изделия {_verb(average_pct)} на {_pct(abs(average_pct), signed=False)}. "
        f"Средняя скидка изменилась с {old['discount_pct']:.1f}% до {new['discount_pct']:.1f}% "
        f"({_pp(discount_delta)})."
    )

    if abs(quantity_pct) >= abs(average_pct) * 2.5 and abs(quantity_pct) >= 3:
        structure = (
            "Динамика оборота сопровождалась прежде всего изменением количества продаж; "
            "изменение средней стоимости изделия было существенно меньше."
        )
    elif abs(average_pct) >= abs(quantity_pct) * 1.7 and abs(average_pct) >= 3:
        structure = (
            "Динамика оборота сопровождалась прежде всего изменением средней стоимости проданного изделия; "
            "изменение количества было заметно слабее."
        )
    else:
        structure = "Изменение оборота сопровождалось одновременной динамикой количества и средней стоимости продаж."

    returns = (
        f"Чистая выручка после возвратов {_verb(net_pct)} до {_money(overall['net_revenue_new'])} "
        f"({_pct(net_pct)}). Сумма возвратов составила {_money(new['return_amount'])} против "
        f"{_money(old['return_amount'])}; их доля в выручке изменилась с "
        f"{overall['return_share_old']:.1f}% до {overall['return_share_new']:.1f}%."
    )
    return "<p>" + escape(first) + "</p><p>" + escape(structure) + "</p><p>" + escape(returns) + "</p>"


def _row_sentence(row: pd.Series) -> str:
    name = str(row.get("Позиция", ""))
    return (
        f"{name}: {_money(row.get('Выручка · Период 1', 0))} → "
        f"{_money(row.get('Выручка · Период 2', 0))}; "
        f"Δ {_money(row.get('Δ выручки', 0))} ({_pct(row.get('Δ выручки, %'))})."
    )


def _driver_cards(snapshot: dict[str, object]) -> None:
    frames = snapshot["dimensions"]
    combined: list[dict[str, object]] = []
    labels = {
        "stores": "Магазин",
        "suppliers": "Поставщик",
        "categories": "Категория",
        "stone_groups": "Группа вставок",
    }
    for key, label in labels.items():
        frame = frames[key]
        if frame.empty:
            continue
        for _, row in frame.iterrows():
            if str(row["Позиция"]) == UNKNOWN_SUPPLIER:
                continue
            combined.append({
                "dimension": label,
                "name": row["Позиция"],
                "delta": float(row["Δ выручки"]),
                "pct": row["Δ выручки, %"],
            })
    positive = sorted((item for item in combined if item["delta"] > 0), key=lambda item: item["delta"], reverse=True)[:5]
    negative = sorted((item for item in combined if item["delta"] < 0), key=lambda item: item["delta"])[:5]

    def lines(items: list[dict[str, object]]) -> str:
        if not items:
            return "<div>Значимых изменений не зафиксировано.</div>"
        return "".join(
            f'<div><b>{escape(str(item["dimension"]))} · {escape(str(item["name"]))}</b> — '
            f'{escape(_money(item["delta"]))} ({escape(_pct(item["pct"]))})</div>'
            for item in items
        )

    st.markdown(
        '<div class="management-highlight-grid">'
        '<div class="management-highlight positive"><b>Крупнейшие положительные вклады</b>'
        + lines(positive)
        + '</div><div class="management-highlight negative"><b>Крупнейшие отрицательные вклады</b>'
        + lines(negative)
        + '</div></div>',
        unsafe_allow_html=True,
    )


def _delta_chart(frame: pd.DataFrame, title: str, *, limit: int = 12) -> go.Figure:
    data = frame.copy()
    if data.empty:
        return go.Figure()
    data["_abs"] = data["Δ выручки"].abs()
    data = data.nlargest(limit, "_abs").sort_values("Δ выручки")
    usd_values = data["Δ выручки"].map(_money_usd_number)
    colors = ["#9f4f43" if value < 0 else "#4f8c61" for value in usd_values]
    labels = [_money(value) for value in data["Δ выручки"]]
    max_abs = float(usd_values.abs().max()) if not usd_values.empty else 0.0
    headroom = max_abs * 1.30 if max_abs else None
    x_range = [-headroom, headroom] if headroom else None
    fig = go.Figure(go.Bar(
        x=usd_values,
        y=data["Позиция"],
        orientation="h",
        marker_color=colors,
        text=labels,
        textposition="outside",
        cliponaxis=False,
        hovertemplate="%{y}<br>Δ: %{text}<extra></extra>",
    ))
    fig.update_layout(
        title=title,
        height=max(330, 44 * len(data) + 100),
        margin=dict(l=20, r=125, t=55, b=45),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,255,255,.68)",
        font=dict(color="#30271d"),
        xaxis=dict(
            title="USD", tickprefix="$", tickformat=",.0f",
            showgrid=True, gridcolor="rgba(90,65,35,.08)", fixedrange=True, range=x_range,
        ),
        yaxis=dict(fixedrange=True, automargin=True),
        showlegend=False,
    )
    return fig


def _comparison_chart(frame: pd.DataFrame, title: str, first_label: str, second_label: str, *, limit: int = 10) -> go.Figure:
    data = frame.copy()
    if data.empty:
        return go.Figure()
    data["_max"] = data[["Выручка · Период 1", "Выручка · Период 2"]].max(axis=1)
    data = data.nlargest(limit, "_max").sort_values("_max")
    first_usd = data["Выручка · Период 1"].map(_money_usd_number)
    second_usd = data["Выручка · Период 2"].map(_money_usd_number)
    max_value = float(pd.concat([first_usd, second_usd]).max()) if not data.empty else 0.0
    x_range = [0, max_value * 1.34] if max_value > 0 else None
    fig = go.Figure()
    fig.add_bar(
        x=first_usd, y=data["Позиция"], orientation="h",
        name=first_label, marker_color="#c9b38e",
        text=[_money(value) for value in data["Выручка · Период 1"]],
        textposition="outside", cliponaxis=False,
        hovertemplate="%{y}<br>" + escape(first_label) + ": %{text}<extra></extra>",
    )
    fig.add_bar(
        x=second_usd, y=data["Позиция"], orientation="h",
        name=second_label, marker_color="#9b6a28",
        text=[_money(value) for value in data["Выручка · Период 2"]],
        textposition="outside", cliponaxis=False,
        hovertemplate="%{y}<br>" + escape(second_label) + ": %{text}<extra></extra>",
    )
    fig.update_layout(
        title=title, barmode="group", height=max(340, 48 * len(data) + 110),
        margin=dict(l=20, r=140, t=55, b=45), paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,255,255,.68)", font=dict(color="#30271d"),
        xaxis=dict(
            title="USD", tickprefix="$", tickformat=",.0f", range=x_range,
            showgrid=True, gridcolor="rgba(90,65,35,.08)", fixedrange=True,
        ),
        yaxis=dict(fixedrange=True, automargin=True), legend=dict(orientation="h", y=1.08),
    )
    return fig


def _prepare_table_display(frame: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, object]]:
    """Convert every monetary column to numeric USD and build readable formats."""
    data = frame.copy()
    rename: dict[str, str] = {}
    rate = get_vnd_per_usd()
    for column in list(data.columns):
        if _is_money_column(column):
            data[column] = pd.to_numeric(data[column], errors="coerce").fillna(0) / rate
            rename[column] = f"{column}, USD"
    if rename:
        data = data.rename(columns=rename)

    config: dict[str, object] = {}
    for column in data.columns:
        name = str(column)
        if name == "Позиция":
            config[column] = st.column_config.TextColumn(width="large")
        elif name.endswith(", USD"):
            config[column] = st.column_config.NumberColumn(step=10, width="medium")
        elif _is_percent_column(name):
            config[column] = st.column_config.NumberColumn(format="%.1f%%", width="small")
        elif name == "Δ доли, п.п.":
            config[column] = st.column_config.NumberColumn(format="%.1f", width="small")
        elif "Количество" in name or name.startswith("Δ количества"):
            config[column] = st.column_config.NumberColumn(format="localized", step=0.001, width="small")
    return data, config


def _is_delta_column(name: object) -> bool:
    text = str(name).strip().casefold()
    return text.startswith(("δ", "Δ".casefold())) or text in {"изменение", "разница"} or "разниц" in text


def _delta_cell_style(value: object) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return ""
    if math.isnan(number) or abs(number) < 1e-12:
        return ""
    if number > 0:
        return "color:#1f6f3d;background-color:rgba(69,145,91,.13);font-weight:700"
    return "color:#a13d35;background-color:rgba(181,76,65,.12);font-weight:700"


def _style_delta_columns(data: pd.DataFrame):
    delta_columns = [column for column in data.columns if _is_delta_column(column)]
    usd_columns = [column for column in data.columns if str(column).endswith(", USD")]
    if not delta_columns and not usd_columns:
        return data
    styled = data.style
    if usd_columns:
        styled = styled.format({
            column: lambda value: f"{round_usd_to_tens(float(value or 0)):,.0f}".replace(",", " ")
            for column in usd_columns
        })
    if delta_columns:
        styled = styled.map(_delta_cell_style, subset=delta_columns)
    return styled


def _table(frame: pd.DataFrame, *, key: str, limit: int | None = 15, columns: Iterable[str] | None = None) -> None:
    if frame.empty:
        st.info("В выбранном разрезе нет данных.")
        return
    data = frame.copy()
    if columns:
        data = data[[column for column in columns if column in data.columns]]
    if limit is not None:
        data = data.head(limit)
    display, config = _prepare_table_display(data)
    st.dataframe(
        _style_delta_columns(display),
        width="stretch",
        hide_index=True,
        key=key,
        column_config=config,
    )


def _section_summary(frame: pd.DataFrame, noun: str) -> str:
    if frame.empty:
        return "Данные отсутствуют."
    positive = significant_rows(frame, positive=True, limit=1)
    negative = significant_rows(frame, positive=False, limit=1)
    fragments = []
    if not positive.empty:
        fragments.append(f"Наибольший положительный вклад среди {noun}: {_row_sentence(positive.iloc[0])}")
    if not negative.empty:
        fragments.append(f"Наибольшее снижение среди {noun}: {_row_sentence(negative.iloc[0])}")
    return " ".join(fragments) or "Существенных разнонаправленных изменений не зафиксировано."


def _render_dimension_section(
    title: str,
    frame: pd.DataFrame,
    first_label: str,
    second_label: str,
    *,
    noun: str,
    key: str,
    table_columns: Iterable[str] | None = None,
) -> None:
    st.markdown(f"## {title}")
    st.markdown(
        f'<div class="management-section-note">{escape(_section_summary(frame, noun))}</div>',
        unsafe_allow_html=True,
    )
    if frame.empty:
        return
    left, right = st.columns(2)
    with left:
        st.plotly_chart(
            _comparison_chart(frame, f"{title}: выручка двух периодов", first_label, second_label),
            width="stretch", config={"displayModeBar": False, "scrollZoom": False}, key=f"{key}_comparison",
        )
    with right:
        st.plotly_chart(
            _delta_chart(frame, f"{title}: изменение выручки"),
            width="stretch", config={"displayModeBar": False, "scrollZoom": False}, key=f"{key}_delta",
        )
    _table(frame.sort_values("Выручка · Период 2", ascending=False), key=f"{key}_top", columns=table_columns)
    with st.expander(f"Полная таблица — {title.lower()}", expanded=False):
        _table(frame.sort_values("Выручка · Период 2", ascending=False), key=f"{key}_full", limit=None)


def _render_upload() -> None:
    st.markdown("## Загрузка двух периодов — три независимых блока")
    st.caption(
        "Для каждого периода загрузите три одинаково датированных отчета 1С. "
        "Название Excel-файла может быть любым: сайт определяет корректность по структуре отчета. "
        "Каждый блок анализируется отдельно. Итоговый объединённый блок сравнения не строится."
    )
    with st.form("management_report_upload_form", clear_on_submit=False):
        st.markdown("### 1. Продажи по магазинам")
        left, right = st.columns(2)
        with left:
            sales_first = st.file_uploader(
                "Период 1 · Продажи по магазинам", type=["xlsx", "xlsm"], key="management_sales_1",
                help="Группировки: Магазин → Камень/вставка → Проба → Номенклатурная группа",
            )
        with right:
            sales_second = st.file_uploader(
                "Период 2 · Продажи по магазинам", type=["xlsx", "xlsm"], key="management_sales_2",
                help="Группировки: Магазин → Камень/вставка → Проба → Номенклатурная группа",
            )

        st.markdown("### 2. Продажи по консультантам")
        left, right = st.columns(2)
        with left:
            consultants_first = st.file_uploader(
                "Период 1 · Продажи по консультантам", type=["xlsx", "xlsm"], key="management_consultants_1",
                help="Группировки: Менеджер → Проба → Номенклатурная группа",
            )
        with right:
            consultants_second = st.file_uploader(
                "Период 2 · Продажи по консультантам", type=["xlsx", "xlsm"], key="management_consultants_2",
                help="Группировки: Менеджер → Проба → Номенклатурная группа",
            )

        st.markdown("### 3. Продажи по поставщикам")
        left, right = st.columns(2)
        with left:
            suppliers_first = st.file_uploader(
                "Период 1 · Продажи по поставщикам", type=["xlsx", "xlsm"], key="management_suppliers_1",
                help="Группировки: Номенклатурная группа → Поставщик",
            )
        with right:
            suppliers_second = st.file_uploader(
                "Период 2 · Продажи по поставщикам", type=["xlsx", "xlsm"], key="management_suppliers_2",
                help="Группировки: Номенклатурная группа → Поставщик",
            )
        submitted = st.form_submit_button("Построить управленческий отчет", type="primary", width="stretch")

    if not submitted:
        return

    uploads = {
        "first": {SALES: sales_first, CONSULTANTS: consultants_first, SUPPLIERS: suppliers_first},
        "second": {SALES: sales_second, CONSULTANTS: consultants_second, SUPPLIERS: suppliers_second},
    }
    missing = [
        f"{period} · {KIND_LABELS[kind]}"
        for period, bundle in uploads.items()
        for kind, uploaded in bundle.items()
        if uploaded is None
    ]
    if missing:
        st.error("Загрузите все шесть файлов: " + "; ".join(missing))
        return

    parsed: dict[str, dict[str, BlockReport]] = {"first": {}, "second": {}}
    try:
        with st.spinner("Читаем шесть выгрузок и сверяем контрольные итоги 1С..."):
            for period, bundle in uploads.items():
                for kind, uploaded in bundle.items():
                    uploaded.seek(0)
                    parsed[period][kind] = parse_block_report(
                        uploaded,
                        kind=kind,
                        source_name=uploaded.name,
                    )
    except Exception as exc:
        st.error(f"Не удалось обработать выгрузки: {exc}")
        return

    errors = [
        *(f"Период 1: {message}" for message in validate_period_bundle(parsed["first"])),
        *(f"Период 2: {message}" for message in validate_period_bundle(parsed["second"])),
    ]
    first_start = parsed["first"][SALES].meta.period_start
    second_start = parsed["second"][SALES].meta.period_start
    if first_start == second_start:
        errors.append("В обоих наборах указан одинаковый период.")
    if errors:
        for message in errors:
            st.error(message)
        return

    st.session_state[REPORT_STATE_KEY] = {
        "format": "three_blocks",
        "first": parsed["first"],
        "second": parsed["second"],
    }
    st.rerun()


def _render_daily(overall: dict[str, object], first_label: str, second_label: str) -> None:
    daily = overall["daily"]
    st.markdown("## В среднем за день")
    cards = [
        _metric_card("Выручка в день", _money(daily["new_revenue"]), _pct(daily["revenue_pct"]), float(daily["revenue_pct"] or 0)),
        _metric_card("Чистая выручка в день", _money(daily["new_net_revenue"]), _pct(daily["net_revenue_pct"]), float(daily["net_revenue_pct"] or 0)),
        _metric_card("Продано в день", _quantity(daily["new_quantity"]), _pct(daily["quantity_pct"]), float(daily["quantity_pct"] or 0)),
        _metric_card("Возвратов, шт. в день", _quantity(daily["new_return_quantity"]), _pct(daily["return_quantity_pct"]), float(daily["return_quantity_pct"] or 0)),
        _metric_card("Сумма возвратов в день", _money(daily["new_return_amount"]), _pct(daily["return_pct"]), float(daily["return_pct"] or 0)),
    ]
    st.markdown('<div class="management-kpi-grid">' + "".join(cards) + "</div>", unsafe_allow_html=True)
    daily_frame = pd.DataFrame([
        {"Показатель": "Выручка", first_label: _money(daily["old_revenue"]), second_label: _money(daily["new_revenue"]), "Изменение": _pct(daily["revenue_pct"])},
        {"Показатель": "Чистая выручка", first_label: _money(daily["old_net_revenue"]), second_label: _money(daily["new_net_revenue"]), "Изменение": _pct(daily["net_revenue_pct"])},
        {"Показатель": "Количество", first_label: _quantity(daily["old_quantity"]), second_label: _quantity(daily["new_quantity"]), "Изменение": _pct(daily["quantity_pct"])},
        {"Показатель": "Количество возвратов", first_label: _quantity(daily["old_return_quantity"]), second_label: _quantity(daily["new_return_quantity"]), "Изменение": _pct(daily["return_quantity_pct"])},
        {"Показатель": "Сумма возвратов", first_label: _money(daily["old_return_amount"]), second_label: _money(daily["new_return_amount"]), "Изменение": _pct(daily["return_pct"])},
    ])
    st.dataframe(
        daily_frame, width="stretch", hide_index=True, key="management_daily_table",
        column_config={
            "Показатель": st.column_config.TextColumn(width="large"),
            first_label: st.column_config.TextColumn(width="medium"),
            second_label: st.column_config.TextColumn(width="medium"),
            "Изменение": st.column_config.TextColumn(width="small"),
        },
    )


def _render_supplier_learning(snapshot: dict[str, object], catalog) -> None:
    quality = snapshot["supplier_quality"]
    st.markdown(
        '<div class="management-quality">'
        f'Распознано по поставщикам: <b>{quality["old_revenue_coverage_pct"]:.1f}%</b> выручки периода 1 и '
        f'<b>{quality["new_revenue_coverage_pct"]:.1f}%</b> выручки периода 2. '
        'Точные SKU имеют приоритет; затем применяются проверенные семейства SKU. Остальное отражается отдельной строкой.'
        '</div>',
        unsafe_allow_html=True,
    )
    first_unknown = quality["old_unknown"]
    second_unknown = quality["new_unknown"]
    combined = pd.concat([
        first_unknown.assign(**{"Период": "Период 1"}) if not first_unknown.empty else pd.DataFrame(),
        second_unknown.assign(**{"Период": "Период 2"}) if not second_unknown.empty else pd.DataFrame(),
    ], ignore_index=True)
    if combined.empty:
        return
    grouped = combined.groupby("SKU", as_index=False).agg(
        Количество=("Количество", "sum"),
        Выручка=("Выручка", "sum"),
        Камень=("Камень", lambda values: "; ".join(sorted(set(filter(None, values))))),
        Категория=("Категория", lambda values: "; ".join(sorted(set(filter(None, values))))),
    ).sort_values("Выручка", ascending=False)

    with st.expander(f"Неопределенные SKU — {len(grouped)}", expanded=False):
        st.caption("Ниже показаны только позиции, для которых нет подтвержденного точного или семейного соответствия.")
        _table(grouped, key="management_unknown_supplier_sku", limit=100, columns=["SKU", "Количество", "Выручка", "Камень", "Категория"])

        edit = grouped.head(30).copy()
        edit["Выручка, USD"] = edit["Выручка"].map(_money_usd_number)
        edit["Поставщик"] = ""
        supplier_options = ["", *catalog.suppliers]
        edited = st.data_editor(
            edit[["SKU", "Выручка, USD", "Поставщик"]],
            width="stretch",
            hide_index=True,
            disabled=["SKU", "Выручка, USD"],
            key="management_supplier_editor",
            column_config={
                "Выручка, USD": st.column_config.NumberColumn(format="localized", step=1, width="medium"),
                "Поставщик": st.column_config.SelectboxColumn(options=supplier_options),
            },
        )
        if st.button("Сохранить выбранные соответствия", key="management_save_supplier_mapping", width="stretch"):
            session_mapping = dict(catalog.overrides)
            added = 0
            for _, row in edited.iterrows():
                supplier = str(row.get("Поставщик", "")).strip()
                sku = normalize_sku(row.get("SKU", ""))
                if sku and supplier:
                    session_mapping[sku] = supplier
                    added += 1
            st.session_state[OVERRIDES_STATE_KEY] = session_mapping
            durable = save_cloud_overrides(session_mapping) if added else False
            if added:
                message = f"Сохранено соответствий: {added}."
                message += " Данные записаны в облачный справочник." if durable else " Данные применены в текущей сессии."
                st.success(message)
                st.rerun()
            else:
                st.info("Поставщики не выбраны.")


def _render_report(first: ParsedReport, second: ParsedReport) -> None:
    catalog = load_supplier_catalog(session_overrides=st.session_state.get(OVERRIDES_STATE_KEY, {}))
    snapshot = build_management_snapshot(first, second, catalog)
    first = snapshot["first"]
    second = snapshot["second"]
    overall = snapshot["overall"]
    first_label = first.meta.period_label
    second_label = second.meta.period_label

    st.markdown(
        f"**Сравниваются:** {escape(first_label)} ({first.meta.period_days} дней) → "
        f"{escape(second_label)} ({second.meta.period_days} дней)"
    )
    if first.meta.period_days != second.meta.period_days:
        st.warning("Периоды различаются по продолжительности. Блок «В среднем за день» рассчитан отдельно для каждого периода.")
    if st.button("Загрузить другие периоды", key="management_report_replace", width="stretch"):
        st.session_state.pop(REPORT_STATE_KEY, None)
        st.session_state.pop("management_report_file_1", None)
        st.session_state.pop("management_report_file_2", None)
        st.rerun()

    st.markdown('<div class="management-summary"><h3>Управленческое резюме</h3>' + _overall_narrative(snapshot) + '</div>', unsafe_allow_html=True)

    old = overall["old"]
    new = overall["new"]
    cards = [
        _metric_card("Выручка", _money(new["revenue"]), _pct(overall["revenue_pct"]), float(overall["revenue_pct"] or 0)),
        _metric_card("Количество", _quantity(new["quantity"]), _pct(overall["quantity_pct"]), float(overall["quantity_pct"] or 0)),
        _metric_card("Средняя цена", _money(new["average_price"]), _pct(overall["average_price_pct"]), float(overall["average_price_pct"] or 0)),
        _metric_card("Средняя скидка", f'{new["discount_pct"]:.1f}%', _pp(overall["discount_delta_pp"]), float(overall["discount_delta_pp"] or 0)),
        _metric_card("Чистая выручка", _money(overall["net_revenue_new"]), _pct(overall["net_revenue_pct"]), float(overall["net_revenue_pct"] or 0)),
        _metric_card("Сумма возвратов", _money(new["return_amount"]), _pct(overall["return_amount_pct"]), float(overall["return_amount_pct"] or 0)),
        _metric_card("Возвратов, шт.", _quantity(new["return_quantity"]), _pct((new["return_quantity"] - old["return_quantity"]) / old["return_quantity"] * 100 if old["return_quantity"] else 0), (new["return_quantity"] - old["return_quantity"])),
        _metric_card("Доля возвратов", f'{overall["return_share_new"]:.1f}%', _pp(overall["return_share_new"] - overall["return_share_old"]), (overall["return_share_new"] - overall["return_share_old"])),
    ]
    st.markdown('<div class="management-kpi-grid">' + "".join(cards) + '</div>', unsafe_allow_html=True)

    _render_daily(overall, first_label, second_label)

    st.markdown("## Главные изменения")
    _driver_cards(snapshot)

    stores = snapshot["dimensions"]["stores"]
    outlet = snapshot["outlet"]
    if not outlet.empty:
        row = outlet.iloc[0]
        st.markdown("## Магазины")
        st.markdown(
            '<div class="management-section-note">'
            + escape(
                f"Outlet (TT + Gifts-TT + Cafe): {_money(row['Выручка · Период 1'])} → "
                f"{_money(row['Выручка · Период 2'])}; Δ {_money(row['Δ выручки'])} "
                f"({_pct(row['Δ выручки, %'])}). " + _section_summary(stores, "магазинов")
            )
            + '</div>',
            unsafe_allow_html=True,
        )
        left, right = st.columns(2)
        with left:
            st.plotly_chart(_comparison_chart(stores, "Магазины: выручка", first_label, second_label), width="stretch", config={"displayModeBar": False}, key="management_stores_compare")
        with right:
            st.plotly_chart(_delta_chart(stores, "Магазины: вклад в изменение"), width="stretch", config={"displayModeBar": False}, key="management_stores_delta")
        _table(stores.sort_values("Выручка · Период 2", ascending=False), key="management_stores_top", columns=TECHNICAL_TABLE_COLUMNS)
        with st.expander("Полная таблица — магазины", expanded=False):
            _table(stores.sort_values("Выручка · Период 2", ascending=False), key="management_stores_full", limit=None)

    _render_dimension_section(
        "Продавцы", snapshot["dimensions"]["managers"], first_label, second_label,
        noun="продавцов", key="management_managers",
        table_columns=[
            *TECHNICAL_TABLE_COLUMNS,
            "Чистая выручка · Период 1", "Чистая выручка · Период 2", "Δ чистой выручки",
            "Количество возвратов · Период 1", "Количество возвратов · Период 2", "Δ количества возвратов",
            "Возвраты · Период 1", "Возвраты · Период 2",
            "Доля возвратов · Период 1, %", "Доля возвратов · Период 2, %",
        ],
    )

    _render_dimension_section(
        "Поставщики", snapshot["dimensions"]["suppliers"], first_label, second_label,
        noun="поставщиков", key="management_suppliers", table_columns=TECHNICAL_TABLE_COLUMNS,
    )
    _render_supplier_learning(snapshot, catalog)

    _render_dimension_section(
        "Номенклатурные группы", snapshot["dimensions"]["categories"], first_label, second_label,
        noun="категорий", key="management_categories", table_columns=TECHNICAL_TABLE_COLUMNS,
    )

    _render_dimension_section(
        "Группы камней и вставок", snapshot["dimensions"]["stone_groups"], first_label, second_label,
        noun="групп вставок", key="management_stone_groups", table_columns=TECHNICAL_TABLE_COLUMNS,
    )
    with st.expander("Детализация по конкретным камням и вставкам", expanded=False):
        stones = snapshot["dimensions"]["stones"].sort_values("Выручка · Период 2", ascending=False)
        st.plotly_chart(_delta_chart(stones, "Конкретные вставки: крупнейшие изменения", limit=16), width="stretch", config={"displayModeBar": False}, key="management_stones_delta")
        _table(stones, key="management_stones_full", limit=None)

    _render_dimension_section(
        "Пробы и металлы", snapshot["dimensions"]["assay_groups"], first_label, second_label,
        noun="групп металла", key="management_assay_groups", table_columns=TECHNICAL_TABLE_COLUMNS,
    )
    with st.expander("Детализация по пробам", expanded=False):
        _table(snapshot["dimensions"]["assays"].sort_values("Выручка · Период 2", ascending=False), key="management_assays_full", limit=None)

    st.markdown("## Товары и SKU")
    sku = snapshot["dimensions"]["sku"]
    new_sku, lost_sku = new_and_lost_sku(sku)
    growth = significant_rows(sku, positive=True, limit=12)
    decline = significant_rows(sku, positive=False, limit=12)
    tabs = st.tabs(("Крупнейший рост", "Крупнейшее снижение", "Новые продажи", "Нет продаж во втором периоде"))
    with tabs[0]:
        _table(growth, key="management_sku_growth", limit=12, columns=TECHNICAL_TABLE_COLUMNS)
    with tabs[1]:
        _table(decline.sort_values("Δ выручки"), key="management_sku_decline", limit=12, columns=TECHNICAL_TABLE_COLUMNS)
    with tabs[2]:
        _table(new_sku, key="management_sku_new", limit=20, columns=TECHNICAL_TABLE_COLUMNS)
    with tabs[3]:
        _table(lost_sku, key="management_sku_lost", limit=20, columns=TECHNICAL_TABLE_COLUMNS)
    with st.expander("Полная таблица SKU", expanded=False):
        _table(sku.sort_values("Выручка · Период 2", ascending=False), key="management_sku_full", limit=None)

    st.markdown("## Возвраты")
    managers = snapshot["dimensions"]["managers"].copy()
    returns = managers.loc[(managers["Возвраты · Период 1"] > 0) | (managers["Возвраты · Период 2"] > 0)].copy()
    returns["Δ возвратов"] = returns["Возвраты · Период 2"] - returns["Возвраты · Период 1"]
    returns = returns.sort_values("Возвраты · Период 2", ascending=False)
    st.markdown(
        f'<div class="management-section-note">Во втором периоде оформлено '
        f'{escape(_quantity(new["return_quantity"]))} возвратов на {escape(_money(new["return_amount"]))}. '
        f'Доля возвратов в выручке составила {overall["return_share_new"]:.1f}% против '
        f'{overall["return_share_old"]:.1f}% в первом периоде.</div>',
        unsafe_allow_html=True,
    )
    _table(
        returns,
        key="management_returns",
        limit=30,
        columns=[
            "Позиция",
            "Количество возвратов · Период 1", "Количество возвратов · Период 2", "Δ количества возвратов",
            "Возвраты · Период 1", "Возвраты · Период 2", "Δ возвратов",
            "Доля возвратов · Период 1, %", "Доля возвратов · Период 2, %",
            "Выручка · Период 2", "Чистая выручка · Период 2",
        ],
    )

    with st.expander("Контроль сверки выгрузок", expanded=False):
        st.json({
            first_label: snapshot["validation"]["first"],
            second_label: snapshot["validation"]["second"],
        })


def _render_block_kpis(title: str, first: BlockReport, second: BlockReport, *, note: str) -> None:
    comparison = compare_totals(first.totals, second.totals)
    old = first.totals
    new = second.totals
    revenue_pct = float(comparison["revenue_pct"] or 0)
    quantity_pct = float(comparison["quantity_pct"] or 0)
    average_pct = float(comparison["average_price_pct"] or 0)
    discount_delta = float(comparison["discount_delta_pp"] or 0)
    net_pct = float(comparison["net_revenue_pct"] or 0)
    return_pct = float(comparison["return_amount_pct"] or 0)

    st.markdown(f"## {title}")
    st.markdown(
        '<div class="management-section-note">' + escape(note) + '</div>',
        unsafe_allow_html=True,
    )
    cards = [
        _metric_card("Выручка", _money(new.revenue), _pct(revenue_pct), revenue_pct),
        _metric_card("Количество", _quantity(new.quantity), _pct(quantity_pct), quantity_pct),
        _metric_card("Средняя цена", _money(new.average_price), _pct(average_pct), average_pct),
        _metric_card("Средняя скидка", f"{new.discount_pct:.1f}%", _pp(discount_delta), discount_delta),
        _metric_card("Чистая выручка", _money(new.net_revenue), _pct(net_pct), net_pct),
        _metric_card("Сумма возвратов", _money(new.return_amount), _pct(return_pct), return_pct),
        _metric_card(
            "Возвратов, шт.",
            _quantity(new.return_quantity),
            _pct(comparison["return_quantity_pct"]),
            float(comparison["return_quantity_pct"] or 0),
        ),
        _metric_card(
            "Доля возвратов",
            f"{float(comparison['return_share_second']):.1f}%",
            _pp(float(comparison["return_share_second"]) - float(comparison["return_share_first"])),
            (float(comparison["return_share_second"]) - float(comparison["return_share_first"])),
        ),
    ]
    st.markdown('<div class="management-kpi-grid">' + "".join(cards) + '</div>', unsafe_allow_html=True)


def _render_small_dimensions(
    tabs_spec: list[tuple[str, pd.DataFrame, str]],
    first_label: str,
    second_label: str,
    *,
    key_prefix: str,
) -> None:
    tabs = st.tabs([title for title, _, _ in tabs_spec])
    for tab, (title, frame, noun) in zip(tabs, tabs_spec):
        with tab:
            st.markdown(
                f'<div class="management-section-note">{escape(_section_summary(frame, noun))}</div>',
                unsafe_allow_html=True,
            )
            if frame.empty:
                st.info("В выбранном разрезе нет данных.")
                continue
            left, right = st.columns(2)
            with left:
                st.plotly_chart(
                    _comparison_chart(frame, f"{title}: выручка", first_label, second_label),
                    width="stretch", config={"displayModeBar": False}, key=f"{key_prefix}_{title}_compare",
                )
            with right:
                st.plotly_chart(
                    _delta_chart(frame, f"{title}: изменение"),
                    width="stretch", config={"displayModeBar": False}, key=f"{key_prefix}_{title}_delta",
                )
            _table(
                frame.sort_values("Выручка · Период 2", ascending=False),
                key=f"{key_prefix}_{title}_table", limit=None, columns=TECHNICAL_TABLE_COLUMNS,
            )


def _render_three_block_report(payload: dict[str, object]) -> None:
    first: dict[str, BlockReport] = payload["first"]  # type: ignore[assignment]
    second: dict[str, BlockReport] = payload["second"]  # type: ignore[assignment]
    first_label = first[SALES].meta.period_label
    second_label = second[SALES].meta.period_label

    st.markdown(
        f"**Сравниваются:** {escape(first_label)} ({first[SALES].meta.period_days} дней) → "
        f"{escape(second_label)} ({second[SALES].meta.period_days} дней)"
    )
    if st.button("Загрузить другие периоды", key="management_report_replace", width="stretch"):
        st.session_state.pop(REPORT_STATE_KEY, None)
        for key in (
            "management_sales_1", "management_sales_2",
            "management_consultants_1", "management_consultants_2",
            "management_suppliers_1", "management_suppliers_2",
        ):
            st.session_state.pop(key, None)
        st.rerun()

    st.markdown("# Аналитика трех независимых блоков")
    st.caption(
        "Сначала каждый источник анализируется самостоятельно. Общий total, количество и возвраты "
        "в каждом блоке берутся только из явной строки «Итого» 1С, а не из суммы округленных групповых строк."
    )

    # Block 1: store/general sales.
    _render_block_kpis(
        "1. Продажи по магазинам",
        first[SALES], second[SALES],
        note=(
            "Источник истины для общей выручки, количества, возвратов и анализа магазинов. "
            "63 Retail и 63 Timing остаются раздельными."
        ),
    )
    stores = compare_metric_maps(first[SALES].dimensions.get("stores", {}), second[SALES].dimensions.get("stores", {}))
    _render_dimension_section(
        "Магазины", stores, first_label, second_label,
        noun="магазинов", key="management_blocks_stores", table_columns=TECHNICAL_TABLE_COLUMNS,
    )
    sales_stones = compare_metric_maps(first[SALES].dimensions.get("stones", {}), second[SALES].dimensions.get("stones", {}))
    sales_assays = compare_metric_maps(first[SALES].dimensions.get("assays", {}), second[SALES].dimensions.get("assays", {}))
    sales_categories = compare_metric_maps(first[SALES].dimensions.get("categories", {}), second[SALES].dimensions.get("categories", {}))
    _render_small_dimensions(
        [
            ("Камни и вставки", sales_stones, "вставок"),
            ("Пробы", sales_assays, "проб"),
            ("Номенклатурные группы", sales_categories, "групп"),
        ],
        first_label, second_label, key_prefix="management_blocks_sales",
    )

    # Block 2: consultants. Do not drop Admin/blank rows from the control total.
    _render_block_kpis(
        "2. Продажи по консультантам",
        first[CONSULTANTS], second[CONSULTANTS],
        note=(
            "Все строки уровня «Менеджер» учитываются, включая Admin и «Менеджер не указан». "
            "Это исключает прежнее занижение общего total и количества продаж."
        ),
    )
    consultants = compare_metric_maps(
        first[CONSULTANTS].dimensions.get("consultants", {}),
        second[CONSULTANTS].dimensions.get("consultants", {}),
    )
    _render_dimension_section(
        "Консультанты", consultants, first_label, second_label,
        noun="консультантов", key="management_blocks_consultants",
        table_columns=[
            *TECHNICAL_TABLE_COLUMNS,
            "Чистая выручка · Период 1", "Чистая выручка · Период 2", "Δ чистой выручки",
            "Количество возвратов · Период 1", "Количество возвратов · Период 2", "Δ количества возвратов",
            "Возвраты · Период 1", "Возвраты · Период 2",
            "Доля возвратов · Период 1, %", "Доля возвратов · Период 2, %",
        ],
    )
    consultant_assays = compare_metric_maps(first[CONSULTANTS].dimensions.get("assays", {}), second[CONSULTANTS].dimensions.get("assays", {}))
    consultant_categories = compare_metric_maps(first[CONSULTANTS].dimensions.get("categories", {}), second[CONSULTANTS].dimensions.get("categories", {}))
    _render_small_dimensions(
        [
            ("Пробы у консультантов", consultant_assays, "проб"),
            ("Группы у консультантов", consultant_categories, "групп"),
        ],
        first_label, second_label, key_prefix="management_blocks_consultants_detail",
    )

    # Block 3: suppliers, sourced only from the dedicated supplier export.
    _render_block_kpis(
        "3. Продажи по поставщикам",
        first[SUPPLIERS], second[SUPPLIERS],
        note=(
            "Поставщики читаются только из отдельного отчета по поставщикам. "
            "Сайт больше не угадывает поставщика по SKU или похожему семейству артикулов."
        ),
    )
    suppliers = compare_metric_maps(
        first[SUPPLIERS].dimensions.get("suppliers", {}),
        second[SUPPLIERS].dimensions.get("suppliers", {}),
    )
    own = suppliers.loc[suppliers["Позиция"] == "Own production"] if not suppliers.empty else pd.DataFrame()
    if not own.empty:
        row = own.iloc[0]
        st.markdown(
            '<div class="management-summary"><h3>Own production</h3><p>'
            + escape(
                f"{first_label}: {_quantity(row['Количество · Период 1'])} изделий, "
                f"{_money(row['Выручка · Период 1'])}. {second_label}: "
                f"{_quantity(row['Количество · Период 2'])} изделий, "
                f"{_money(row['Выручка · Период 2'])}."
            )
            + '</p></div>',
            unsafe_allow_html=True,
        )
    _render_dimension_section(
        "Поставщики", suppliers, first_label, second_label,
        noun="поставщиков", key="management_blocks_suppliers", table_columns=TECHNICAL_TABLE_COLUMNS,
    )
    supplier_categories = compare_metric_maps(first[SUPPLIERS].dimensions.get("categories", {}), second[SUPPLIERS].dimensions.get("categories", {}))
    _render_small_dimensions(
        [("Номенклатурные группы поставщиков", supplier_categories, "групп")],
        first_label, second_label, key_prefix="management_blocks_supplier_detail",
    )

    # Три независимых блока завершают управленческий отчет.


def render_management_report_dashboard() -> None:
    _css()
    parsed = st.session_state.get(REPORT_STATE_KEY)
    if not parsed:
        _render_upload()
        return
    try:
        if isinstance(parsed, dict) and parsed.get("format") == "three_blocks":
            _render_three_block_report(parsed)
        else:
            first, second = parsed
            _render_report(first, second)
    finally:
        gc.collect()
