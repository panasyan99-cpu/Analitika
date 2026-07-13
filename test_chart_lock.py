import plotly.graph_objects as go

from streamlit_app import LOCKED_CHART_CONFIG, lock_chart_interactions


def test_chart_config_disables_mutating_controls_but_keeps_hover():
    assert LOCKED_CHART_CONFIG["displayModeBar"] is False
    assert LOCKED_CHART_CONFIG["scrollZoom"] is False
    assert LOCKED_CHART_CONFIG["doubleClick"] is False
    assert LOCKED_CHART_CONFIG["editable"] is False
    assert LOCKED_CHART_CONFIG["staticPlot"] is False


def test_cartesian_axes_and_legend_are_locked():
    fig = go.Figure(go.Bar(x=[1, 2], y=[3, 4]))
    locked = lock_chart_interactions(fig)
    assert locked.layout.dragmode is False
    assert locked.layout.clickmode == "event"
    assert locked.layout.hovermode == "closest"
    assert locked.layout.xaxis.fixedrange is True
    assert locked.layout.yaxis.fixedrange is True
    assert locked.layout.legend.itemclick is False
    assert locked.layout.legend.itemdoubleclick is False


def test_pie_keeps_hover_without_creating_fixed_axes():
    fig = go.Figure(go.Pie(labels=["A", "B"], values=[2, 1]))
    locked = lock_chart_interactions(fig)
    assert locked.layout.dragmode is False
    assert locked.layout.hovermode == "closest"
    assert locked.layout.xaxis.fixedrange is None
    assert locked.layout.yaxis.fixedrange is None
