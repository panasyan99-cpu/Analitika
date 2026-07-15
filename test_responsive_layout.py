from pathlib import Path


def source() -> str:
    return Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")


def test_mobile_navigation_is_conditional_and_present():
    text = source()
    assert 'def mobile_navigation(has_report: bool, *, comparison: bool = False)' in text
    assert 'class="mobile-nav-shell"' in text
    assert 'mobile_navigation(bool(active_files), comparison=False)' in text
    assert '("#summary", "📊 Сводка")' in text
    assert 'if has_report:' in text


def test_responsive_breakpoints_and_table_scroll_are_defined():
    text = source()
    assert '@media (max-width: 900px)' in text
    assert '@media (max-width: 820px)' in text
    assert '@media (max-width: 600px)' in text
    assert ':has(> [data-testid="stColumn"]:nth-child(3))' in text
    assert '[data-testid="stDataFrame"] { overflow-x:auto' in text
    assert '[data-baseweb="tab-list"]' in text


def test_release_history_includes_locked_responsive_stability_and_current():
    text = source()
    assert 'APP_VERSION = "1.2.0"' in text
    assert 'Analitika Web 1.1.11 — Comparison workspace' in text
    assert 'Analitika Web 1.1.10 — Executive brief clarity' in text
    assert 'Analitika Web 1.1.7 — Stability and memory optimization' in text
    assert 'Analitika Web 1.1.6 — Responsive mobile layout' in text
    assert 'Analitika Web 1.1.5 — Locked chart interactions' in text
