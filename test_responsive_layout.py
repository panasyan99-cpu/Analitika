from pathlib import Path


ROOT = Path(__file__).parent


def source() -> str:
    return (ROOT / "streamlit_app.py").read_text(encoding="utf-8")


def test_mobile_navigation_is_shared_and_present():
    app = source()
    navigation = (ROOT / "src" / "navigation.py").read_text(encoding="utf-8")
    assert 'def mobile_navigation(has_report: bool, *, comparison: bool = False)' in app
    assert 'class="mobile-nav-shell"' in navigation
    assert 'mobile_navigation(bool(active_files), comparison=False)' in app
    assert '("summary", "Сводка", "#summary", has_report)' in app
    assert 'render_mobile_navigation(report_navigation_items' in app


def test_responsive_breakpoints_and_table_scroll_are_defined():
    text = source()
    assert '@media (max-width: 900px)' in text
    assert '@media (max-width: 820px)' in text
    assert '@media (max-width: 600px)' in text
    assert ':has(> [data-testid="stColumn"]:nth-child(3))' in text
    assert '[data-testid="stDataFrame"] { overflow-x:auto' in text
    assert '[data-baseweb="tab-list"]' in text
    assert '.st-key-global_fx_compact [data-testid="stHorizontalBlock"]' in text


def test_release_history_includes_locked_responsive_stability_and_current():
    changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    version = (ROOT / "version.json").read_text(encoding="utf-8")
    assert '"version": "1.7.4"' in version
    assert '## 1.3.0 — Full-page unified UX' in changelog
    assert '## 1.2.8 — Stone groups, unified controls and responsive audit' in changelog
    assert '## 1.1.11 — Comparison workspace' in changelog
    assert '## 1.1.10 — Executive brief clarity and table sorting' in changelog
    assert '## 1.1.7 — Stability and memory optimization' in changelog
    assert '## 1.1.6 — Responsive mobile layout' in changelog
    assert '## 1.1.5 — Locked chart interactions' in changelog
