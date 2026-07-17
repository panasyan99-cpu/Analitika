from pathlib import Path


ROOT = Path(__file__).parent


def source() -> str:
    return (ROOT / "streamlit_app.py").read_text(encoding="utf-8")


def test_sidebar_opens_again_after_collapse_on_current_streamlit():
    app = source()
    assert 'initial_sidebar_state="expanded"' in app
    assert '[data-testid="stExpandSidebarButton"]' in app
    assert '[data-testid="stSidebarCollapseButton"]' in app
    assert '[data-testid="stToolbar"] {' in app
    assert 'display:flex !important' in app
    assert 'visibility:visible !important' in app
    assert 'pointer-events:auto !important' in app
    assert 'z-index:1000000 !important' in app
    assert '[data-testid="stSidebarCollapsedControl"]' in app
    assert 'content:"МЕНЮ"' in app
    assert 'min-width:54px' in app


def test_sidebar_reopen_release_version():
    version = (ROOT / "version.json").read_text(encoding="utf-8")
    changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    assert '"version": "1.7.1"' in version
    assert '## 1.6.4 — Global metal filter across all modules' in changelog
