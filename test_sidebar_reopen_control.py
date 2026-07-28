from pathlib import Path

ROOT = Path(__file__).parent


def source() -> str:
    return (ROOT / "streamlit_app.py").read_text(encoding="utf-8")


def test_legacy_sidebar_is_removed_in_visual_system():
    app = source()
    assert 'initial_sidebar_state="collapsed"' in app
    assert '/* No legacy black sidebar in any 2.0 workspace. */' in app
    assert '[data-testid="stSidebar"],' in app
    assert 'display:none !important; visibility:hidden !important;' in app


def test_visual_system_release_version():
    version = (ROOT / "version.json").read_text(encoding="utf-8")
    changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    assert '"version": "2.4.1"' in version
    assert '## 2.0 — Unified visual system' in changelog
