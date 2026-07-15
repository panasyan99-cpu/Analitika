from pathlib import Path


def source() -> str:
    return Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")


def test_comparison_submit_commits_state_before_rerun():
    text = source()
    submit_pos = text.index('submitted = st.form_submit_button(')
    ready_pos = text.index('st.session_state["comparison_ready"] = True', submit_pos)
    rerun_pos = text.index('st.rerun()', ready_pos)
    assert submit_pos < ready_pos < rerun_pos


def test_navigation_is_rendered_from_persisted_ready_state():
    text = source()
    ready_pos = text.index('ready = bool(st.session_state.get("comparison_ready"))')
    nav_pos = text.index('sidebar_navigation(both_loaded, comparison=True)', ready_pos)
    assert ready_pos < nav_pos
    assert 'mobile_navigation(both_loaded, comparison=True)' in text


def test_release_history_contains_navigation_fix():
    changelog = Path(__file__).with_name('CHANGELOG.md').read_text(encoding='utf-8')
    assert '## 1.1.13 — Comparison navigation state fix' in changelog
