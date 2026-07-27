from pathlib import Path


def source() -> str:
    return Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")


def test_comparison_submit_commits_state_before_rerun():
    text = source()
    submit_pos = text.index('submitted = st.form_submit_button(')
    ready_pos = text.index('st.session_state["comparison_ready"] = True', submit_pos)
    rerun_pos = text.index('st.rerun()', ready_pos)
    assert submit_pos < ready_pos < rerun_pos


def test_comparison_uses_full_width_and_has_inline_replace_action():
    text = source()
    comparison = text[text.index('def render_comparison_mode() -> None:'):text.index('def render_warehouse_mode() -> None:')]
    assert 'sidebar_navigation(' not in comparison
    assert 'mobile_navigation(' not in comparison
    assert '"Загрузить другие периоды"' in comparison
    assert 'clear_comparison_uploads()' in comparison


def test_release_history_contains_navigation_fix():
    changelog = Path(__file__).with_name('CHANGELOG.md').read_text(encoding='utf-8')
    assert '## 1.1.13 — Comparison navigation state fix' in changelog
