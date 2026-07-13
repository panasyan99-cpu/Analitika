from pathlib import Path


def source() -> str:
    return Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")


def test_first_click_forces_rerun_before_navigation_is_rendered():
    text = source()
    button_pos = text.index('key="start_comparison"')
    rerun_pos = text.index('st.rerun()', button_pos)
    nav_pos = text.index('sidebar_navigation(has_comparison_report, comparison=True)', rerun_pos)
    assert button_pos < rerun_pos < nav_pos


def test_navigation_state_is_computed_after_button_action():
    text = source()
    assert 'has_comparison_report = ready and both_loaded' in text
    assert 'mobile_navigation(has_comparison_report, comparison=True)' in text
    assert 'ready = True' not in text[text.index('key="start_comparison"'):text.index('if not both_loaded:')]


def test_release_history_contains_navigation_fix():
    text = source()
    assert 'Analitika Web 1.1.13 — Comparison navigation state fix' in text
