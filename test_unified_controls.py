from pathlib import Path


def test_non_navigation_actions_use_light_gold_with_white_text():
    text = Path("streamlit_app.py").read_text(encoding="utf-8")
    assert 'div.stButton > button,' in text
    assert 'div.stDownloadButton > button,' in text
    assert '[data-testid="stFormSubmitButton"] button,' in text
    assert '[data-testid="stSegmentedControl"] button,' in text
    assert 'background:linear-gradient(135deg,#e0bd78' in text
    assert 'color:#ffffff !important' in text
    assert 'min-height:44px' in text


def test_navigation_keeps_shared_dark_gold_style_only_in_navigation_scope():
    text = Path("streamlit_app.py").read_text(encoding="utf-8")
    assert '.sidebar-nav-item,' in text
    assert '[class*="st-key-sidebar_navigation_controls"]' in text
    assert '--sidebar-nav-bg:linear-gradient(135deg,#181006' in text
    assert '--sidebar-nav-current:linear-gradient(135deg,#c5903b' in text
    assert 'justify-content:center !important' in text
    assert '.mobile-nav-item,' in text
    assert '#d4a95c' in text
