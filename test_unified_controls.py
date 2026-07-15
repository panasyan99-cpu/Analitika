from pathlib import Path


def test_all_streamlit_controls_share_black_gold_theme():
    text = Path("streamlit_app.py").read_text(encoding="utf-8")
    assert 'div.stButton > button,' in text
    assert 'div.stDownloadButton > button,' in text
    assert '[data-testid="stFormSubmitButton"] button,' in text
    assert '[data-testid="stSegmentedControl"] button,' in text
    assert 'background:linear-gradient(135deg,#0d0b08' in text
    assert 'border-color:#d4a95c' in text
    assert 'min-height:44px' in text


def test_mobile_navigation_uses_same_dark_gold_style():
    text = Path("streamlit_app.py").read_text(encoding="utf-8")
    mobile = text[text.index('.mobile-nav a,'):text.index('[data-testid="stHorizontalBlock"]', text.index('.mobile-nav a,'))]
    assert '#0d0b08' in mobile
    assert '#f5ead8' in mobile
    assert '#d4a95c' in mobile
