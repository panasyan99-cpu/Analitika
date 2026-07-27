from pathlib import Path


def test_anchor_and_button_navigation_share_one_visual_rule():
    text = Path("streamlit_app.py").read_text(encoding="utf-8")
    shared = (
        '.sidebar-nav-item:visited,\n'
        '[data-testid="stSidebar"] [class*="st-key-sidebar_navigation_controls"] div.stButton > button {'
    )
    assert shared in text
    assert 'justify-content:center !important' in text
    assert 'text-align:center !important' in text
    assert '--sidebar-nav-current:' in text


def test_sales_analysis_sidebar_is_removed_but_sonu_navigation_remains():
    app = Path("streamlit_app.py").read_text(encoding="utf-8")
    sonu = Path("src/sonu.py").read_text(encoding="utf-8")
    assert 'def sidebar_navigation(' not in app
    assert 'current=(item_id == "sonu-upload" and not has_report)' in sonu


def test_disabled_navigation_keeps_shared_geometry_and_readability():
    text = Path("streamlit_app.py").read_text(encoding="utf-8")
    assert 'opacity:.62 !important' in text
    assert 'min-height:44px !important' in text
    assert 'color:#d8cbb8 !important' in text
