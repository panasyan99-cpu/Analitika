from pathlib import Path


def test_release_history_is_separate_and_dynamic():
    text = Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")
    assert 'def render_release_history()' in text
    assert 'class="updates-scroll updates-scroll-standalone"' in text
    assert 'release_history_html(Path(__file__).with_name("CHANGELOG.md"))' in text
    assert 'latest_updates_html(Path(__file__).with_name("CHANGELOG.md"), 3)' in text
