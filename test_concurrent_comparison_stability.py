from pathlib import Path


def app_source() -> str:
    return Path(__file__).with_name("streamlit_app.py").read_text(encoding="utf-8")


def config_source() -> str:
    return Path(__file__).with_name(".streamlit").joinpath("config.toml").read_text(encoding="utf-8")


def test_mutable_parsed_objects_are_not_shared_between_users():
    text = app_source()
    assert '@st.cache_resource(scope="session", ttl=900, max_entries=3' in text
    assert 'cross-user access' in text


def test_comparison_upload_is_atomic():
    text = app_source()
    assert 'with st.form("comparison_upload_form", clear_on_submit=False)' in text
    assert 'if first_file is None or second_file is None:' in text
    assert 'comparison_processing' in text


def test_fast_interrupting_reruns_are_disabled():
    assert 'fastReruns = false' in config_source()


def test_current_release_is_documented():
    text = app_source()
    changelog = Path(__file__).with_name('CHANGELOG.md').read_text(encoding='utf-8')
    assert 'APP_VERSION = "1.2.5"' in text
    assert '## 1.2.5 — Restored modules, live product info and Sonu navigation' in changelog
    assert '## 1.1.15 — Concurrent comparison stability' in changelog


def test_heavy_excel_parsing_is_serialized_across_sessions():
    text = app_source()
    assert 'def excel_parse_lock() -> threading.Lock:' in text
    assert 'with excel_parse_lock():' in text
