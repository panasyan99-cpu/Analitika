from pathlib import Path


def test_global_fx_release_is_full_and_clean():
    root = Path(__file__).parent
    app = (root / "streamlit_app.py").read_text(encoding="utf-8")
    currency = (root / "src" / "currency.py").read_text(encoding="utf-8")
    sonu = (root / "src" / "sonu.py").read_text(encoding="utf-8")

    assert 'APP_VERSION = "1.2.3"' in app
    assert 'GLOBAL_FX_SESSION_KEY = "site_vnd_per_usd"' in currency
    assert 'DEFAULT_VND_PER_USD = 26_300' in currency
    assert 'key=GLOBAL_FX_SESSION_KEY' in currency
    assert 'render_global_fx_control' in app
    assert 'from src.currency import render_global_fx_control' in sonu
    assert 'key="analytics_fx_rate"' not in app
    assert 'key="sonu_fx_rate"' not in sonu


def test_standard_and_comparison_money_use_global_rate():
    root = Path(__file__).parent
    app = (root / "streamlit_app.py").read_text(encoding="utf-8")
    assert 'return get_vnd_per_usd()' in app
    assert 'display[col] = pd.to_numeric(display[col], errors="coerce").fillna(0) / analytics_fx_rate()' in app
    assert 'return render_global_fx_control()' in app
