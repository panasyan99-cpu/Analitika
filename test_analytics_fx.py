from pathlib import Path
import json

ROOT = Path(__file__).parent


def test_global_fx_is_rendered_once_on_main_page():
    app = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    currency = (ROOT / "src" / "currency.py").read_text(encoding="utf-8")
    sonu = (ROOT / "src" / "sonu.py").read_text(encoding="utf-8")
    version = json.loads((ROOT / "version.json").read_text(encoding="utf-8"))

    assert version["version"] == "1.9.2"
    assert app.count("render_global_fx_control()") == 1
    assert "render_hero(active_mode)" in app
    assert app.index("render_hero(active_mode)") < app.index("render_global_fx_control()")
    assert "with st.sidebar" not in currency
    assert 'key="global_fx_compact"' in currency
    assert "Курс VND за 1 USD" in currency
    assert 'label_visibility="collapsed"' in currency
    assert "get_vnd_per_usd()" in sonu
    assert "render_global_fx_control" not in sonu


def test_global_fx_default_and_conversion():
    currency = (ROOT / "src" / "currency.py").read_text(encoding="utf-8")
    assert "DEFAULT_VND_PER_USD = 26_300" in currency
    assert "float(value or 0) / active_rate" in currency
