from pathlib import Path


def test_sonu_defines_locked_chart_config():
    text = Path('src/sonu.py').read_text(encoding='utf-8')
    assert 'LOCKED_CHART_CONFIG = {' in text
    assert 'config=LOCKED_CHART_CONFIG' in text


def test_metal_filter_is_rendered_for_sonu_too():
    text = Path('streamlit_app.py').read_text(encoding='utf-8')
    assert 'if mode != "Заказ Sonu"' not in text
    assert 'render_metal_filter_control(mode)' in text
    assert 'render_sonu_order_dashboard(selected_metal_groups())' in text
