from pathlib import Path


ROOT = Path(__file__).parent


def test_purity_filter_is_between_mode_selector_and_fx_control():
    source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    main = source[source.index("def main() -> None:"):]
    mode_pos = main.index('mode = st.segmented_control(')
    filter_pos = main.index('render_metal_filter_control()')
    fx_pos = main.index('render_global_fx_control()')
    report_pos = main.index('render_comparison_mode()')
    assert mode_pos < filter_pos < fx_pos < report_pos


def test_comparison_report_reuses_top_filter_without_duplicate_render():
    source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    comparison = source[source.index("def render_comparison_mode() -> None:"):source.index("def render_warehouse_mode() -> None:")]
    assert 'selected_metals = selected_metal_groups()' in comparison
    assert 'selected_metals = render_metal_filter_control()' not in comparison
    assert '<b>Фильтры по пробам</b>' in source
