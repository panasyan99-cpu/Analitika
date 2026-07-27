from pathlib import Path


def test_all_active_workspaces_are_inline_and_sidebar_free():
    app = Path("streamlit_app.py").read_text(encoding="utf-8")
    sonu = Path("src/sonu.py").read_text(encoding="utf-8")
    warehouse = Path("src/warehouse.py").read_text(encoding="utf-8")
    order = Path("src/order_workflow.py").read_text(encoding="utf-8")
    assert 'initial_sidebar_state="collapsed"' in app
    assert 'display:none !important; visibility:hidden !important;' in app
    assert '_sonu_sidebar_navigation(' not in sonu[sonu.index("def render_sonu_order_dashboard"): ]
    assert 'status_slot = render_navigation()' not in warehouse[warehouse.index("def render_warehouse_dashboard"): ]
    assert '_render_sidebar(parsed, draft)' not in order[order.index("def render_supplier_order_dashboard"): ]
