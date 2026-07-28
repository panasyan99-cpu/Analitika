from pathlib import Path

ROOT = Path(__file__).parent
SOURCE = (ROOT / "src" / "order_workflow.py").read_text(encoding="utf-8")


def test_new_order_upload_is_above_orders_button():
    upload_start = SOURCE.index('st.markdown("## Новый заказ")')
    uploader = SOURCE.index('st.file_uploader(', upload_start)
    orders_button = SOURCE.index('"Заказы",', uploader)
    library_render = SOURCE.index('_render_saved_order_library()', orders_button)
    assert upload_start < uploader < orders_button < library_render


def test_all_analitika_orders_are_shown_without_completed_toggle():
    library_start = SOURCE.index('def _render_saved_order_library')
    library_end = SOURCE.index('def _render_upload', library_start)
    block = SOURCE[library_start:library_end]
    assert 'include_completed=True' in block
    assert 'Показать завершённые' not in block


def test_analitika_orders_are_rendered_before_manual_orders():
    library_start = SOURCE.index('def _render_saved_order_library')
    library_end = SOURCE.index('def _render_upload', library_start)
    block = SOURCE[library_start:library_end]
    analitika = block.index('Заказы, созданные в Analitika')
    workspace_loop = block.index('for index, workspace in enumerate(workspaces)')
    manual = block.rindex('_render_manual_transit_orders()')
    assert analitika < workspace_loop < manual
    assert '### Остальные заказы' in SOURCE
