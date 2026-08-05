from __future__ import annotations

import ast
from pathlib import Path


def _function_source(module_source: str, function_name: str) -> str:
    tree = ast.parse(module_source)
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == function_name:
            segment = ast.get_source_segment(module_source, node)
            assert segment is not None
            return segment
    raise AssertionError(f"Function {function_name} not found")


def test_three_block_report_ends_after_supplier_analytics():
    source = Path("src/management_report.py").read_text(encoding="utf-8")
    body = _function_source(source, "_render_three_block_report")

    assert '"1. Продажи по магазинам"' in body
    assert '"2. Продажи по консультантам"' in body
    assert '"3. Продажи по поставщикам"' in body
    assert "Комплексный анализ итогов трех блоков" not in body
    assert "Итоговое управленческое резюме" not in body
    assert "management_blocks_control_first" not in body
    assert "management_combined_growth" not in body
    assert "management_combined_decline" not in body


def test_upload_still_validates_three_sources_before_rendering():
    source = Path("src/management_report.py").read_text(encoding="utf-8")
    upload = _function_source(source, "_render_upload")

    assert "validate_period_bundle" in upload
    assert "Загрузите все шесть файлов" in upload
    assert "Итоговый объединённый блок сравнения не строится" in upload


def test_obsolete_combined_driver_renderer_was_removed():
    source = Path("src/management_report.py").read_text(encoding="utf-8")
    tree = ast.parse(source)
    function_names = {
        node.name for node in tree.body if isinstance(node, ast.FunctionDef)
    }
    assert "_render_combined_drivers" not in function_names
