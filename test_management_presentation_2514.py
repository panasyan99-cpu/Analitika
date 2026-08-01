from __future__ import annotations

import ast
from pathlib import Path

from src.currency import round_usd_to_tens
from src.management_block_reports import CONSULTANTS, KIND_LABELS, SALES, SUPPLIERS


def test_usd_display_rounds_to_nearest_ten_half_away_from_zero():
    assert round_usd_to_tens(64_004) == 64_000
    assert round_usd_to_tens(64_005) == 64_010
    assert round_usd_to_tens(-64_005) == -64_010


def test_management_blocks_have_plain_business_names():
    assert KIND_LABELS == {
        SALES: "Продажи по магазинам",
        CONSULTANTS: "Продажи по консультантам",
        SUPPLIERS: "Продажи по поставщикам",
    }


def test_management_report_source_has_arbitrary_filename_copy_and_delta_styling():
    source = Path("src/management_report.py").read_text(encoding="utf-8")
    ast.parse(source)
    assert "Название Excel-файла может быть любым" in source
    assert "def _style_delta_columns" in source
    assert "round_usd_to_tens" in source
    assert "styled.format" in source


def test_main_site_source_has_global_comparison_delta_styling_and_ten_dollar_rounding():
    source = Path("streamlit_app.py").read_text(encoding="utf-8")
    ast.parse(source)
    assert "def style_delta_columns" in source
    assert "round_usd_to_tens" in source


def test_sonu_usd_displays_use_ten_dollar_rounding():
    source = Path("src/sonu.py").read_text(encoding="utf-8")
    ast.parse(source)
    assert "round_usd_to_tens" in source
    assert "numeric.map(round_usd_to_tens) if _is_money_column(column)" in source
