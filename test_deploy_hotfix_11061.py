from pathlib import Path
import json

ROOT = Path(__file__).resolve().parent


def test_streamlit_pyarrow_pins_are_compatible():
    requirements = (ROOT / "requirements.txt").read_text(encoding="utf-8")
    assert "streamlit==1.60.0" in requirements
    assert "pyarrow==24.0.0" in requirements
    assert "pyarrow==25.0.0" not in requirements


def test_supplier_order_mode_uses_session_state_without_default_conflict():
    source = (ROOT / "src" / "order_workflow.py").read_text(encoding="utf-8")
    block_start = source.index('mode_key = "supplier_order_mode"')
    block_end = source.index("_flush_previous_mode_on_change", block_start)
    block = source[block_start:block_end]
    assert "st.session_state[mode_key] = ORDER_MODE_STONES" in block
    assert "key=mode_key" in block
    assert "default=" not in block


def test_hotfix_public_version():
    version = json.loads((ROOT / "version.json").read_text(encoding="utf-8"))
    assert version["version"] == "2.0"


def test_warehouse_annotations_are_not_postponed():
    source = (ROOT / "src" / "warehouse.py").read_text(encoding="utf-8")
    assert not source.startswith("from __future__ import annotations")
    assert "@dataclass(frozen=True)\nclass WarehouseConfig:" in source
