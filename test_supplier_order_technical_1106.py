from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace

import pytest

import src.diagnostics as diagnostics
import src.order_workflow as workflow
from src.order_persistence import CloudStorageError
from src.order_workflow import ORDER_MODE_STONES, OrderDraft


def test_runtime_dependencies_match_streamlit_160_contract() -> None:
    requirements = Path("requirements.txt").read_text(encoding="utf-8")
    assert "streamlit==1.60.0" in requirements
    assert "pyarrow==24.0.0" in requirements


def test_release_uses_timed_fragments_and_bounded_pages() -> None:
    source = Path(workflow.__file__).read_text(encoding="utf-8")
    assert "@st.fragment(run_every=CLOUD_AUTOSAVE_INTERVAL_SECONDS)" in source
    assert "def _render_order_stage_fragment" in source
    assert "def _render_ring_stage_fragment" in source
    assert workflow.ORDER_PAGE_SIZE == 10


def test_generated_files_keep_only_latest_payload_of_requested_kind(monkeypatch: pytest.MonkeyPatch) -> None:
    state = {
        "supplier_excel::hash::Камни::old": b"old-main",
        "supplier_excel::hash::Камни::new": b"new-main",
        "limited_excel::hash::Камни::current": b"limited",
        "supplier_excel::other::Камни::old": b"other",
    }
    monkeypatch.setattr(workflow, "st", SimpleNamespace(session_state=state))

    released = workflow._clear_generated_payloads(
        "hash",
        "Камни",
        keep_keys=("supplier_excel::hash::Камни::new",),
        kinds=("main",),
    )

    assert released == len(b"old-main")
    assert "supplier_excel::hash::Камни::old" not in state
    assert state["supplier_excel::hash::Камни::new"] == b"new-main"
    assert state["limited_excel::hash::Камни::current"] == b"limited"
    assert state["supplier_excel::other::Камни::old"] == b"other"


def test_failed_cloud_flush_remains_dirty_for_timed_retry(monkeypatch: pytest.MonkeyPatch) -> None:
    state: dict[str, object] = {}
    dummy_st = SimpleNamespace(session_state=state)
    monkeypatch.setattr(workflow, "st", dummy_st)
    monkeypatch.setattr(workflow, "diagnostic_event", lambda *args, **kwargs: None)
    monkeypatch.setattr(
        workflow,
        "save_draft",
        lambda draft, sync_cloud=True: (_ for _ in ()).throw(CloudStorageError("offline")),
    )

    draft = OrderDraft(source_hash="hash", source_name="report.xlsx", mode=ORDER_MODE_STONES)
    workflow._save_session_draft(draft, sync_cloud=True)

    assert state[workflow._draft_dirty_key(draft)] is True
    assert "облако временно недоступно" in str(state["supplier_order_save_status"])


def test_diagnostics_rotates_bounded_log(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    log = tmp_path / "diagnostics.jsonl"
    rotated = tmp_path / "diagnostics.1.jsonl"
    monkeypatch.setattr(diagnostics, "_RUNTIME", tmp_path)
    monkeypatch.setattr(diagnostics, "_LOG_FILE", log)
    monkeypatch.setattr(diagnostics, "_ROTATED_LOG_FILE", rotated)
    monkeypatch.setattr(diagnostics, "_MAX_LOG_BYTES", 80)

    diagnostics.diagnostic_event("first", payload="x" * 120)
    diagnostics.diagnostic_event("second", payload="y")

    assert rotated.exists()
    assert '"event": "first"' in rotated.read_text(encoding="utf-8")
    assert '"event": "second"' in log.read_text(encoding="utf-8")


def test_1106_release_metadata() -> None:
    version = json.loads(Path("version.json").read_text(encoding="utf-8"))
    assert version["version"] == "2.4.4"
    assert version["channel"] == "stable"


def test_four_month_supplier_workbook_parses_with_fixed_window(tmp_path: Path) -> None:
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "TDSheet"
    sheet["A1"] = "Продажи товаров за период 01.03.2026 - 30.06.2026"
    sheet["A2"] = "Поставщик(и): Y&J"
    sheet["E7"] = "Продажи за период"
    sheet["G7"] = "Остатки"
    sheet["O7"] = "ТВП"
    sheet["G8"] = "TT"
    sheet["H8"] = "AB"
    sheet["I8"] = "NTR1"
    sheet["J8"] = "NTR2"
    sheet["K8"] = "SCR"
    sheet["L8"] = "63"
    sheet["M8"] = "20"
    sheet["N8"] = "Всего"
    sheet["A11"] = "Set# TEST"
    sheet["A12"] = "SKE24A001"
    sheet["B12"] = "Ruby"
    sheet["C12"] = "Earrings"
    sheet["E12"] = 8
    sheet["G12"] = 0
    sheet["H12"] = 0
    sheet["I12"] = 0
    sheet["J12"] = 0
    sheet["K12"] = 0
    sheet["L12"] = 0
    sheet["M12"] = 0
    sheet["N12"] = 0
    sheet["O12"] = 0
    path = tmp_path / "four-month-order.xlsx"
    workbook.save(path)

    parsed = workflow.parse_order_workbook(path)

    assert parsed.period == "01.03.2026 - 30.06.2026"
    assert parsed.items
    assert all(item.report_months == 4 for item in parsed.items)
