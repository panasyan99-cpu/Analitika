from __future__ import annotations

import json
from pathlib import Path

import pytest

import src.order_workflow as workflow
from src.order_workflow import ORDER_MODE_STONES, OrderDraft, _detect_columns, report_month_count, save_draft


def test_supplier_order_always_uses_four_months() -> None:
    assert report_month_count("01.01.2026 - 31.03.2026") == 4
    assert report_month_count("20.04.2026 - 20.07.2026") == 4
    assert report_month_count("") == 4


def test_missing_required_columns_stop_processing() -> None:
    rows = {7: {"E": "Продажи за период", "G": "Остатки"}, 8: {"G": "TT", "N": "Всего"}}
    with pytest.raises(ValueError, match="ТВП"):
        _detect_columns(rows)


def test_local_first_save_does_not_wait_for_cloud(monkeypatch: pytest.MonkeyPatch) -> None:
    calls = {"local": 0, "cloud": 0}
    monkeypatch.setattr(workflow, "_save_draft_locally", lambda payload: calls.__setitem__("local", calls["local"] + 1))

    class Storage:
        def save_draft(self, payload):
            calls["cloud"] += 1

    monkeypatch.setattr(workflow, "get_cloud_storage", lambda: Storage())
    draft = OrderDraft(source_hash="hash", source_name="report.xlsx", mode=ORDER_MODE_STONES)
    save_draft(draft, sync_cloud=False)
    assert calls == {"local": 1, "cloud": 0}
    save_draft(draft, sync_cloud=True)
    assert calls == {"local": 2, "cloud": 1}


def test_performance_configuration_is_enabled() -> None:
    config = Path(".streamlit/config.toml").read_text(encoding="utf-8")
    assert "maxUploadSize = 150" in config
    assert "fastReruns = true" in config


def test_ui_uses_batched_forms_and_lazy_excel() -> None:
    source = Path(workflow.__file__).read_text(encoding="utf-8")
    assert "Применить количество" in source
    assert "Применить замок" in source
    assert "Применить размеры" in source
    assert "Excel строится только по кнопке" in source
    assert "CLOUD_AUTOSAVE_INTERVAL_SECONDS" in source


def test_release_metadata_is_current() -> None:
    version = json.loads(Path("version.json").read_text(encoding="utf-8"))
    assert version["version"] == "1.11.1"
