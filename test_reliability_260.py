from __future__ import annotations

import ast
import json
from dataclasses import dataclass
from pathlib import Path
from types import SimpleNamespace

import pytest

import src.audit_log as audit_log
import src.order_workflow as workflow
import src.sonu as sonu
from src.order_workflow import DraftPersistenceResult, ORDER_MODE_STONES, OrderDraft
from src.warehouse_management.client import WarehouseClientError
from src.warehouse_management.models import Product
from src.warehouse_management.schema import SILVER_LINE_FIELDS
from src.warehouse_management.service import WarehouseService


ROOT = Path(__file__).resolve().parent


def test_failed_local_and_cloud_save_is_blocking_and_never_claims_success(monkeypatch: pytest.MonkeyPatch) -> None:
    state: dict[str, object] = {}
    monkeypatch.setattr(workflow, "st", SimpleNamespace(session_state=state))
    monkeypatch.setattr(workflow, "diagnostic_event", lambda *args, **kwargs: None)
    monkeypatch.setattr(
        workflow,
        "persist_draft",
        lambda draft, sync_cloud=True: DraftPersistenceResult(
            saved_at=draft.updated_at,
            local_saved=False,
            cloud_configured=True,
            cloud_saved=False,
            local_error="disk full",
            cloud_error="offline",
        ),
    )
    draft = OrderDraft(source_hash="hash", source_name="report.xlsx", mode=ORDER_MODE_STONES)
    assert workflow._save_session_draft(draft, sync_cloud=True) is False
    assert state[workflow._draft_dirty_key(draft)] is True
    assert state["supplier_order_save_blocked"] is True
    assert str(state["supplier_order_save_status"]).startswith("НЕ СОХРАНЕНО")


def test_audit_log_redacts_secrets(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(audit_log, "_RUNTIME", tmp_path)
    monkeypatch.setattr(audit_log, "_LOCAL_FILE", tmp_path / "audit.jsonl")
    event = audit_log.audit_event(
        "test.action",
        module="tests",
        persist_cloud=False,
        password="secret",
        nested={"token": "abc", "safe": "visible"},
    )
    assert event["details"]["password"] == "[REDACTED]"
    assert event["details"]["nested"]["token"] == "[REDACTED]"
    assert event["details"]["nested"]["safe"] == "visible"
    assert "secret" not in (tmp_path / "audit.jsonl").read_text(encoding="utf-8")


def test_silver_schema_uses_exact_weight_receiving_field_names() -> None:
    assert "Расчётное количество по весу" in SILVER_LINE_FIELDS
    assert "Погрешность веса, г" in SILVER_LINE_FIELDS
    assert "Расчётное количество, шт." not in SILVER_LINE_FIELDS


@dataclass
class FakeConfig:
    souvenirs_table_id: int = 642
    components_table_id: int = 643
    operations_table_id: int = 644
    supplies_table_id: int = 645
    supply_lines_table_id: int = 646


class ResumableClient:
    def __init__(self) -> None:
        self.config = FakeConfig()
        self.rows: dict[int, list[dict[str, object]]] = {642: [], 643: [], 644: [], 645: [], 646: []}
        self.counter = 100
        self.fail_second_line_once = True

    def _id(self) -> int:
        self.counter += 1
        return self.counter

    def list_rows(self, table_id: int, refresh: bool = False):
        return [dict(row) for row in self.rows[int(table_id)]]

    def batch_id(self, prefix: str) -> str:
        return f"{prefix}-TEST"

    def create_row(self, table_id: int, payload: dict[str, object]):
        table_id = int(table_id)
        if table_id == 646 and str(payload.get("Строка поставки", "")).endswith("SKU-2") and self.fail_second_line_once:
            self.fail_second_line_once = False
            raise WarehouseClientError("temporary line failure")
        row = {"id": self._id(), **payload}
        self.rows[table_id].append(row)
        return dict(row)

    def batch_update(self, table_id: int, payloads):
        for payload in payloads:
            row_id = int(payload["id"])
            row = next(item for item in self.rows[int(table_id)] if int(item["id"]) == row_id)
            row.update(dict(payload))

    def upload_file(self, path: Path):
        return {"name": path.name}

    def create_operations(self, payloads, *, batch_id: str, command_id: str = ""):
        created = []
        for payload in payloads:
            created.append(self.create_row(644, {**payload, "Command ID": command_id, "Статус документа": "Создаётся"}))
        return created

    def mark_operations_status(self, rows, status: str):
        ids = {int(row["id"]) for row in rows}
        for row in self.rows[644]:
            if int(row["id"]) in ids:
                row["Статус документа"] = status


def _import_products() -> list[Product]:
    return [
        Product(number=1, boxes="A", sku="SKU-1", qty_document=5, description="One", category="Серьги", material="Brass", stone="", color="", unit_weight_kg=None, image_path=""),
        Product(number=2, boxes="B", sku="SKU-2", qty_document=7, description="Two", category="Подвески", material="Brass", stone="", color="", unit_weight_kg=None, image_path=""),
    ]


def test_supply_import_resumes_after_partial_failure_without_duplicates() -> None:
    client = ResumableClient()
    service = WarehouseService(client)
    with pytest.raises(WarehouseClientError, match="temporary line failure"):
        service.create_supply_from_products(
            supply_id="SUP-RESUME",
            supplier="Supplier",
            invoice="invoice.xlsx",
            comment="",
            products=_import_products(),
            section="Сувенирка",
            command_id="IMPORT-ONE",
        )
    assert len(client.rows[642]) == 2
    assert len(client.rows[646]) == 1
    assert client.rows[645][0]["Статус импорта"] == "Ошибка"

    plan = service.plan_supply_import(supply_id="SUP-RESUME", products=_import_products(), section="Сувенирка")
    assert plan["resume_lines"] == 1
    assert plan["new_lines"] == 1

    result = service.create_supply_from_products(
        supply_id="SUP-RESUME",
        supplier="Supplier",
        invoice="invoice.xlsx",
        comment="",
        products=_import_products(),
        section="Сувенирка",
        command_id="IMPORT-TWO",
    )
    assert result["command_id"] == "IMPORT-ONE"
    assert result["resumed"] >= 1
    assert len(client.rows[642]) == 2
    assert len(client.rows[646]) == 2
    assert client.rows[645][0]["Статус импорта"] == "Завершён"


def test_ui_action_inventory_and_critical_controls_exist() -> None:
    actions: list[tuple[str, int, str, str]] = []
    paths = [ROOT / "streamlit_app.py", *(ROOT / "src").rglob("*.py")]
    for path in paths:
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Attribute):
                continue
            if node.func.attr not in {"button", "form_submit_button", "download_button", "file_uploader"}:
                continue
            label = ast.unparse(node.args[0]) if node.args else ""
            actions.append((str(path.relative_to(ROOT)), node.lineno, node.func.attr, label))
    labels = "\n".join(row[3] for row in actions)
    assert len(actions) >= 105
    for critical in (
        "Проверить облако снова",
        "Исправить структуру Baserow",
        "Скачать журнал действий JSON",
        "Проверить всё снова",
    ):
        assert critical in labels
    warehouse_source = (ROOT / "src/warehouse_management/ui.py").read_text(encoding="utf-8")
    assert "Продолжить импорт поставки" in warehouse_source


def test_reliability_source_contracts_are_present() -> None:
    order = (ROOT / "src/order_workflow.py").read_text(encoding="utf-8")
    warehouse = (ROOT / "src/warehouse_management/ui.py").read_text(encoding="utf-8")
    sonu_source = (ROOT / "src/sonu.py").read_text(encoding="utf-8")
    assert "НЕ СОХРАНЕНО" in order
    assert "supplier_order_save_blocked" in order
    assert "CLOUD_STATUS_TTL_SECONDS" in (ROOT / "src/order_persistence.py").read_text(encoding="utf-8")
    assert "Открытие склада ничего не изменяет автоматически" in warehouse
    assert "План безопасного импорта" in warehouse
    assert "bracelet-overrides/entries" in sonu_source
