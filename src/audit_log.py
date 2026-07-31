from __future__ import annotations

import json
import logging
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping

try:
    import streamlit as st
except ImportError:  # pragma: no cover - tests install a lightweight stub
    st = None  # type: ignore[assignment]

from src.diagnostics import diagnostic_event

_RUNTIME = Path(__file__).resolve().parents[1] / ".runtime"
_LOCAL_FILE = _RUNTIME / "action_audit.jsonl"
_MAX_LOCAL_BYTES = 4 * 1024 * 1024
_LOCK = threading.RLock()
_LOGGER = logging.getLogger("analitika.audit")
_PROCESS_SESSION_ID = uuid.uuid4().hex[:12]
_SENSITIVE_FRAGMENTS = (
    "password", "secret", "token", "authorization", "cookie", "access_key",
    "credential", "jwt", "session_key",
)


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")


def _session_id() -> str:
    if st is None:
        return _PROCESS_SESSION_ID
    try:
        state = st.session_state
        key = "_analitika_audit_session_id"
        if not state.get(key):
            state[key] = uuid.uuid4().hex[:12]
        return str(state[key])
    except Exception:
        return _PROCESS_SESSION_ID


def _clean(value: Any, *, depth: int = 0) -> Any:
    if depth > 3:
        return "…"
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, str):
        return value if len(value) <= 800 else value[:797] + "…"
    if isinstance(value, Mapping):
        result: dict[str, Any] = {}
        for raw_key, raw_value in value.items():
            key = str(raw_key)
            if any(fragment in key.casefold() for fragment in _SENSITIVE_FRAGMENTS):
                result[key] = "[REDACTED]"
            else:
                result[key] = _clean(raw_value, depth=depth + 1)
        return result
    if isinstance(value, (list, tuple, set)):
        items = list(value)
        return [_clean(item, depth=depth + 1) for item in items[:50]]
    return _clean(str(value), depth=depth + 1)


def _rotate_if_needed() -> None:
    try:
        if _LOCAL_FILE.exists() and _LOCAL_FILE.stat().st_size >= _MAX_LOCAL_BYTES:
            rotated = _LOCAL_FILE.with_name("action_audit.1.jsonl")
            rotated.unlink(missing_ok=True)
            _LOCAL_FILE.replace(rotated)
    except OSError:
        _LOGGER.exception("Unable to rotate audit log")


def _write_local(payload: Mapping[str, Any]) -> bool:
    try:
        with _LOCK:
            _RUNTIME.mkdir(parents=True, exist_ok=True)
            _rotate_if_needed()
            with _LOCAL_FILE.open("a", encoding="utf-8") as handle:
                handle.write(json.dumps(dict(payload), ensure_ascii=False, default=str) + "\n")
        return True
    except OSError:
        _LOGGER.exception("Unable to write local audit log")
        return False


def audit_event(
    action: str,
    *,
    module: str,
    result: str = "success",
    correlation_id: str = "",
    actor: str = "operator",
    persist_cloud: bool = True,
    **details: Any,
) -> dict[str, Any]:
    """Write one sanitized business-action event locally and, when available, to R2.

    Audit failures never turn a completed business operation into a failure. The
    returned payload contains storage flags so the diagnostics page can surface
    whether the durable copy was written.
    """
    payload: dict[str, Any] = {
        "event_id": uuid.uuid4().hex,
        "timestamp": _now_iso(),
        "module": str(module),
        "action": str(action),
        "result": str(result),
        "actor": str(actor or "operator"),
        "session_id": _session_id(),
        "correlation_id": str(correlation_id or ""),
        "details": _clean(details),
    }
    local_saved = _write_local(payload)
    cloud_saved = False
    cloud_error = ""
    if persist_cloud:
        try:
            from src.order_persistence import get_cloud_storage

            storage = get_cloud_storage()
            if storage is not None:
                storage.append_audit_event(payload)
                cloud_saved = True
        except Exception as exc:  # Audit is best-effort by design.
            cloud_error = str(exc)
            diagnostic_event("audit.cloud_write_error", action=action, error=cloud_error)
    payload["local_saved"] = local_saved
    payload["cloud_saved"] = cloud_saved
    if cloud_error:
        payload["cloud_error"] = cloud_error
    return payload


def read_local_audit_events(limit: int = 100) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for path in (_LOCAL_FILE.with_name("action_audit.1.jsonl"), _LOCAL_FILE):
        try:
            lines = path.read_text(encoding="utf-8").splitlines()
        except OSError:
            continue
        for line in lines:
            try:
                row = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(row, dict):
                rows.append(row)
    rows.sort(key=lambda row: str(row.get("timestamp", "")), reverse=True)
    return rows[: max(1, int(limit))]


def read_recent_audit_events(limit: int = 100, *, include_cloud: bool = True) -> list[dict[str, Any]]:
    rows = read_local_audit_events(limit=max(limit, 100))
    if include_cloud:
        try:
            from src.order_persistence import get_cloud_storage

            storage = get_cloud_storage()
            if storage is not None:
                rows.extend(storage.list_audit_events(limit=max(limit, 100)))
        except Exception:
            pass
    deduplicated: dict[str, dict[str, Any]] = {}
    for row in rows:
        key = str(row.get("event_id") or "") or json.dumps(row, sort_keys=True, default=str)
        deduplicated[key] = row
    result = list(deduplicated.values())
    result.sort(key=lambda row: str(row.get("timestamp", "")), reverse=True)
    return result[: max(1, int(limit))]
