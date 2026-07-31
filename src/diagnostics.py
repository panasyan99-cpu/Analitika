from __future__ import annotations

import json
import logging
import threading
import time
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterator

_RUNTIME = Path(__file__).resolve().parents[1] / ".runtime"
_LOG_FILE = _RUNTIME / "diagnostics.jsonl"
_ROTATED_LOG_FILE = _RUNTIME / "diagnostics.1.jsonl"
_MAX_LOG_BYTES = 2 * 1024 * 1024
_LOGGER = logging.getLogger("analitika.diagnostics")
_LOG_LOCK = threading.Lock()


def _rotate_log_if_needed() -> None:
    try:
        if not _LOG_FILE.exists() or _LOG_FILE.stat().st_size < _MAX_LOG_BYTES:
            return
        _ROTATED_LOG_FILE.unlink(missing_ok=True)
        _LOG_FILE.replace(_ROTATED_LOG_FILE)
    except OSError:
        _LOGGER.exception("Unable to rotate diagnostics")


def diagnostic_event(event: str, **details: Any) -> None:
    payload = {
        "timestamp": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "event": str(event),
        **{str(key): value for key, value in details.items()},
    }
    try:
        with _LOG_LOCK:
            _RUNTIME.mkdir(parents=True, exist_ok=True)
            _rotate_log_if_needed()
            with _LOG_FILE.open("a", encoding="utf-8") as handle:
                handle.write(json.dumps(payload, ensure_ascii=False, default=str) + "\n")
    except OSError:
        _LOGGER.exception("Unable to write diagnostics")


@contextmanager
def timed_operation(event: str, **details: Any) -> Iterator[None]:
    started = time.perf_counter()
    try:
        yield
    except Exception as exc:
        diagnostic_event(
            event,
            ok=False,
            duration_ms=round((time.perf_counter() - started) * 1000, 1),
            error=str(exc),
            **details,
        )
        raise
    else:
        diagnostic_event(
            event,
            ok=True,
            duration_ms=round((time.perf_counter() - started) * 1000, 1),
            **details,
        )


def read_diagnostic_events(limit: int = 100) -> list[dict[str, Any]]:
    """Read recent local technical events for the diagnostics workspace."""
    rows: list[dict[str, Any]] = []
    for path in (_ROTATED_LOG_FILE, _LOG_FILE):
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
