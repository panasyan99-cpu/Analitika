from __future__ import annotations

import json
import logging
import time
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterator

_RUNTIME = Path(__file__).resolve().parents[1] / ".runtime"
_LOG_FILE = _RUNTIME / "diagnostics.jsonl"
_LOGGER = logging.getLogger("analitika.diagnostics")

def diagnostic_event(event: str, **details: Any) -> None:
    payload = {
        "timestamp": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "event": str(event),
        **{str(key): value for key, value in details.items()},
    }
    try:
        _RUNTIME.mkdir(parents=True, exist_ok=True)
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
        diagnostic_event(event, ok=False, duration_ms=round((time.perf_counter()-started)*1000,1), error=str(exc), **details)
        raise
    else:
        diagnostic_event(event, ok=True, duration_ms=round((time.perf_counter()-started)*1000,1), **details)
