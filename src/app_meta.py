from __future__ import annotations

import json
from pathlib import Path


def read_app_version() -> str:
    """Read the public app version from the single version.json source."""
    version_path = Path(__file__).resolve().parents[1] / "version.json"
    try:
        payload = json.loads(version_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, TypeError):
        return "1.4.0"
    value = str(payload.get("version", "")).strip()
    return value or "1.4.0"


APP_VERSION = read_app_version()
