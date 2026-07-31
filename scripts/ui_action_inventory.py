#!/usr/bin/env python3
"""Generate a static inventory of Streamlit actions and detect unsafe duplicate keys."""
from __future__ import annotations

import ast
import json
from collections import Counter
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
ACTION_METHODS = {
    "button",
    "form_submit_button",
    "download_button",
    "file_uploader",
}


def _literal(call: ast.Call, keyword: str) -> str:
    for item in call.keywords:
        if item.arg == keyword:
            try:
                return ast.literal_eval(item.value) if isinstance(ast.literal_eval(item.value), str) else ast.unparse(item.value)
            except Exception:
                return ast.unparse(item.value)
    return ""


def collect() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    paths = [ROOT / "streamlit_app.py", *sorted((ROOT / "src").rglob("*.py"))]
    for path in paths:
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Attribute):
                continue
            method = node.func.attr
            if method not in ACTION_METHODS:
                continue
            label = ast.unparse(node.args[0]) if node.args else ""
            rows.append(
                {
                    "file": str(path.relative_to(ROOT)),
                    "line": node.lineno,
                    "method": method,
                    "label_expression": label,
                    "key_expression": _literal(node, "key"),
                    "disabled_expression": _literal(node, "disabled"),
                }
            )
    return rows


def main() -> int:
    rows = collect()
    output = ROOT / "UI_ACTION_INVENTORY_2.6.0.json"
    output.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    literal_keys = [row["key_expression"] for row in rows if row["key_expression"] and not any(ch in row["key_expression"] for ch in "{}()")]
    duplicates = {key: count for key, count in Counter(literal_keys).items() if count > 1}
    missing_keys = [row for row in rows if row["method"] != "form_submit_button" and not row["key_expression"]]
    print(f"Actions: {len(rows)}")
    print(f"Actions without explicit key: {len(missing_keys)}")
    print(f"Duplicate literal keys: {len(duplicates)}")
    if missing_keys:
        for row in missing_keys:
            print(f"MISSING KEY {row['file']}:{row['line']} {row['method']} {row['label_expression']}")
    if duplicates:
        for key, count in sorted(duplicates.items()):
            print(f"DUPLICATE KEY {key}: {count}")
    return 1 if missing_keys or duplicates else 0


if __name__ == "__main__":
    raise SystemExit(main())
