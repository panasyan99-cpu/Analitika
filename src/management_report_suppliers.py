from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache
import json
from pathlib import Path
import re
from typing import Mapping

MAPPING_PATH = Path(__file__).resolve().parents[1] / "data" / "management_supplier_mapping.json"
CLOUD_MAPPING_NAME = "management-report/supplier-overrides.json"
UNKNOWN_SUPPLIER = "Не определен"


def normalize_sku(value: object) -> str:
    return re.sub(r"\s+", "", str(value or "").strip().upper())


def _rule_keys(sku: str) -> tuple[str, ...]:
    value = normalize_sku(sku)
    keys: list[str] = []
    alpha = re.match(r"([A-Z]+)", value)
    if alpha and len(alpha.group(1)) >= 2:
        keys.append("ALPHA:" + alpha.group(1))
    tokens = [part for part in re.split(r"[-_/]+", value) if part]
    if tokens:
        keys.append("T1:" + tokens[0])
        if len(tokens) >= 2:
            keys.append("T2:" + "-".join(tokens[:2]))
        if len(tokens) >= 3:
            keys.append("T3:" + "-".join(tokens[:3]))
    compact = re.sub(r"[^A-Z0-9]", "", value)
    for length in range(3, min(9, len(compact) + 1)):
        keys.append(f"P{length}:" + compact[:length])
    keys.append("FAM:" + re.sub(r"\d+", "#", value))
    if tokens:
        keys.append("TF1:" + re.sub(r"\d+", "#", tokens[0]))
        if len(tokens) >= 2:
            keys.append("TF2:" + "-".join(re.sub(r"\d+", "#", part) for part in tokens[:2]))
    return tuple(dict.fromkeys(keys))


@dataclass(frozen=True)
class SupplierResolution:
    supplier: str
    source: str
    confidence: str


@dataclass(frozen=True)
class SupplierCatalog:
    exact: Mapping[str, str]
    family_rules: Mapping[str, tuple[str, int]]
    overrides: Mapping[str, str]
    suppliers: tuple[str, ...]

    def resolve(self, sku: object) -> SupplierResolution:
        normalized = normalize_sku(sku)
        if not normalized:
            return SupplierResolution(UNKNOWN_SUPPLIER, "unknown", "none")
        override = self.overrides.get(normalized)
        if override:
            return SupplierResolution(override, "manual", "confirmed")
        exact = self.exact.get(normalized)
        if exact:
            return SupplierResolution(exact, "exact", "confirmed")

        matches = [self.family_rules[key] for key in _rule_keys(normalized) if key in self.family_rules]
        suppliers = {supplier for supplier, _ in matches}
        if len(suppliers) == 1 and matches:
            supplier = next(iter(suppliers))
            support = max(item[1] for item in matches)
            return SupplierResolution(supplier, "learned_family", f"verified_family_{support}")
        return SupplierResolution(UNKNOWN_SUPPLIER, "unknown", "none")


@lru_cache(maxsize=1)
def _base_payload() -> dict:
    try:
        return json.loads(MAPPING_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, TypeError):
        return {"exact": {}, "family_rules": []}


def load_cloud_overrides() -> dict[str, str]:
    try:
        from src.order_persistence import CloudStorageError, get_cloud_storage
        storage = get_cloud_storage()
        if storage is None:
            return {}
        payload = storage.load_shared_json(CLOUD_MAPPING_NAME) or {}
    except (ImportError, OSError, TypeError, ValueError, RuntimeError):
        return {}
    raw = payload.get("mapping", payload) if isinstance(payload, dict) else {}
    if not isinstance(raw, Mapping):
        return {}
    return {
        normalize_sku(key): str(value).strip()
        for key, value in raw.items()
        if normalize_sku(key) and str(value or "").strip()
    }


def save_cloud_overrides(mapping: Mapping[str, str]) -> bool:
    normalized = {
        normalize_sku(key): str(value).strip()
        for key, value in mapping.items()
        if normalize_sku(key) and str(value or "").strip() and str(value).strip() != UNKNOWN_SUPPLIER
    }
    try:
        from src.order_persistence import get_cloud_storage
        storage = get_cloud_storage()
        if storage is None:
            return False
        current_payload = storage.load_shared_json(CLOUD_MAPPING_NAME) or {}
        current_raw = current_payload.get("mapping", current_payload) if isinstance(current_payload, dict) else {}
        merged = {
            normalize_sku(key): str(value).strip()
            for key, value in dict(current_raw).items()
            if normalize_sku(key) and str(value or "").strip()
        } if isinstance(current_raw, Mapping) else {}
        merged.update(normalized)
        storage.save_shared_json(
            CLOUD_MAPPING_NAME,
            {"schema_version": 1, "mapping": dict(sorted(merged.items()))},
        )
        return True
    except (ImportError, OSError, TypeError, ValueError, RuntimeError):
        return False


def load_supplier_catalog(*, session_overrides: Mapping[str, str] | None = None) -> SupplierCatalog:
    payload = _base_payload()
    exact = {
        normalize_sku(key): str(value).strip()
        for key, value in dict(payload.get("exact", {})).items()
        if normalize_sku(key) and str(value or "").strip()
    }
    family_rules: dict[str, tuple[str, int]] = {}
    for row in payload.get("family_rules", []):
        if not isinstance(row, Mapping):
            continue
        key = str(row.get("key", "")).strip()
        supplier = str(row.get("supplier", "")).strip()
        try:
            support = int(row.get("support", 0) or 0)
        except (TypeError, ValueError):
            support = 0
        if key and supplier and support >= 4:
            family_rules[key] = (supplier, support)

    overrides = load_cloud_overrides()
    if session_overrides:
        overrides.update({normalize_sku(key): str(value).strip() for key, value in session_overrides.items()})
    suppliers = sorted(set(exact.values()) | {value[0] for value in family_rules.values()} | set(overrides.values()))
    return SupplierCatalog(
        exact=exact,
        family_rules=family_rules,
        overrides=overrides,
        suppliers=tuple(suppliers),
    )
