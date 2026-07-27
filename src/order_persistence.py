from __future__ import annotations

import io
import json
import os
import threading
from dataclasses import dataclass
from datetime import datetime, timezone
from functools import lru_cache
from pathlib import Path
from typing import Any, Mapping, Sequence

import streamlit as st

try:
    import boto3
    from boto3.s3.transfer import TransferConfig
    from botocore.config import Config as BotocoreConfig
    from botocore.exceptions import BotoCoreError, ClientError
except ImportError:  # pragma: no cover - surfaced as a configuration error in UI
    boto3 = None
    TransferConfig = None
    BotocoreConfig = None
    BotoCoreError = Exception
    ClientError = Exception


class CloudStorageError(RuntimeError):
    """Raised when the durable object store cannot complete an operation."""


@dataclass(frozen=True)
class S3StorageConfig:
    endpoint_url: str
    access_key_id: str
    secret_access_key: str
    bucket: str
    region: str = "auto"
    prefix: str = "analitika/supplier-orders"
    required: bool = False

    @property
    def configured(self) -> bool:
        return bool(self.access_key_id and self.secret_access_key and self.bucket)


@dataclass(frozen=True)
class CloudStorageStatus:
    configured: bool
    available: bool
    required: bool
    backend_name: str
    message: str


MODE_FILE_NAMES = {
    "Камни": "stones",
    "Жемчуг": "pearls",
}

DELIVERY_STATUS_SENT = "sent"
DELIVERY_STATUS_APPROVED = "approved"
DELIVERY_STATUS_IN_PROGRESS = "in_progress"
DELIVERY_STATUS_RECEIVED = "received"
DELIVERY_STATUSES = (
    DELIVERY_STATUS_SENT,
    DELIVERY_STATUS_APPROVED,
    DELIVERY_STATUS_IN_PROGRESS,
    DELIVERY_STATUS_RECEIVED,
)
DELIVERY_DATE_FIELDS = {
    DELIVERY_STATUS_SENT: "sent_at",
    DELIVERY_STATUS_APPROVED: "approved_at",
    DELIVERY_STATUS_IN_PROGRESS: "in_progress_at",
    DELIVERY_STATUS_RECEIVED: "received_at",
}

_LOCK_GUARD = threading.Lock()
_WORKSPACE_LOCKS: dict[str, threading.RLock] = {}
_INDEX_LOCK = threading.RLock()


def _workspace_lock(source_hash: str) -> threading.RLock:
    key = str(source_hash).strip() or "__unknown__"
    with _LOCK_GUARD:
        return _WORKSPACE_LOCKS.setdefault(key, threading.RLock())


def _now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _safe_int(value: object) -> int:
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


def _normalize_delivery_status(value: object, *, received: object = False) -> str:
    status = str(value or "").strip()
    if status in DELIVERY_STATUSES:
        return status
    return DELIVERY_STATUS_RECEIVED if bool(received) else DELIVERY_STATUS_SENT


def _date_only(value: object) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    try:
        return datetime.fromisoformat(raw.replace("Z", "+00:00")).date().isoformat()
    except ValueError:
        return raw[:10] if len(raw) >= 10 else raw


def _normalize_delivery_dates(
    value: object,
    *,
    order_date: object = "",
    received_at: object = "",
    status: object = DELIVERY_STATUS_SENT,
    status_updated_at: object = "",
) -> dict[str, str]:
    raw = dict(value) if isinstance(value, Mapping) else {}
    result = {
        field: _date_only(raw.get(field, ""))
        for field in DELIVERY_DATE_FIELDS.values()
    }
    if not result["sent_at"]:
        result["sent_at"] = _date_only(order_date)
    normalized_status = _normalize_delivery_status(status)
    active_field = DELIVERY_DATE_FIELDS[normalized_status]
    if not result[active_field]:
        fallback = received_at if normalized_status == DELIVERY_STATUS_RECEIVED else status_updated_at
        result[active_field] = _date_only(fallback)
    if normalized_status == DELIVERY_STATUS_RECEIVED and not result["received_at"]:
        result["received_at"] = _date_only(received_at)
    return result


def _mapping_value(mapping: Mapping[str, Any] | None, name: str, default: object = "") -> object:
    if not mapping:
        return default
    try:
        value = mapping.get(name, default)
    except (AttributeError, KeyError, TypeError):
        return default
    return value


def _secret_section() -> Mapping[str, Any]:
    try:
        section = st.secrets.get("order_storage", {})
    except (FileNotFoundError, KeyError, TypeError):
        return {}
    return section if isinstance(section, Mapping) else {}


def _truthy(value: object) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "yes", "on", "да"}


@lru_cache(maxsize=1)
def load_storage_config() -> S3StorageConfig:
    section = _secret_section()
    return S3StorageConfig(
        endpoint_url=str(
            _mapping_value(section, "endpoint_url", os.getenv("ORDER_STORAGE_ENDPOINT_URL", ""))
        ).strip(),
        access_key_id=str(
            _mapping_value(section, "access_key_id", os.getenv("ORDER_STORAGE_ACCESS_KEY_ID", ""))
        ).strip(),
        secret_access_key=str(
            _mapping_value(section, "secret_access_key", os.getenv("ORDER_STORAGE_SECRET_ACCESS_KEY", ""))
        ).strip(),
        bucket=str(_mapping_value(section, "bucket", os.getenv("ORDER_STORAGE_BUCKET", ""))).strip(),
        region=str(_mapping_value(section, "region", os.getenv("ORDER_STORAGE_REGION", "auto"))).strip() or "auto",
        prefix=str(
            _mapping_value(section, "prefix", os.getenv("ORDER_STORAGE_PREFIX", "analitika/supplier-orders"))
        ).strip().strip("/")
        or "analitika/supplier-orders",
        required=_truthy(_mapping_value(section, "required", os.getenv("ORDER_STORAGE_REQUIRED", "false"))),
    )


def reset_storage_config_cache() -> None:
    load_storage_config.cache_clear()
    get_cloud_storage.cache_clear()
    get_cloud_storage_status.cache_clear()


class S3OrderStorage:
    """Durable supplier-order storage over any S3-compatible object store.

    Every source workbook has an isolated workspace prefix. A compact cloud
    index is maintained next to the workspaces so the order library can be
    rendered with one small JSON request rather than downloading every source
    workbook or every manifest.
    """

    def __init__(self, config: S3StorageConfig):
        if boto3 is None or TransferConfig is None or BotocoreConfig is None:
            raise CloudStorageError("Не установлен пакет boto3.")
        if not config.configured:
            raise CloudStorageError("Облачное хранилище заказов не настроено.")
        self.config = config
        client_kwargs: dict[str, Any] = {
            "service_name": "s3",
            "aws_access_key_id": config.access_key_id,
            "aws_secret_access_key": config.secret_access_key,
            "region_name": config.region,
            "config": BotocoreConfig(
                retries={"max_attempts": 8, "mode": "adaptive"},
                connect_timeout=20,
                read_timeout=180,
                s3={"addressing_style": "path"},
            ),
        }
        if config.endpoint_url:
            client_kwargs["endpoint_url"] = config.endpoint_url
        self.client = boto3.client(**client_kwargs)
        self.transfer_config = TransferConfig(
            multipart_threshold=8 * 1024 * 1024,
            multipart_chunksize=8 * 1024 * 1024,
            max_concurrency=4,
            use_threads=True,
        )

    def _key(self, suffix: str) -> str:
        return f"{self.config.prefix}/{suffix.lstrip('/')}"

    def _workspace_prefix(self, source_hash: str) -> str:
        return self._key(f"workspaces/{source_hash}")

    def manifest_key(self, source_hash: str) -> str:
        return f"{self._workspace_prefix(source_hash)}/manifest.json"

    def mode_manifest_key(self, source_hash: str, mode: str) -> str:
        mode_name = MODE_FILE_NAMES.get(mode, "draft")
        return f"{self._workspace_prefix(source_hash)}/manifest-{mode_name}.json"

    def index_key(self) -> str:
        return self._key("orders-index.json")

    def index_entry_key(self, source_hash: str) -> str:
        return self._key(f"index-entries/{source_hash}.json")

    def manual_order_prefix(self) -> str:
        return self._key("manual-orders/")

    def manual_order_key(self, order_id: str) -> str:
        clean_id = "".join(ch for ch in str(order_id) if ch.isalnum() or ch in {"-", "_"})
        if not clean_id:
            raise CloudStorageError("Не указан идентификатор ручного заказа.")
        return self._key(f"manual-orders/{clean_id}.json")

    def draft_key(self, source_hash: str, mode: str) -> str:
        mode_name = MODE_FILE_NAMES.get(mode, "draft")
        return f"{self._workspace_prefix(source_hash)}/draft-{mode_name}.json"

    def workbook_key(self, source_hash: str, source_name: str) -> str:
        suffix = Path(source_name).suffix.lower()
        if suffix not in {".xlsx", ".xlsm"}:
            suffix = ".xlsx"
        return f"{self._workspace_prefix(source_hash)}/source{suffix}"

    def check(self) -> None:
        try:
            self.client.head_bucket(Bucket=self.config.bucket)
        except (BotoCoreError, ClientError, OSError) as exc:
            raise CloudStorageError(f"Хранилище недоступно: {exc}") from exc

    def exists(self, key: str) -> bool:
        try:
            self.client.head_object(Bucket=self.config.bucket, Key=key)
            return True
        except ClientError as exc:
            status = int(exc.response.get("ResponseMetadata", {}).get("HTTPStatusCode", 0) or 0)
            code = str(exc.response.get("Error", {}).get("Code", ""))
            if status == 404 or code in {"404", "NoSuchKey", "NotFound"}:
                return False
            raise CloudStorageError(f"Не удалось проверить объект {key}: {exc}") from exc
        except (BotoCoreError, OSError) as exc:
            raise CloudStorageError(f"Не удалось проверить объект {key}: {exc}") from exc

    def put_bytes(self, key: str, payload: bytes, content_type: str) -> None:
        try:
            self.client.upload_fileobj(
                io.BytesIO(payload),
                self.config.bucket,
                key,
                ExtraArgs={"ContentType": content_type, "CacheControl": "no-store"},
                Config=self.transfer_config,
            )
        except (BotoCoreError, ClientError, OSError) as exc:
            raise CloudStorageError(f"Не удалось сохранить данные в облако: {exc}") from exc

    def get_bytes(self, key: str) -> bytes:
        try:
            response = self.client.get_object(Bucket=self.config.bucket, Key=key)
            return bytes(response["Body"].read())
        except (BotoCoreError, ClientError, OSError, KeyError) as exc:
            raise CloudStorageError(f"Не удалось загрузить сохранённые данные: {exc}") from exc

    def download_file(self, key: str, destination: Path) -> Path:
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary = destination.with_suffix(destination.suffix + ".download")
        try:
            self.client.download_file(
                self.config.bucket,
                key,
                str(temporary),
                Config=self.transfer_config,
            )
            temporary.replace(destination)
        except (BotoCoreError, ClientError, OSError) as exc:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass
            raise CloudStorageError(f"Не удалось восстановить исходный Excel: {exc}") from exc
        return destination

    def get_json(self, key: str) -> dict[str, Any] | None:
        if not self.exists(key):
            return None
        try:
            value = json.loads(self.get_bytes(key).decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise CloudStorageError(f"Сохранённые данные повреждены: {exc}") from exc
        return value if isinstance(value, dict) else None

    def put_json(self, key: str, payload: Mapping[str, Any]) -> None:
        body = json.dumps(dict(payload), ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        self.put_bytes(key, body, "application/json; charset=utf-8")

    def _index_entry_from_manifest(self, manifest: Mapping[str, Any]) -> dict[str, Any]:
        drafts_raw = manifest.get("drafts", {})
        drafts = dict(drafts_raw) if isinstance(drafts_raw, Mapping) else {}
        normalized_drafts: dict[str, dict[str, Any]] = {}
        for mode, details_raw in drafts.items():
            if mode not in MODE_FILE_NAMES or not isinstance(details_raw, Mapping):
                continue
            details = dict(details_raw)
            delivery_status = _normalize_delivery_status(
                details.get("delivery_status", ""),
                received=details.get("received", False),
            )
            normalized_drafts[str(mode)] = {
                "key": str(details.get("key", "")),
                "created_at": str(details.get("created_at", "")),
                "updated_at": str(details.get("updated_at", "")),
                "selected_positions": max(0, _safe_int(details.get("selected_positions", 0))),
                "total_quantity": max(0, _safe_int(details.get("total_quantity", 0))),
                "limited_positions": max(0, _safe_int(details.get("limited_positions", 0))),
                "stage": str(details.get("stage", "order")),
                "status": "completed" if str(details.get("status", "draft")) == "completed" else "draft",
                "delivery_status": delivery_status,
                "delivery_dates": _normalize_delivery_dates(
                    details.get("delivery_dates", {}),
                    order_date=details.get("updated_at", "") or details.get("created_at", ""),
                    received_at=details.get("received_at", ""),
                    status=delivery_status,
                    status_updated_at=details.get("status_updated_at", ""),
                ),
                "status_updated_at": str(details.get("status_updated_at", "")),
                # Compatibility fields retained for older clients and manifests.
                "received": delivery_status == DELIVERY_STATUS_RECEIVED,
                "received_at": str(details.get("received_at", "")),
            }
        statuses = [str(row.get("status", "draft")) for row in normalized_drafts.values()]
        workspace_status = "completed" if statuses and all(value == "completed" for value in statuses) else "draft"
        updated_at = str(manifest.get("updated_at", ""))
        created_at = str(manifest.get("created_at", ""))
        if not created_at:
            candidates = [
                str(row.get("created_at", "")) or str(row.get("updated_at", ""))
                for row in normalized_drafts.values()
                if str(row.get("created_at", "")) or str(row.get("updated_at", ""))
            ]
            created_at = min(candidates) if candidates else updated_at
        return {
            "source_hash": str(manifest.get("source_hash", "")),
            "source_name": str(manifest.get("source_name", "")),
            "workbook_key": str(manifest.get("workbook_key", "")),
            "workbook_size": max(0, _safe_int(manifest.get("workbook_size", 0))),
            "created_at": created_at,
            "updated_at": updated_at,
            "status": workspace_status,
            "drafts": normalized_drafts,
        }

    def _list_index_entries(self) -> dict[str, dict[str, Any]]:
        """Read independent per-workspace index rows.

        Each workspace owns a separate object, so two users saving different
        orders cannot overwrite one another. ``orders-index.json`` remains a
        compact compatibility cache and is rebuilt/merged from these rows.
        """
        prefix = self._key("index-entries/")
        result: dict[str, dict[str, Any]] = {}
        if not hasattr(self, "client"):
            return result
        continuation: str | None = None
        while True:
            kwargs: dict[str, Any] = {
                "Bucket": self.config.bucket,
                "Prefix": prefix,
                "MaxKeys": 1000,
            }
            if continuation:
                kwargs["ContinuationToken"] = continuation
            try:
                response = self.client.list_objects_v2(**kwargs)
            except (BotoCoreError, ClientError, OSError) as exc:
                raise CloudStorageError(f"Не удалось прочитать индекс заказов: {exc}") from exc
            for row in response.get("Contents", []):
                key = str(row.get("Key", ""))
                if not key.endswith(".json"):
                    continue
                payload = self.get_json(key)
                if not isinstance(payload, Mapping):
                    continue
                source_hash = str(payload.get("source_hash", "")).strip()
                if source_hash:
                    result[source_hash] = dict(payload)
            if not response.get("IsTruncated"):
                break
            continuation = str(response.get("NextContinuationToken", "")) or None
            if not continuation:
                break
        return result

    def _load_index(self, *, merge_entries: bool = False) -> dict[str, Any]:
        """Read the compact index with optional independent-entry recovery.

        Normal library opens perform one JSON read. Independent per-workspace
        rows are scanned only during recovery/refresh, avoiding one cloud request
        per historical order on every page open.
        """
        index = self.get_json(self.index_key()) or {}
        orders_raw = index.get("orders", {})
        orders = dict(orders_raw) if isinstance(orders_raw, Mapping) else {}
        if merge_entries:
            try:
                orders.update(self._list_index_entries())
            except CloudStorageError:
                pass
        return {
            "schema_version": 3,
            "updated_at": str(index.get("updated_at", "")) or _now_iso(),
            "orders": orders,
        }

    def _save_index(self, index: Mapping[str, Any]) -> None:
        payload = dict(index)
        payload["schema_version"] = 3
        payload["updated_at"] = _now_iso()
        self.put_json(self.index_key(), payload)

    def _upsert_index_from_manifest(self, manifest: Mapping[str, Any]) -> None:
        source_hash = str(manifest.get("source_hash", "")).strip()
        if not source_hash:
            return
        entry = self._index_entry_from_manifest(manifest)
        # Independent row is the durable recovery source. The compact index is
        # updated under a process lock so concurrent Streamlit sessions cannot
        # overwrite one another within the deployed app instance.
        self.put_json(self.index_entry_key(source_hash), entry)
        with _INDEX_LOCK:
            index = self._load_index()
            orders = dict(index.get("orders", {}))
            orders[source_hash] = entry
            index["orders"] = orders
            self._save_index(index)

    def save_workbook(self, source_hash: str, source_name: str, payload: bytes) -> dict[str, Any]:
        with _workspace_lock(source_hash):
            workbook_key = self.workbook_key(source_hash, source_name)
            if not self.exists(workbook_key):
                self.put_bytes(
                    workbook_key,
                    payload,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            now = _now_iso()
            manifest = self.get_json(self.manifest_key(source_hash)) or {}
            created_at = str(manifest.get("created_at", "")) or now
            manifest.update(
                {
                    "schema_version": 3,
                    "source_hash": source_hash,
                    "source_name": source_name,
                    "workbook_key": workbook_key,
                    "workbook_size": len(payload),
                    "created_at": created_at,
                    "updated_at": now,
                    "status": str(manifest.get("status", "draft")) or "draft",
                }
            )
            manifest.setdefault("drafts", {})
            self.put_json(self.manifest_key(source_hash), manifest)
            self._upsert_index_from_manifest(manifest)
            return manifest

    def save_draft(self, payload: Mapping[str, Any]) -> dict[str, Any]:
        source_hash = str(payload.get("source_hash", "")).strip()
        source_name = str(payload.get("source_name", "")).strip()
        mode = str(payload.get("mode", "")).strip()
        if not source_hash or mode not in MODE_FILE_NAMES:
            raise CloudStorageError("Черновик не содержит идентификатор отчёта или тип заказа.")
        with _workspace_lock(source_hash):
            now = str(payload.get("updated_at", "")) or _now_iso()
            created_at = str(payload.get("created_at", "")) or now
            draft_key = self.draft_key(source_hash, mode)
            self.put_json(draft_key, payload)
            orders = payload.get("orders", {})
            if not isinstance(orders, Mapping):
                orders = {}
            limited_orders = payload.get("limited_orders", {})
            if not isinstance(limited_orders, Mapping):
                limited_orders = {}
            previous_mode = self.get_json(self.mode_manifest_key(source_hash, mode)) or {}
            previous_status = "completed" if str(previous_mode.get("status", "draft")) == "completed" else "draft"
            status = "completed" if str(payload.get("status", "draft")) == "completed" else "draft"
            delivery_status = _normalize_delivery_status(
                previous_mode.get("delivery_status", ""),
                received=previous_mode.get("received", False),
            )
            mode_created_at = str(previous_mode.get("created_at", "")) or created_at
            delivery_dates = _normalize_delivery_dates(
                previous_mode.get("delivery_dates", {}),
                order_date=mode_created_at,
                received_at=previous_mode.get("received_at", ""),
                status=delivery_status,
                status_updated_at=previous_mode.get("status_updated_at", ""),
            )
            if status == "completed" and previous_status != "completed":
                delivery_status = DELIVERY_STATUS_SENT
                delivery_dates = {field: "" for field in DELIVERY_DATE_FIELDS.values()}
                delivery_dates["sent_at"] = _date_only(now)
            mode_details = {
                "mode": mode,
                "key": draft_key,
                "created_at": mode_created_at,
                "updated_at": now,
                "selected_positions": sum(1 for value in orders.values() if _safe_int(value) > 0),
                "total_quantity": sum(max(0, _safe_int(value)) for value in orders.values()),
                "limited_positions": sum(1 for value in limited_orders.values() if bool(value)),
                "stage": str(payload.get("stage", "order")),
                "status": status,
                "delivery_status": delivery_status,
                "delivery_dates": delivery_dates,
                "status_updated_at": str(previous_mode.get("status_updated_at", "")),
                "received": delivery_status == DELIVERY_STATUS_RECEIVED,
                "received_at": str(previous_mode.get("received_at", "")),
            }
            # Each mode owns an independent metadata object. If two sessions edit
            # Stones and Pearls at the same time, neither mode summary is lost.
            self.put_json(self.mode_manifest_key(source_hash, mode), mode_details)

            manifest = self.get_json(self.manifest_key(source_hash)) or {
                "schema_version": 3,
                "source_hash": source_hash,
                "source_name": source_name,
                "workbook_key": self.workbook_key(source_hash, source_name),
                "workbook_size": 0,
                "created_at": created_at,
                "drafts": {},
            }
            existing_drafts = manifest.get("drafts", {})
            existing_drafts = dict(existing_drafts) if isinstance(existing_drafts, Mapping) else {}
            drafts: dict[str, dict[str, Any]] = {}
            for candidate_mode in MODE_FILE_NAMES:
                if candidate_mode == mode:
                    candidate = mode_details
                else:
                    candidate = self.get_json(self.mode_manifest_key(source_hash, candidate_mode))
                    if candidate is None:
                        fallback = existing_drafts.get(candidate_mode)
                        candidate = dict(fallback) if isinstance(fallback, Mapping) else None
                if isinstance(candidate, Mapping):
                    drafts[candidate_mode] = dict(candidate)
            statuses = [str(row.get("status", "draft")) for row in drafts.values()]
            manifest["schema_version"] = 3
            manifest["source_hash"] = source_hash
            manifest["source_name"] = source_name or str(manifest.get("source_name", ""))
            manifest["created_at"] = str(manifest.get("created_at", "")) or created_at
            manifest["updated_at"] = max(
                [now, *(str(row.get("updated_at", "")) for row in drafts.values())]
            )
            manifest["drafts"] = drafts
            manifest["status"] = "completed" if statuses and all(value == "completed" for value in statuses) else "draft"
            self.put_json(self.manifest_key(source_hash), manifest)
            self._upsert_index_from_manifest(manifest)
            return manifest

    def set_mode_delivery_status(
        self,
        source_hash: str,
        mode: str,
        delivery_status: str,
        *,
        status_date: str = "",
        delivery_dates: Mapping[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Persist one operational status and its dated timeline."""
        source_hash = str(source_hash).strip()
        mode = str(mode).strip()
        normalized_status = _normalize_delivery_status(delivery_status)
        if not source_hash or mode not in MODE_FILE_NAMES:
            raise CloudStorageError("Не указан заказ или его тип.")
        with _workspace_lock(source_hash):
            mode_key = self.mode_manifest_key(source_hash, mode)
            mode_details = self.get_json(mode_key)
            if not isinstance(mode_details, Mapping):
                raise CloudStorageError("Сведения об этом типе заказа не найдены.")
            mode_details = dict(mode_details)
            if str(mode_details.get("status", "draft")) != "completed":
                raise CloudStorageError("Статус поставки можно менять только у завершённого заказа.")

            now = _now_iso()
            dates = _normalize_delivery_dates(
                delivery_dates if delivery_dates is not None else mode_details.get("delivery_dates", {}),
                order_date=mode_details.get("created_at", ""),
                received_at=mode_details.get("received_at", ""),
                status=mode_details.get("delivery_status", DELIVERY_STATUS_SENT),
                status_updated_at=mode_details.get("status_updated_at", ""),
            )
            active_field = DELIVERY_DATE_FIELDS[normalized_status]
            chosen_date = _date_only(status_date) or dates.get(active_field) or _date_only(now)
            dates[active_field] = chosen_date
            active_rank = DELIVERY_STATUSES.index(normalized_status)
            for later_status in DELIVERY_STATUSES[active_rank + 1:]:
                dates[DELIVERY_DATE_FIELDS[later_status]] = ""

            mode_details["delivery_status"] = normalized_status
            mode_details["delivery_dates"] = dates
            mode_details["status_updated_at"] = now
            mode_details["received"] = normalized_status == DELIVERY_STATUS_RECEIVED
            mode_details["received_at"] = dates.get("received_at", "") if mode_details["received"] else ""
            mode_details["updated_at"] = now
            self.put_json(mode_key, mode_details)

            manifest = self.get_json(self.manifest_key(source_hash)) or {}
            drafts_raw = manifest.get("drafts", {})
            drafts = dict(drafts_raw) if isinstance(drafts_raw, Mapping) else {}
            drafts[mode] = mode_details
            manifest["drafts"] = drafts
            manifest["updated_at"] = now
            statuses = [str(row.get("status", "draft")) for row in drafts.values() if isinstance(row, Mapping)]
            manifest["status"] = "completed" if statuses and all(value == "completed" for value in statuses) else "draft"
            self.put_json(self.manifest_key(source_hash), manifest)
            self._upsert_index_from_manifest(manifest)
            return mode_details

    def set_mode_received(self, source_hash: str, mode: str, received: bool) -> dict[str, Any]:
        """Compatibility wrapper for the former two-state delivery control."""
        target = DELIVERY_STATUS_RECEIVED if bool(received) else DELIVERY_STATUS_SENT
        return self.set_mode_delivery_status(source_hash, mode, target)

    def save_manual_order(self, payload: Mapping[str, Any]) -> dict[str, Any]:
        order_id = str(payload.get("order_id", "")).strip()
        title = str(payload.get("title", "")).strip()
        if not order_id or not title:
            raise CloudStorageError("У ручного заказа должны быть идентификатор и название.")
        now = _now_iso()
        delivery_status = _normalize_delivery_status(
            payload.get("delivery_status", ""),
            received=payload.get("received", False),
        )
        order_date = str(payload.get("order_date", ""))
        status_updated_at = str(payload.get("status_updated_at", "")) or now
        delivery_dates = _normalize_delivery_dates(
            payload.get("delivery_dates", {}),
            order_date=order_date,
            received_at=payload.get("received_at", ""),
            status=delivery_status,
            status_updated_at=status_updated_at,
        )
        normalized = {
            "schema_version": 3,
            "order_id": order_id,
            "title": title,
            "order_date": order_date,
            "note": str(payload.get("note", "")).strip(),
            "quantity": max(0, _safe_int(payload.get("quantity", 0))),
            "delivery_status": delivery_status,
            "delivery_dates": delivery_dates,
            "status_updated_at": status_updated_at,
            "received": delivery_status == DELIVERY_STATUS_RECEIVED,
            "created_at": str(payload.get("created_at", "")) or now,
            "updated_at": str(payload.get("updated_at", "")) or now,
            "received_at": str(payload.get("received_at", "")),
        }
        if normalized["received"]:
            normalized["received_at"] = delivery_dates.get("received_at", "") or _date_only(now)
            normalized["delivery_dates"]["received_at"] = normalized["received_at"]
        else:
            normalized["received_at"] = ""
        self.put_json(self.manual_order_key(order_id), normalized)
        return normalized

    def list_manual_orders(self) -> tuple[dict[str, Any], ...]:
        prefix = self.manual_order_prefix()
        result: list[dict[str, Any]] = []
        continuation: str | None = None
        while True:
            kwargs: dict[str, Any] = {
                "Bucket": self.config.bucket,
                "Prefix": prefix,
                "MaxKeys": 1000,
            }
            if continuation:
                kwargs["ContinuationToken"] = continuation
            try:
                response = self.client.list_objects_v2(**kwargs)
            except (BotoCoreError, ClientError, OSError) as exc:
                raise CloudStorageError(f"Не удалось получить список ручных заказов: {exc}") from exc
            for item in response.get("Contents", []):
                key = str(item.get("Key", ""))
                if not key.endswith(".json"):
                    continue
                payload = self.get_json(key)
                if isinstance(payload, Mapping):
                    result.append(dict(payload))
            if not response.get("IsTruncated"):
                break
            continuation = str(response.get("NextContinuationToken", "")) or None
            if not continuation:
                break
        result.sort(
            key=lambda row: (str(row.get("order_date", "")), str(row.get("updated_at", ""))),
            reverse=True,
        )
        return tuple(result)

    def delete_manual_order(self, order_id: str) -> bool:
        key = self.manual_order_key(order_id)
        if not self.exists(key):
            return False
        try:
            self.client.delete_object(Bucket=self.config.bucket, Key=key)
        except (BotoCoreError, ClientError, OSError) as exc:
            raise CloudStorageError(f"Не удалось удалить ручной заказ: {exc}") from exc
        return True

    def load_draft(self, source_hash: str, mode: str) -> dict[str, Any] | None:
        return self.get_json(self.draft_key(source_hash, mode))

    def restore_workbook(self, source_hash: str, destination_dir: Path) -> tuple[Path, dict[str, Any]]:
        manifest = self.get_json(self.manifest_key(source_hash))
        if not manifest:
            raise CloudStorageError("Сохранённый заказ не найден в облаке.")
        key = str(manifest.get("workbook_key", ""))
        if not key:
            raise CloudStorageError("В сохранённом заказе отсутствует ссылка на исходный Excel.")
        suffix = Path(key).suffix.lower() or ".xlsx"
        destination = destination_dir / f"{source_hash}{suffix}"
        expected_size = _safe_int(manifest.get("workbook_size", 0))
        if not destination.exists() or (expected_size > 0 and destination.stat().st_size != expected_size):
            self.download_file(key, destination)
        return destination, manifest

    def list_manifests(self) -> tuple[dict[str, Any], ...]:
        """Compatibility/migration scan. Normal library reads use orders-index.json."""
        prefix = self._key("workspaces/")
        manifests: list[dict[str, Any]] = []
        continuation: str | None = None
        while True:
            kwargs: dict[str, Any] = {
                "Bucket": self.config.bucket,
                "Prefix": prefix,
                "MaxKeys": 1000,
            }
            if continuation:
                kwargs["ContinuationToken"] = continuation
            try:
                response = self.client.list_objects_v2(**kwargs)
            except (BotoCoreError, ClientError, OSError) as exc:
                raise CloudStorageError(f"Не удалось получить список сохранённых заказов: {exc}") from exc
            for item in response.get("Contents", []):
                key = str(item.get("Key", ""))
                if not key.endswith("/manifest.json"):
                    continue
                manifest = self.get_json(key)
                if manifest:
                    manifests.append(manifest)
            if not response.get("IsTruncated"):
                break
            continuation = str(response.get("NextContinuationToken", "")) or None
            if not continuation:
                break
        manifests.sort(key=lambda item: str(item.get("updated_at", "")), reverse=True)
        return tuple(manifests)

    def rebuild_index(self) -> tuple[dict[str, Any], ...]:
        manifests = self.list_manifests()
        orders = {
            str(manifest.get("source_hash", "")): self._index_entry_from_manifest(manifest)
            for manifest in manifests
            if str(manifest.get("source_hash", "")).strip()
        }
        for source_hash, entry in orders.items():
            self.put_json(self.index_entry_key(source_hash), entry)
        index = {"schema_version": 3, "updated_at": _now_iso(), "orders": orders}
        with _INDEX_LOCK:
            self._save_index(index)
        return tuple(sorted(orders.values(), key=lambda row: str(row.get("updated_at", "")), reverse=True))

    def list_order_index(self, *, refresh: bool = False) -> tuple[dict[str, Any], ...]:
        if refresh:
            return self.rebuild_index()
        index = self._load_index()
        orders = index.get("orders", {})
        if not isinstance(orders, Mapping) or not orders:
            recovered = self._load_index(merge_entries=True)
            orders = recovered.get("orders", {})
            if isinstance(orders, Mapping) and orders:
                self._save_index(recovered)
            else:
                return self.rebuild_index()
        values = [dict(value) for value in orders.values() if isinstance(value, Mapping)]
        values.sort(key=lambda row: str(row.get("updated_at", "")), reverse=True)
        return tuple(values)

    def list_workspace_keys(self, source_hash: str) -> tuple[str, ...]:
        prefix = self._workspace_prefix(source_hash).rstrip("/") + "/"
        result: list[str] = []
        continuation: str | None = None
        while True:
            kwargs: dict[str, Any] = {
                "Bucket": self.config.bucket,
                "Prefix": prefix,
                "MaxKeys": 1000,
            }
            if continuation:
                kwargs["ContinuationToken"] = continuation
            try:
                response = self.client.list_objects_v2(**kwargs)
            except (BotoCoreError, ClientError, OSError) as exc:
                raise CloudStorageError(f"Не удалось проверить файлы удаляемого заказа: {exc}") from exc
            result.extend(str(item.get("Key", "")) for item in response.get("Contents", []) if item.get("Key"))
            if not response.get("IsTruncated"):
                break
            continuation = str(response.get("NextContinuationToken", "")) or None
            if not continuation:
                break
        return tuple(result)

    def delete_keys(self, keys: Sequence[str]) -> tuple[str, ...]:
        failures: list[str] = []
        for start in range(0, len(keys), 1000):
            batch = [str(key) for key in keys[start : start + 1000] if str(key)]
            if not batch:
                continue
            try:
                response = self.client.delete_objects(
                    Bucket=self.config.bucket,
                    Delete={"Objects": [{"Key": key} for key in batch], "Quiet": True},
                )
            except (BotoCoreError, ClientError, OSError) as exc:
                failures.extend(batch)
                continue
            failures.extend(str(row.get("Key", "")) for row in response.get("Errors", []) if row.get("Key"))
        return tuple(dict.fromkeys(failures))

    def delete_workspace(self, source_hash: str) -> tuple[str, ...]:
        """Delete every object below a workspace, verify, then remove its index row.

        The index is deliberately updated last. If Cloudflare reports a partial
        failure, the order remains visible and the exception includes every key
        that is still present, preventing a false successful deletion.
        """
        source_hash = str(source_hash).strip()
        if not source_hash:
            raise CloudStorageError("Не указан идентификатор удаляемого заказа.")
        keys = self.list_workspace_keys(source_hash)
        self.delete_keys(keys)
        remaining = self.list_workspace_keys(source_hash)
        unresolved = tuple(dict.fromkeys(remaining))
        if unresolved:
            preview = ", ".join(unresolved[:8])
            suffix = "…" if len(unresolved) > 8 else ""
            raise CloudStorageError(
                f"Заказ удалён не полностью. Остались объекты: {preview}{suffix}"
            )

        self.delete_keys((self.index_entry_key(source_hash),))
        with _INDEX_LOCK:
            index = self._load_index()
            orders = dict(index.get("orders", {}))
            orders.pop(source_hash, None)
            index["orders"] = orders
            self._save_index(index)
        return keys


@lru_cache(maxsize=1)
def get_cloud_storage() -> S3OrderStorage | None:
    config = load_storage_config()
    if not config.configured:
        return None
    return S3OrderStorage(config)


@lru_cache(maxsize=1)
def get_cloud_storage_status() -> CloudStorageStatus:
    config = load_storage_config()
    if not config.configured:
        return CloudStorageStatus(
            configured=False,
            available=False,
            required=config.required,
            backend_name="S3",
            message="Облачное хранилище заказов не настроено.",
        )
    try:
        storage = get_cloud_storage()
        if storage is None:
            raise CloudStorageError("Хранилище не создано.")
        storage.check()
    except CloudStorageError as exc:
        return CloudStorageStatus(
            configured=True,
            available=False,
            required=config.required,
            backend_name="S3",
            message=str(exc),
        )
    return CloudStorageStatus(
        configured=True,
        available=True,
        required=config.required,
        backend_name="S3",
        message="Облачное хранилище подключено. Исходный Excel хранится постоянно, изменения черновика синхронизируются пакетно.",
    )
