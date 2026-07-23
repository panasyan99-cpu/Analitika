from __future__ import annotations

import io
import json
import os
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


def _now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _safe_int(value: object) -> int:
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


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

    def index_key(self) -> str:
        return self._key("orders-index.json")

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
            normalized_drafts[str(mode)] = {
                "key": str(details.get("key", "")),
                "created_at": str(details.get("created_at", "")),
                "updated_at": str(details.get("updated_at", "")),
                "selected_positions": max(0, _safe_int(details.get("selected_positions", 0))),
                "total_quantity": max(0, _safe_int(details.get("total_quantity", 0))),
                "limited_positions": max(0, _safe_int(details.get("limited_positions", 0))),
                "stage": str(details.get("stage", "order")),
                "status": "completed" if str(details.get("status", "draft")) == "completed" else "draft",
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

    def _load_index(self) -> dict[str, Any]:
        index = self.get_json(self.index_key())
        if not index:
            return {"schema_version": 1, "updated_at": _now_iso(), "orders": {}}
        orders = index.get("orders", {})
        if not isinstance(orders, dict):
            orders = {}
        return {
            "schema_version": 1,
            "updated_at": str(index.get("updated_at", "")) or _now_iso(),
            "orders": orders,
        }

    def _save_index(self, index: Mapping[str, Any]) -> None:
        payload = dict(index)
        payload["schema_version"] = 1
        payload["updated_at"] = _now_iso()
        self.put_json(self.index_key(), payload)

    def _upsert_index_from_manifest(self, manifest: Mapping[str, Any]) -> None:
        source_hash = str(manifest.get("source_hash", "")).strip()
        if not source_hash:
            return
        index = self._load_index()
        orders = dict(index.get("orders", {}))
        orders[source_hash] = self._index_entry_from_manifest(manifest)
        index["orders"] = orders
        self._save_index(index)

    def save_workbook(self, source_hash: str, source_name: str, payload: bytes) -> dict[str, Any]:
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
                "schema_version": 2,
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
        selected_positions = sum(1 for value in orders.values() if _safe_int(value) > 0)
        total_quantity = sum(max(0, _safe_int(value)) for value in orders.values())
        limited_positions = sum(1 for value in limited_orders.values() if bool(value))
        manifest = self.get_json(self.manifest_key(source_hash)) or {
            "schema_version": 2,
            "source_hash": source_hash,
            "source_name": source_name,
            "workbook_key": self.workbook_key(source_hash, source_name),
            "workbook_size": 0,
            "created_at": created_at,
            "drafts": {},
        }
        drafts = manifest.setdefault("drafts", {})
        previous = drafts.get(mode, {}) if isinstance(drafts.get(mode), Mapping) else {}
        status = "completed" if str(payload.get("status", "draft")) == "completed" else "draft"
        drafts[mode] = {
            "key": draft_key,
            "created_at": str(previous.get("created_at", "")) or created_at,
            "updated_at": now,
            "selected_positions": selected_positions,
            "total_quantity": total_quantity,
            "limited_positions": limited_positions,
            "stage": str(payload.get("stage", "order")),
            "status": status,
        }
        statuses = [str(row.get("status", "draft")) for row in drafts.values() if isinstance(row, Mapping)]
        manifest["schema_version"] = 2
        manifest["source_name"] = source_name or str(manifest.get("source_name", ""))
        manifest["created_at"] = str(manifest.get("created_at", "")) or created_at
        manifest["updated_at"] = now
        manifest["status"] = "completed" if statuses and all(value == "completed" for value in statuses) else "draft"
        self.put_json(self.manifest_key(source_hash), manifest)
        self._upsert_index_from_manifest(manifest)
        return manifest

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
        index = {"schema_version": 1, "updated_at": _now_iso(), "orders": orders}
        self._save_index(index)
        return tuple(sorted(orders.values(), key=lambda row: str(row.get("updated_at", "")), reverse=True))

    def list_order_index(self, *, refresh: bool = False) -> tuple[dict[str, Any], ...]:
        if refresh:
            return self.rebuild_index()
        index = self.get_json(self.index_key())
        if not index:
            return self.rebuild_index()
        orders = index.get("orders", {})
        if not isinstance(orders, Mapping):
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
        message="Облачное хранилище подключено. Исходный Excel и каждое изменение сохраняются вне Streamlit.",
    )
