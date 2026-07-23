from __future__ import annotations

import io
import json
import os
from dataclasses import dataclass
from datetime import datetime, timezone
from functools import lru_cache
from pathlib import Path
from typing import Any, Mapping

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

    Source workbooks are immutable objects addressed by SHA-256. Draft JSON is
    updated in place. boto3 automatically switches to multipart upload for large
    workbooks, so reports much larger than the Streamlit instance disk survive
    app restarts and deployments.
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

    def save_workbook(self, source_hash: str, source_name: str, payload: bytes) -> dict[str, Any]:
        workbook_key = self.workbook_key(source_hash, source_name)
        if not self.exists(workbook_key):
            self.put_bytes(
                workbook_key,
                payload,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        manifest = self.get_json(self.manifest_key(source_hash)) or {}
        manifest.update(
            {
                "schema_version": 1,
                "source_hash": source_hash,
                "source_name": source_name,
                "workbook_key": workbook_key,
                "workbook_size": len(payload),
                "updated_at": _now_iso(),
            }
        )
        manifest.setdefault("drafts", {})
        self.put_json(self.manifest_key(source_hash), manifest)
        return manifest

    def save_draft(self, payload: Mapping[str, Any]) -> dict[str, Any]:
        source_hash = str(payload.get("source_hash", "")).strip()
        source_name = str(payload.get("source_name", "")).strip()
        mode = str(payload.get("mode", "")).strip()
        if not source_hash or mode not in MODE_FILE_NAMES:
            raise CloudStorageError("Черновик не содержит идентификатор отчёта или тип заказа.")
        draft_key = self.draft_key(source_hash, mode)
        self.put_json(draft_key, payload)
        orders = payload.get("orders", {})
        if not isinstance(orders, Mapping):
            orders = {}
        selected_positions = sum(1 for value in orders.values() if int(value or 0) > 0)
        total_quantity = sum(max(0, int(value or 0)) for value in orders.values())
        manifest = self.get_json(self.manifest_key(source_hash)) or {
            "schema_version": 1,
            "source_hash": source_hash,
            "source_name": source_name,
            "workbook_key": self.workbook_key(source_hash, source_name),
            "workbook_size": 0,
            "drafts": {},
        }
        drafts = manifest.setdefault("drafts", {})
        drafts[mode] = {
            "key": draft_key,
            "updated_at": str(payload.get("updated_at", "")) or _now_iso(),
            "selected_positions": selected_positions,
            "total_quantity": total_quantity,
            "stage": str(payload.get("stage", "order")),
        }
        manifest["source_name"] = source_name or str(manifest.get("source_name", ""))
        manifest["updated_at"] = str(payload.get("updated_at", "")) or _now_iso()
        self.put_json(self.manifest_key(source_hash), manifest)
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
        expected_size = int(manifest.get("workbook_size", 0) or 0)
        if not destination.exists() or (expected_size > 0 and destination.stat().st_size != expected_size):
            self.download_file(key, destination)
        return destination, manifest

    def list_manifests(self) -> tuple[dict[str, Any], ...]:
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
