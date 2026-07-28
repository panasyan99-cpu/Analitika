from __future__ import annotations

from datetime import datetime
from io import BytesIO
import mimetypes
from pathlib import Path
import time
from typing import Any, Iterable, Protocol

from PIL import Image, ImageChops, ImageOps
import requests


class WarehouseClientError(RuntimeError):
    pass


class ConfigProtocol(Protocol):
    base_url: str
    token: str
    souvenirs_table_id: int
    components_table_id: int
    operations_table_id: int
    supplies_table_id: int
    supply_lines_table_id: int


def as_int(value: Any, default: int = 0) -> int:
    if value in (None, ""):
        return default
    if isinstance(value, bool):
        return int(value)
    try:
        return int(float(str(value).replace(" ", "").replace(",", ".")))
    except (TypeError, ValueError):
        return default


def select_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, dict):
        return str(value.get("value") or value.get("name") or "")
    if isinstance(value, list):
        return "; ".join(filter(None, (select_text(item) for item in value)))
    return str(value)


def link_ids(value: Any) -> list[int]:
    result: list[int] = []
    values = value if isinstance(value, list) else [value] if value else []
    for item in values:
        raw = item.get("id") if isinstance(item, dict) else item
        try:
            result.append(int(raw))
        except (TypeError, ValueError):
            continue
    return result


class WarehouseClient:
    """Small Baserow Database API client used by the online warehouse module."""

    def __init__(self, config: ConfigProtocol) -> None:
        self.config = config
        self.base_url = str(config.base_url).rstrip("/")
        self.token = str(config.token).strip()
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Authorization": f"Token {self.token}",
                "Accept": "application/json, text/plain, */*",
                "User-Agent": "Princess-Analitika-Warehouse-Web/2.0",
            }
        )
        self._fields: dict[int, list[dict[str, Any]]] = {}

    @staticmethod
    def batch_id(prefix: str) -> str:
        return f"{prefix}-{datetime.now():%Y%m%d-%H%M%S-%f}"

    def _request(
        self,
        method: str,
        path: str,
        *,
        payload: Any = None,
        params: dict[str, Any] | None = None,
        files: Any = None,
        timeout: int = 90,
        retries: int = 2,
    ) -> Any:
        url = f"{self.base_url}{path}"
        for attempt in range(retries + 1):
            try:
                response = self.session.request(
                    method,
                    url,
                    json=payload if files is None else None,
                    params=params,
                    files=files,
                    timeout=timeout,
                )
            except requests.RequestException as exc:
                if attempt < retries:
                    time.sleep(0.7 * (attempt + 1))
                    continue
                raise WarehouseClientError(f"Не удалось подключиться к Baserow: {exc}") from exc

            if 200 <= response.status_code < 300:
                return response.json() if response.content else None
            if response.status_code in {429, 500, 502, 503, 504} and attempt < retries:
                time.sleep(0.8 * (attempt + 1))
                continue
            if response.status_code in {401, 403}:
                raise WarehouseClientError(
                    "Baserow отклонил токен или у токена нет права на эту операцию."
                )
            raise WarehouseClientError(
                f"Baserow HTTP {response.status_code}: {response.text[:900]}"
            )
        raise WarehouseClientError("Неожиданная ошибка Baserow.")

    def list_rows(self, table_id: int, *, query: str = "") -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        page = 1
        while True:
            params: dict[str, Any] = {
                "user_field_names": "true",
                "size": 200,
                "page": page,
            }
            if query:
                for part in query.lstrip("?&").split("&"):
                    if "=" in part:
                        key, value = part.split("=", 1)
                        params[key] = value
            result = self._request(
                "GET",
                f"/api/database/rows/table/{int(table_id)}/",
                params=params,
            )
            rows.extend(result.get("results", []))
            if not result.get("next"):
                break
            page += 1
        return rows

    def fields(self, table_id: int, *, refresh: bool = False) -> list[dict[str, Any]]:
        table_id = int(table_id)
        if not refresh and table_id in self._fields:
            return self._fields[table_id]
        result = self._request("GET", f"/api/database/fields/table/{table_id}/")
        fields = list(result or [])
        self._fields[table_id] = fields
        return fields

    def field_map(self, table_id: int) -> dict[str, dict[str, Any]]:
        return {
            str(field.get("name") or ""): field
            for field in self.fields(table_id)
            if str(field.get("name") or "")
        }

    @staticmethod
    def _select_parts(value: Any) -> list[str]:
        if isinstance(value, list):
            return [str(item).strip() for item in value if str(item).strip()]
        return [part.strip() for part in str(value or "").split(";") if part.strip()]

    def normalize_payload(self, table_id: int, payload: dict[str, Any]) -> dict[str, Any]:
        """Drop unknown fields and adapt values to the current Baserow schema."""
        fields = self.field_map(table_id)
        result: dict[str, Any] = {}
        for name, value in payload.items():
            field = fields.get(name)
            if field is None:
                continue
            field_type = str(field.get("type") or "")
            if field_type == "multiple_select":
                result[name] = self._select_parts(value)
            elif field_type == "single_select":
                parts = self._select_parts(value)
                result[name] = parts[0] if parts else None
            elif field_type == "link_row":
                result[name] = link_ids(value)
            elif field_type == "boolean":
                result[name] = bool(value)
            elif field_type in {"number", "rating"}:
                result[name] = value if value not in (None, "") else None
            elif field_type == "file":
                result[name] = value if isinstance(value, list) else []
            else:
                result[name] = value
        return result

    def create_row(self, table_id: int, payload: dict[str, Any]) -> dict[str, Any]:
        clean = self.normalize_payload(table_id, payload)
        return self._request(
            "POST",
            f"/api/database/rows/table/{int(table_id)}/",
            params={"user_field_names": "true"},
            payload=clean,
        )

    def batch_create(self, table_id: int, items: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
        source = [self.normalize_payload(table_id, item) for item in items]
        created: list[dict[str, Any]] = []
        for start in range(0, len(source), 100):
            result = self._request(
                "POST",
                f"/api/database/rows/table/{int(table_id)}/batch/",
                params={"user_field_names": "true"},
                payload={"items": source[start : start + 100]},
            )
            if isinstance(result, dict):
                created.extend(result.get("items", []))
            elif isinstance(result, list):
                created.extend(result)
        return created

    def batch_update(self, table_id: int, items: Iterable[dict[str, Any]]) -> None:
        source: list[dict[str, Any]] = []
        for item in items:
            row_id = int(item["id"])
            clean = self.normalize_payload(table_id, {k: v for k, v in item.items() if k != "id"})
            source.append({"id": row_id, **clean})
        for start in range(0, len(source), 100):
            self._request(
                "PATCH",
                f"/api/database/rows/table/{int(table_id)}/batch/",
                params={"user_field_names": "true"},
                payload={"items": source[start : start + 100]},
            )

    def delete_row(self, table_id: int, row_id: int) -> None:
        self._request(
            "DELETE",
            f"/api/database/rows/table/{int(table_id)}/{int(row_id)}/",
        )

    @staticmethod
    def _prepare_image(path: Path) -> tuple[bytes, str, str]:
        try:
            with Image.open(path) as source:
                image = ImageOps.exif_transpose(source).convert("RGBA")
                background = Image.new("RGBA", image.size, "white")
                background.alpha_composite(image)
                rgb = background.convert("RGB")
                white = Image.new("RGB", rgb.size, "white")
                diff = ImageChops.difference(rgb, white).convert("L")
                bbox = diff.point(lambda value: 255 if value > 12 else 0).getbbox()
                if bbox:
                    left, top, right, bottom = bbox
                    margin = max(8, int(max(right - left, bottom - top) * 0.08))
                    rgb = rgb.crop(
                        (
                            max(0, left - margin),
                            max(0, top - margin),
                            min(rgb.width, right + margin),
                            min(rgb.height, bottom + margin),
                        )
                    )
                target = 980
                scale = min(target / max(rgb.width, 1), target / max(rgb.height, 1), 1.0)
                size = (max(1, int(rgb.width * scale)), max(1, int(rgb.height * scale)))
                rgb = rgb.resize(size, getattr(Image, "Resampling", Image).LANCZOS)
                canvas = Image.new("RGB", (1100, 1100), "white")
                canvas.paste(rgb, ((1100 - size[0]) // 2, (1100 - size[1]) // 2))
                output = BytesIO()
                canvas.save(output, format="JPEG", quality=92, optimize=True)
                return output.getvalue(), f"{path.stem}_large.jpg", "image/jpeg"
        except Exception:
            mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
            return path.read_bytes(), path.name, mime

    def upload_file(self, path: Path) -> dict[str, Any]:
        data, name, mime = self._prepare_image(path)
        result = self._request(
            "POST",
            "/api/user-files/upload-file/",
            files={"file": (name, data, mime)},
            timeout=240,
        )
        if not isinstance(result, dict):
            raise WarehouseClientError("Baserow не вернул описание загруженного файла.")
        return result

    def operation_exists(self, batch_id: str) -> bool:
        if not batch_id:
            return False
        for row in self.list_rows(self.config.operations_table_id):
            if str(row.get("Batch ID") or "").strip() == batch_id:
                return True
        return False

    def create_operations(self, items: list[dict[str, Any]], *, batch_id: str) -> list[dict[str, Any]]:
        if not items:
            return []
        if self.operation_exists(batch_id):
            raise WarehouseClientError(
                f"Документ {batch_id} уже существует. Повторное проведение заблокировано."
            )
        return self.batch_create(self.config.operations_table_id, items)
