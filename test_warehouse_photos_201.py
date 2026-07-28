from __future__ import annotations

from pathlib import Path

from PIL import Image

from src.warehouse_management.ui import _local_thumbnail_data_uri, _photo_url


def test_photo_url_prefers_small_baserow_thumbnail() -> None:
    value = [{"url": "/media/full.jpg", "thumbnails": {"small": {"url": "/media/small.jpg"}}}]
    assert _photo_url(value, "https://storage.example") == "https://storage.example/media/small.jpg"


def test_photo_url_supports_original_fallback() -> None:
    value = [{"url": "/media/user_files/photo.jpg"}]
    assert _photo_url(value, "https://storage.example") == "https://storage.example/media/user_files/photo.jpg"


def test_local_thumbnail_is_compact_data_uri(tmp_path: Path) -> None:
    path = tmp_path / "photo.png"
    Image.new("RGB", (800, 600), "white").save(path)
    result = _local_thumbnail_data_uri(str(path), path.stat().st_mtime_ns)
    assert result.startswith("data:image/jpeg;base64,")
    assert len(result) < 100_000
