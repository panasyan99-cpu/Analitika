from __future__ import annotations

import json
import re
import subprocess
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlparse


class UpdateError(RuntimeError):
    """User-friendly update error."""


@dataclass(frozen=True, order=True)
class AppVersion:
    major: int
    minor: int
    patch: int
    stable_rank: int
    rc_number: int

    @classmethod
    def parse(cls, value: str) -> "AppVersion":
        text = value.strip().lower().lstrip("v")
        match = re.fullmatch(r"(\d+)\.(\d+)\.(\d+)(?:[-_.]?rc(\d+))?", text)
        if not match:
            raise ValueError(f"Некорректный номер версии: {value}")
        major, minor, patch, rc = match.groups()
        # Stable is newer than any RC of the same base version.
        return cls(int(major), int(minor), int(patch), 1 if rc is None else 0, int(rc or 0))


def config_path(base_dir: Path) -> Path:
    return base_dir / "update_config.json"


def load_update_config(base_dir: Path) -> dict:
    path = config_path(base_dir)
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise UpdateError(f"Не удалось прочитать настройки обновлений: {exc}") from exc


def _github_json(url: str, timeout: int = 12):
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "Analitika-Updater",
            "X-GitHub-Api-Version": "2022-11-28",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            raise UpdateError("Релизы пока не опубликованы в GitHub Releases.") from exc
        raise UpdateError(f"GitHub вернул ошибку HTTP {exc.code}.") from exc
    except urllib.error.URLError as exc:
        raise UpdateError("Не удалось подключиться к GitHub. Проверьте интернет.") from exc
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        raise UpdateError("GitHub вернул некорректный ответ.") from exc


def _pick_installer_asset(release: dict) -> dict | None:
    return next(
        (
            item
            for item in release.get("assets", [])
            if str(item.get("name", "")).lower().endswith(".exe")
            and "setup" in str(item.get("name", "")).lower()
        ),
        None,
    )


def check_for_update(
    current_version: str,
    base_dir: Path,
    channel: str = "stable",
) -> dict | None:
    """Return the newest suitable GitHub release or None.

    stable: only non-prerelease releases.
    rc: stable and prerelease releases.
    """
    cfg = load_update_config(base_dir)
    repo = str(cfg.get("github_repo", "")).strip()
    if not repo or repo.startswith("OWNER/") or "/" not in repo:
        return None

    channel = channel.lower().strip()
    if channel not in {"stable", "rc"}:
        channel = "stable"

    releases = _github_json(f"https://api.github.com/repos/{repo}/releases?per_page=30")
    if not isinstance(releases, list):
        raise UpdateError("GitHub вернул неожиданный список релизов.")

    try:
        installed = AppVersion.parse(current_version)
    except ValueError as exc:
        raise UpdateError(str(exc)) from exc

    candidates: list[tuple[AppVersion, dict]] = []
    for release in releases:
        if release.get("draft"):
            continue
        if channel == "stable" and release.get("prerelease"):
            continue
        tag = str(release.get("tag_name") or release.get("name") or "").strip()
        try:
            parsed = AppVersion.parse(tag)
        except ValueError:
            continue
        if parsed > installed:
            candidates.append((parsed, release))

    if not candidates:
        return None

    _, newest = max(candidates, key=lambda item: item[0])
    tag = str(newest.get("tag_name") or newest.get("name") or "").strip().lstrip("vV")
    asset = _pick_installer_asset(newest)
    return {
        "version": tag,
        "notes": str(newest.get("body") or "").strip(),
        "download_url": asset.get("browser_download_url") if asset else None,
        "asset_name": asset.get("name") if asset else None,
        "asset_size": int(asset.get("size") or 0) if asset else 0,
        "page_url": newest.get("html_url"),
        "prerelease": bool(newest.get("prerelease")),
        "channel": "rc" if newest.get("prerelease") else "stable",
    }


def download_installer(
    url: str,
    updates_dir: Path,
    expected_size: int = 0,
    filename: str | None = None,
) -> Path:
    updates_dir.mkdir(parents=True, exist_ok=True)
    safe_name = filename or Path(urlparse(url).path).name or "Analitika_Update_Setup.exe"
    target = updates_dir / safe_name
    partial = target.with_suffix(target.suffix + ".part")

    request = urllib.request.Request(url, headers={"User-Agent": "Analitika-Updater"})
    try:
        with urllib.request.urlopen(request, timeout=120) as response, partial.open("wb") as output:
            header_size = int(response.headers.get("Content-Length") or 0)
            while True:
                chunk = response.read(1024 * 1024)
                if not chunk:
                    break
                output.write(chunk)
    except (urllib.error.URLError, OSError) as exc:
        partial.unlink(missing_ok=True)
        raise UpdateError(f"Не удалось скачать обновление: {exc}") from exc

    actual_size = partial.stat().st_size
    required_size = expected_size or header_size
    if required_size and actual_size != required_size:
        partial.unlink(missing_ok=True)
        raise UpdateError(
            f"Файл обновления скачан не полностью: {actual_size} из {required_size} байт."
        )
    if actual_size < 100_000:
        partial.unlink(missing_ok=True)
        raise UpdateError("Скачанный установщик имеет подозрительно маленький размер.")

    target.unlink(missing_ok=True)
    partial.replace(target)
    return target


def launch_installer(installer: Path, silent: bool = True) -> None:
    if not installer.exists():
        raise UpdateError("Файл установщика не найден.")
    if not sys.platform.startswith("win"):
        raise UpdateError("Автоматическая установка обновлений поддерживается только в Windows.")

    args = [str(installer)]
    if silent:
        args.extend([
            "/SILENT",
            "/SUPPRESSMSGBOXES",
            "/CLOSEAPPLICATIONS",
            "/NORESTART",
            "/SP-",
        ])
    try:
        subprocess.Popen(args, close_fds=True)
    except OSError as exc:
        raise UpdateError(f"Не удалось запустить установщик: {exc}") from exc
