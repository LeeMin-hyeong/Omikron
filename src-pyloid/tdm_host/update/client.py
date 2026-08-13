"""Update checker embedded in the merged TDM application."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
import threading
import uuid
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any
from urllib.request import Request, urlopen

from tdm_host.update.state import post_update_transaction_id, update_root, write_update_state


GITHUB_REPOSITORY = "LeeMin-hyeong/TestDataManagement"
RELEASE_API = f"https://api.github.com/repos/{GITHUB_REPOSITORY}/releases/latest"
ARCHIVE_NAME = "tdm-win.zip"
CHECKSUM_NAME = f"{ARCHIVE_NAME}.sha256"
MANIFEST_NAME = "update-manifest.json"
HELPER_NAME = "update-helper.exe"
HTTP_TIMEOUT_SECONDS = 30
_update_gate_handle: int | None = None


class UpdatePreparationError(RuntimeError):
    """Raised when a downloaded release cannot be safely prepared."""


def _claim_update_gate() -> bool:
    """Let only the first application process perform startup update work."""
    global _update_gate_handle
    if os.name != "nt":
        return True
    import ctypes

    kernel32 = ctypes.windll.kernel32
    kernel32.CreateMutexW.restype = ctypes.c_void_p
    handle = kernel32.CreateMutexW(None, False, "Local\\TDM.Application.UpdateGate")
    if not handle:
        return False
    if kernel32.GetLastError() == 183:
        kernel32.CloseHandle(handle)
        return False
    _update_gate_handle = int(handle)
    return True


def _request(url: str):
    return urlopen(
        Request(
            url,
            headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": "tdm-updater",
            },
        ),
        timeout=HTTP_TIMEOUT_SECONDS,
    )


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _version_parts(value: str) -> tuple[int, ...]:
    parts: list[int] = []
    for part in value.strip().lstrip("vV").split("."):
        digits = "".join(character for character in part if character.isdigit())
        parts.append(int(digits) if digits else 0)
    while parts and parts[-1] == 0:
        parts.pop()
    return tuple(parts)


def _read_local_version(root: Path) -> str:
    try:
        return (root / "version.txt").read_text(encoding="utf-8-sig").strip().lstrip("vV")
    except OSError:
        return "0.0.0"


def _latest_release() -> tuple[str, str, str]:
    with _request(RELEASE_API) as response:
        release = json.loads(response.read().decode("utf-8"))
    version = str(release.get("tag_name") or release.get("name") or "").lstrip("vV")
    assets = {
        str(asset.get("name")): str(asset.get("browser_download_url"))
        for asset in release.get("assets", [])
        if isinstance(asset, dict)
    }
    archive_url = assets.get(ARCHIVE_NAME)
    checksum_url = assets.get(CHECKSUM_NAME)
    if not version or not archive_url or not checksum_url:
        raise UpdatePreparationError("릴리스에 업데이트 파일 또는 체크섬이 없습니다.")
    return version, archive_url, checksum_url


def _download(url: str, target: Path) -> None:
    partial = target.with_suffix(f"{target.suffix}.part")
    partial.unlink(missing_ok=True)
    try:
        with _request(url) as response, partial.open("wb") as stream:
            shutil.copyfileobj(response, stream, length=1024 * 1024)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(partial, target)
    finally:
        partial.unlink(missing_ok=True)


def _safe_extract(archive: Path, destination: Path) -> None:
    destination_resolved = destination.resolve()
    with zipfile.ZipFile(archive) as zipped:
        for member in zipped.infolist():
            candidate = (destination / member.filename).resolve()
            try:
                candidate.relative_to(destination_resolved)
            except ValueError as exc:
                raise UpdatePreparationError("업데이트 압축 파일에 안전하지 않은 경로가 있습니다.") from exc
        zipped.extractall(destination)


def _verify_manifest(source: Path, expected_version: str) -> dict[str, Any]:
    try:
        manifest = json.loads((source / MANIFEST_NAME).read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise UpdatePreparationError("업데이트 manifest를 읽을 수 없습니다.") from exc
    if manifest.get("schemaVersion") != 1:
        raise UpdatePreparationError("지원하지 않는 업데이트 manifest입니다.")
    manifest_version = str(manifest.get("version") or "").lstrip("vV")
    if manifest_version != expected_version.lstrip("vV"):
        raise UpdatePreparationError("릴리스 버전과 업데이트 파일 버전이 일치하지 않습니다.")
    files = manifest.get("files")
    if not isinstance(files, dict):
        raise UpdatePreparationError("업데이트 manifest에 파일 목록이 없습니다.")
    source_root = source.resolve()
    for relative_name, entry in files.items():
        if not isinstance(relative_name, str) or not isinstance(entry, dict):
            raise UpdatePreparationError("업데이트 manifest 파일 정보가 올바르지 않습니다.")
        path = (source / Path(relative_name)).resolve()
        try:
            path.relative_to(source_root)
        except ValueError as exc:
            raise UpdatePreparationError("업데이트 manifest에 안전하지 않은 경로가 있습니다.") from exc
        expected_hash = str(entry.get("sha256") or "").lower()
        if not path.is_file() or len(expected_hash) != 64 or _sha256(path) != expected_hash:
            raise UpdatePreparationError(f"업데이트 파일 무결성 검증에 실패했습니다: {relative_name}")
    for required in ("main.exe", "tdm.exe", HELPER_NAME, "version.txt"):
        if required not in files:
            raise UpdatePreparationError(f"업데이트 manifest에 {required} 정보가 없습니다.")
    if not (source / "_internal").is_dir():
        raise UpdatePreparationError("업데이트 파일에 _internal 디렉터리가 없습니다.")
    return manifest


def prepare_update(root: Path, version: str, archive_url: str, checksum_url: str) -> list[str]:
    transaction_id = uuid.uuid4().hex
    transaction_root = update_root(root) / "transactions" / transaction_id
    download_directory = transaction_root / "downloads"
    staging_directory = transaction_root / "staging"
    download_directory.mkdir(parents=True, exist_ok=False)
    archive = download_directory / ARCHIVE_NAME
    _download(archive_url, archive)
    with _request(checksum_url) as response:
        expected_hash = response.read().decode("ascii", errors="strict").strip().split()[0].lower()
    if len(expected_hash) != 64 or _sha256(archive) != expected_hash:
        raise UpdatePreparationError("다운로드한 업데이트 파일의 SHA256이 일치하지 않습니다.")
    staging_directory.mkdir(parents=True)
    _safe_extract(archive, staging_directory)
    source = staging_directory / "tdm-win"
    _verify_manifest(source, version)

    helper_directory = update_root(root) / "helper" / transaction_id
    helper_directory.mkdir(parents=True, exist_ok=False)
    helper = helper_directory / HELPER_NAME
    shutil.copy2(source / HELPER_NAME, helper)
    if _sha256(helper) != _sha256(source / HELPER_NAME):
        raise UpdatePreparationError("업데이트 helper 복사 검증에 실패했습니다.")
    write_update_state(
        root,
        mode="update",
        transactionId=transaction_id,
        fromVersion=_read_local_version(root),
        toVersion=version,
        status="prepared",
        sourcePath=str(source),
        helperPath=str(helper),
    )
    return [
        str(helper),
        "--mode",
        "update",
        "--root",
        str(root),
        "--source",
        str(source),
        "--wait-pid",
        str(os.getpid()),
        "--transaction-id",
        transaction_id,
        "--to-version",
        version,
    ]


class _UpdateWindow:
    def __init__(self, root: Path):
        import tkinter as tk
        from tkinter import ttk

        self.application_root = root
        self.should_exit = False
        self.window = tk.Tk()
        self.window.title("tdm 업데이트")
        self.window.resizable(False, False)
        width, height = 380, 150
        x = (self.window.winfo_screenwidth() - width) // 2
        y = (self.window.winfo_screenheight() - height) // 2
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        self.status = tk.StringVar(value="업데이트 확인 중…")
        tk.Label(self.window, text="tdm", font=("Segoe UI", 18, "bold")).pack(pady=(22, 8))
        tk.Label(self.window, textvariable=self.status, font=("Segoe UI", 10)).pack()
        self.progress = ttk.Progressbar(self.window, mode="indeterminate", length=280)
        self.progress.pack(pady=14)
        self.progress.start(12)
        self.window.protocol("WM_DELETE_WINDOW", lambda: None)

    def set_status(self, value: str) -> None:
        self.window.after(0, self.status.set, value)

    def close(self) -> None:
        self.window.after(0, self.window.destroy)

    def run(self, action: Callable[[], bool]) -> bool:
        def worker() -> None:
            try:
                self.should_exit = action()
            finally:
                self.close()

        self.window.after(50, lambda: threading.Thread(target=worker, daemon=True).start())
        self.window.mainloop()
        return self.should_exit


def maybe_install_available_update() -> bool:
    """Check, stage and hand off an update; return true when this process must exit."""
    if not getattr(sys, "frozen", False) or not _claim_update_gate():
        return False
    if post_update_transaction_id() is not None:
        return False
    root = Path(sys.executable).resolve().parent
    window = _UpdateWindow(root)

    def action() -> bool:
        try:
            current = _read_local_version(root)
            latest, archive_url, checksum_url = _latest_release()
            if _version_parts(current) >= _version_parts(latest):
                window.set_status("최신 버전입니다.")
                return False
            window.set_status(f"업데이트 다운로드 중… ({current} → {latest})")
            command = prepare_update(root, latest, archive_url, checksum_url)
            window.set_status("업데이트 적용 준비 중…")
            subprocess.Popen(
                command,
                cwd=str(root),
                close_fds=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            return True
        except Exception as exc:  # noqa: BLE001 - update failure must not block app startup
            write_update_state(
                root,
                mode="update",
                status="preparation_failed",
                error=type(exc).__name__,
                detail=str(exc),
            )
            return False

    return window.run(action)
