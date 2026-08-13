"""Hand off from the already-deployed standalone updater to the merged app."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
import uuid
from pathlib import Path
from typing import Any

from tdm_host.update.state import health_marker_path, update_root, write_update_state


LEGACY_STAGING_RELATIVE = Path(".update_tmp") / "staging" / "tdm-win"
MANIFEST_NAME = "update-manifest.json"
HELPER_NAME = "update-helper.exe"
TARGET_NAME = "tdm.exe"
MIGRATION_MARKER = "legacy-updater-migrated.json"


class MigrationPreparationError(RuntimeError):
    """Raised when a legacy migration package is missing or inconsistent."""


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _load_manifest(source: Path) -> dict[str, Any]:
    manifest_path = source / MANIFEST_NAME
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise MigrationPreparationError("업데이트 manifest를 읽을 수 없습니다.") from exc
    if manifest.get("schemaVersion") != 1:
        raise MigrationPreparationError("지원하지 않는 업데이트 manifest입니다.")
    version = manifest.get("version")
    if not isinstance(version, str) or not version.strip():
        raise MigrationPreparationError("업데이트 버전이 없습니다.")
    return manifest


def _verify_manifest_file(source: Path, manifest: dict[str, Any], name: str) -> Path:
    entry = manifest.get("files", {}).get(name)
    if not isinstance(entry, dict):
        raise MigrationPreparationError(f"manifest에 {name} 정보가 없습니다.")
    expected = str(entry.get("sha256") or "").lower()
    if len(expected) != 64:
        raise MigrationPreparationError(f"manifest의 {name} SHA256이 올바르지 않습니다.")
    path = source / name
    if not path.is_file():
        raise MigrationPreparationError(f"전환 패키지에 {name} 파일이 없습니다.")
    if _sha256(path) != expected:
        raise MigrationPreparationError(f"{name} 파일 무결성 검증에 실패했습니다.")
    return path


def prepare_legacy_migration(root: Path, source: Path, *, pid: int) -> list[str]:
    manifest = _load_manifest(source)
    target_source = _verify_manifest_file(source, manifest, TARGET_NAME)
    helper_source = _verify_manifest_file(source, manifest, HELPER_NAME)
    version = str(manifest["version"]).strip().lstrip("vV")
    transaction_id = uuid.uuid4().hex
    helper_directory = update_root(root) / "helper" / transaction_id
    helper_directory.mkdir(parents=True, exist_ok=False)
    helper_path = helper_directory / HELPER_NAME
    shutil.copy2(helper_source, helper_path)
    if _sha256(helper_path) != _sha256(helper_source):
        raise MigrationPreparationError("교체 helper 복사 검증에 실패했습니다.")

    health_path = health_marker_path(root, transaction_id)
    health_path.unlink(missing_ok=True)
    write_update_state(
        root,
        mode="legacy-migration",
        transactionId=transaction_id,
        fromVersion=_read_version(root / "version.txt"),
        toVersion=version,
        status="prepared",
        sourcePath=str(source),
        helperPath=str(helper_path),
        targetSourcePath=str(target_source),
    )
    return [
        str(helper_path),
        "--root",
        str(root),
        "--source",
        str(source),
        "--wait-pid",
        str(pid),
        "--transaction-id",
        transaction_id,
        "--to-version",
        version,
    ]


def _read_version(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8").strip().lstrip("vV") or "unknown"
    except OSError:
        return "unknown"


def maybe_start_legacy_migration() -> bool:
    """Launch the helper once when a legacy updater starts the new main.exe."""
    if not getattr(sys, "frozen", False):
        return False
    executable = Path(sys.executable).resolve()
    if executable.name.lower() != "main.exe":
        return False
    root = executable.parent
    if (update_root(root) / MIGRATION_MARKER).exists():
        return False
    source = root / LEGACY_STAGING_RELATIVE
    if not (source / TARGET_NAME).is_file():
        return False
    try:
        command = prepare_legacy_migration(root, source, pid=os.getpid())
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen(
            command,
            cwd=str(root),
            close_fds=True,
            creationflags=creation_flags,
        )
        return True
    except Exception as exc:  # noqa: BLE001 - migration failure must not block the app
        write_update_state(
            root,
            mode="legacy-migration",
            status="preparation_failed",
            error=type(exc).__name__,
            detail=str(exc),
        )
        return False
