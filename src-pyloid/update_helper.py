"""Independent Windows update switcher used after the main process exits.

This module intentionally depends only on the Python standard library so its
one-file executable does not load anything from the installation's _internal
directory while that directory is being updated.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


TARGET_NAME = "tdm.exe"
HELPER_NAME = "update-helper.exe"
MANIFEST_NAME = "update-manifest.json"
UPDATE_PAYLOAD_NAMES = (
    "main.exe",
    TARGET_NAME,
    HELPER_NAME,
    "version.txt",
    "LICENSE",
    "_internal",
)
WAIT_TIMEOUT_SECONDS = 60.0
RETRY_INTERVAL_SECONDS = 0.5
TRANSACTION_ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]{1,64}$")


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _atomic_write_json(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    try:
        with temporary.open("w", encoding="utf-8", newline="\n") as stream:
            json.dump(value, stream, ensure_ascii=False, indent=2, sort_keys=True)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def _write_state(update_root: Path, **values: Any) -> None:
    _atomic_write_json(
        update_root / "update-state.json",
        {
            "schemaVersion": 1,
            "updatedAt": datetime.now(timezone.utc).isoformat(),
            **values,
        },
    )


def _append_log(update_root: Path, message: str) -> None:
    log_path = update_root / "logs" / "update.log"
    log_path.parent.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now(timezone.utc).isoformat()
    with log_path.open("a", encoding="utf-8", newline="\n") as stream:
        stream.write(f"{timestamp} {message}\n")


def _prune_directories(parent: Path, *, keep: int = 2) -> None:
    if not parent.is_dir():
        return
    directories = sorted(
        (path for path in parent.iterdir() if path.is_dir()),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    for directory in directories[keep:]:
        shutil.rmtree(directory, ignore_errors=True)


def _process_exists(pid: int) -> bool:
    if pid <= 0:
        return False
    if os.name == "nt":
        import ctypes

        process = ctypes.windll.kernel32.OpenProcess(0x00100000, False, pid)
        if not process:
            return False
        try:
            return ctypes.windll.kernel32.WaitForSingleObject(process, 0) == 0x102
        finally:
            ctypes.windll.kernel32.CloseHandle(process)
    try:
        os.kill(pid, 0)
    except OSError:
        return False
    return True


def _wait_for_process_exit(pid: int, timeout: float = WAIT_TIMEOUT_SECONDS) -> None:
    deadline = time.monotonic() + timeout
    while _process_exists(pid):
        if time.monotonic() >= deadline:
            raise TimeoutError("메인 프로그램 종료를 기다리는 시간이 초과되었습니다.")
        time.sleep(RETRY_INTERVAL_SECONDS)


def _load_and_verify_source(source: Path) -> tuple[Path, str]:
    manifest = json.loads((source / MANIFEST_NAME).read_text(encoding="utf-8-sig"))
    if manifest.get("schemaVersion") != 1:
        raise RuntimeError("지원하지 않는 업데이트 manifest입니다.")
    version = str(manifest.get("version") or "").strip().lstrip("vV")
    if not version:
        raise RuntimeError("업데이트 버전이 없습니다.")
    entry = manifest.get("files", {}).get(TARGET_NAME)
    if not isinstance(entry, dict):
        raise RuntimeError("manifest에 tdm.exe 정보가 없습니다.")
    expected = str(entry.get("sha256") or "").lower()
    target_source = source / TARGET_NAME
    if not target_source.is_file() or len(expected) != 64:
        raise RuntimeError("새 tdm.exe 파일 정보가 올바르지 않습니다.")
    if _sha256(target_source) != expected:
        raise RuntimeError("새 tdm.exe 파일 무결성 검증에 실패했습니다.")
    return target_source, version


def _load_and_verify_update_source(source: Path) -> str:
    manifest = json.loads((source / MANIFEST_NAME).read_text(encoding="utf-8-sig"))
    if manifest.get("schemaVersion") != 1:
        raise RuntimeError("지원하지 않는 업데이트 manifest입니다.")
    version = str(manifest.get("version") or "").strip().lstrip("vV")
    files = manifest.get("files")
    if not version or not isinstance(files, dict):
        raise RuntimeError("업데이트 manifest 정보가 올바르지 않습니다.")
    source_root = source.resolve()
    for relative_name, entry in files.items():
        if not isinstance(relative_name, str) or not isinstance(entry, dict):
            raise RuntimeError("업데이트 manifest 파일 정보가 올바르지 않습니다.")
        path = (source / Path(relative_name)).resolve()
        try:
            path.relative_to(source_root)
        except ValueError as exc:
            raise RuntimeError("업데이트 manifest에 안전하지 않은 경로가 있습니다.") from exc
        expected = str(entry.get("sha256") or "").lower()
        if not path.is_file() or len(expected) != 64 or _sha256(path) != expected:
            raise RuntimeError(f"업데이트 파일 무결성 검증에 실패했습니다: {relative_name}")
    for required in UPDATE_PAYLOAD_NAMES:
        path = source / required
        if required == "_internal":
            if not path.is_dir():
                raise RuntimeError("업데이트 파일에 _internal 디렉터리가 없습니다.")
        elif not path.is_file():
            raise RuntimeError(f"업데이트 파일에 {required} 파일이 없습니다.")
    return version


def _replace_with_retry(source: Path, target: Path, timeout: float = WAIT_TIMEOUT_SECONDS) -> None:
    deadline = time.monotonic() + timeout
    while True:
        try:
            os.replace(source, target)
            return
        except PermissionError:
            if time.monotonic() >= deadline:
                raise
            time.sleep(RETRY_INTERVAL_SECONDS)


def _copy_fsynced(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with source.open("rb") as input_stream, target.open("wb") as output_stream:
        shutil.copyfileobj(input_stream, output_stream, length=1024 * 1024)
        output_stream.flush()
        os.fsync(output_stream.fileno())


def _launch_application(root: Path, transaction_id: str) -> subprocess.Popen[bytes]:
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return subprocess.Popen(
        [str(root / TARGET_NAME), "--post-update", transaction_id],
        cwd=str(root),
        close_fds=True,
        creationflags=creation_flags,
    )


def _wait_for_health(
    process: subprocess.Popen[bytes], marker: Path, timeout: float = WAIT_TIMEOUT_SECONDS
) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if marker.is_file():
            return
        return_code = process.poll()
        if return_code is not None:
            raise RuntimeError(f"새 프로그램이 시작 중 종료되었습니다. exitCode={return_code}")
        time.sleep(RETRY_INTERVAL_SECONDS)
    raise TimeoutError("새 프로그램의 정상 시작 확인 시간이 초과되었습니다.")


def install_legacy_migration(
    *, root: Path, source: Path, wait_pid: int, transaction_id: str, to_version: str
) -> None:
    if not TRANSACTION_ID_PATTERN.fullmatch(transaction_id):
        raise ValueError("Invalid update transaction ID")
    update_root = root / ".update"
    backup_directory = update_root / "backup" / transaction_id
    backup_directory.mkdir(parents=True, exist_ok=False)
    backup_target = backup_directory / TARGET_NAME
    target = root / TARGET_NAME
    next_target = root / f"{TARGET_NAME}.next"
    marker = update_root / "health" / f"{transaction_id}.startup-ok"
    marker.unlink(missing_ok=True)
    process: subprocess.Popen[bytes] | None = None
    switched = False

    try:
        _append_log(update_root, f"legacy migration {transaction_id} started")
        target_source, manifest_version = _load_and_verify_source(source)
        if manifest_version != to_version.lstrip("vV"):
            raise RuntimeError("helper 인수와 manifest 버전이 일치하지 않습니다.")
        _write_state(
            update_root,
            mode="legacy-migration",
            transactionId=transaction_id,
            toVersion=manifest_version,
            status="waiting_for_exit",
            sourcePath=str(source),
            backupPath=str(backup_directory),
        )
        _wait_for_process_exit(wait_pid)

        _write_state(
            update_root,
            mode="legacy-migration",
            transactionId=transaction_id,
            toVersion=manifest_version,
            status="switching",
            sourcePath=str(source),
            backupPath=str(backup_directory),
        )
        if not target.is_file():
            raise FileNotFoundError(f"기존 {TARGET_NAME} 파일이 없습니다.")
        _replace_with_retry(target, backup_target)
        switched = True
        _copy_fsynced(target_source, next_target)
        if _sha256(next_target) != _sha256(target_source):
            raise RuntimeError("설치된 tdm.exe 복사 검증에 실패했습니다.")
        _replace_with_retry(next_target, target)

        _write_state(
            update_root,
            mode="legacy-migration",
            transactionId=transaction_id,
            toVersion=manifest_version,
            status="awaiting_health",
            sourcePath=str(source),
            backupPath=str(backup_directory),
        )
        process = _launch_application(root, transaction_id)
        _wait_for_health(process, marker)
        _atomic_write_json(
            update_root / "legacy-updater-migrated.json",
            {
                "schemaVersion": 1,
                "transactionId": transaction_id,
                "version": manifest_version,
                "completedAt": datetime.now(timezone.utc).isoformat(),
            },
        )
        _write_state(
            update_root,
            mode="legacy-migration",
            transactionId=transaction_id,
            toVersion=manifest_version,
            status="committed",
            backupPath=str(backup_directory),
        )
        _append_log(update_root, f"legacy migration {transaction_id} committed")
        try:
            shutil.rmtree(root / ".update_tmp", ignore_errors=True)
            _prune_directories(update_root / "helper")
        except OSError as cleanup_error:
            _append_log(update_root, f"legacy cleanup failed: {cleanup_error!r}")
    except Exception as exc:
        _append_log(update_root, f"legacy migration {transaction_id} failed: {exc!r}")
        if process is not None and process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
        next_target.unlink(missing_ok=True)
        if switched:
            failed_target = update_root / "failed" / transaction_id / TARGET_NAME
            failed_target.parent.mkdir(parents=True, exist_ok=True)
            if target.exists():
                _replace_with_retry(target, failed_target)
            if backup_target.exists():
                _replace_with_retry(backup_target, target)
        _write_state(
            update_root,
            mode="legacy-migration",
            transactionId=transaction_id,
            toVersion=to_version,
            status="rollback" if target.exists() else "rollback_failed",
            error=type(exc).__name__,
            backupPath=str(backup_directory),
        )
        if (root / "main.exe").is_file():
            subprocess.Popen([str(root / "main.exe")], cwd=str(root), close_fds=True)
        raise


def install_update(
    *, root: Path, source: Path, wait_pid: int, transaction_id: str, to_version: str
) -> None:
    if not TRANSACTION_ID_PATTERN.fullmatch(transaction_id):
        raise ValueError("Invalid update transaction ID")
    update_root = root / ".update"
    backup_directory = update_root / "backup" / transaction_id
    failed_directory = update_root / "failed" / transaction_id
    backup_directory.mkdir(parents=True, exist_ok=False)
    marker = update_root / "health" / f"{transaction_id}.startup-ok"
    marker.unlink(missing_ok=True)
    process: subprocess.Popen[bytes] | None = None
    backed_up: list[str] = []
    installed: list[str] = []

    try:
        _append_log(update_root, f"update {transaction_id} started")
        manifest_version = _load_and_verify_update_source(source)
        if manifest_version != to_version.lstrip("vV"):
            raise RuntimeError("helper 인수와 manifest 버전이 일치하지 않습니다.")
        _write_state(
            update_root,
            mode="update",
            transactionId=transaction_id,
            toVersion=manifest_version,
            status="waiting_for_exit",
            sourcePath=str(source),
            backupPath=str(backup_directory),
        )
        _wait_for_process_exit(wait_pid)
        _write_state(
            update_root,
            mode="update",
            transactionId=transaction_id,
            toVersion=manifest_version,
            status="switching",
            sourcePath=str(source),
            backupPath=str(backup_directory),
        )

        for name in UPDATE_PAYLOAD_NAMES:
            target = root / name
            if target.exists():
                _replace_with_retry(target, backup_directory / name)
                backed_up.append(name)
        for name in UPDATE_PAYLOAD_NAMES:
            _replace_with_retry(source / name, root / name)
            installed.append(name)

        _write_state(
            update_root,
            mode="update",
            transactionId=transaction_id,
            toVersion=manifest_version,
            status="awaiting_health",
            backupPath=str(backup_directory),
        )
        process = _launch_application(root, transaction_id)
        _wait_for_health(process, marker)
        _write_state(
            update_root,
            mode="update",
            transactionId=transaction_id,
            toVersion=manifest_version,
            status="committed",
            backupPath=str(backup_directory),
        )
        _append_log(update_root, f"update {transaction_id} committed")
        try:
            shutil.rmtree(update_root / "transactions" / transaction_id, ignore_errors=True)
            _prune_directories(update_root / "backup")
            _prune_directories(update_root / "helper")
        except OSError as cleanup_error:
            _append_log(update_root, f"update cleanup failed: {cleanup_error!r}")
    except Exception as exc:
        _append_log(update_root, f"update {transaction_id} failed: {exc!r}")
        if process is not None and process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
        failed_directory.mkdir(parents=True, exist_ok=True)
        for name in reversed(installed):
            target = root / name
            if target.exists():
                _replace_with_retry(target, failed_directory / name)
        for name in reversed(backed_up):
            backup = backup_directory / name
            if backup.exists():
                _replace_with_retry(backup, root / name)
        rollback_succeeded = (root / TARGET_NAME).is_file() and all(
            (root / name).exists() for name in backed_up
        )
        _write_state(
            update_root,
            mode="update",
            transactionId=transaction_id,
            toVersion=to_version,
            status="rollback" if rollback_succeeded else "rollback_failed",
            error=type(exc).__name__,
            backupPath=str(backup_directory),
        )
        if (root / TARGET_NAME).is_file():
            subprocess.Popen([str(root / TARGET_NAME)], cwd=str(root), close_fds=True)
        raise


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="TDM update file switcher")
    parser.add_argument(
        "--mode",
        choices=("legacy-migration", "update"),
        default="legacy-migration",
    )
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--wait-pid", type=int, required=True)
    parser.add_argument("--transaction-id", required=True)
    parser.add_argument("--to-version", required=True)
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    try:
        arguments = {
            "root": args.root.resolve(),
            "source": args.source.resolve(),
            "wait_pid": args.wait_pid,
            "transaction_id": args.transaction_id,
            "to_version": args.to_version,
        }
        if args.mode == "update":
            install_update(**arguments)
        else:
            install_legacy_migration(**arguments)
    except Exception:  # noqa: BLE001 - details are persisted to the update log
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
