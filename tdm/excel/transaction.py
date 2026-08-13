"""Recoverable multi-workbook save transactions without a global lock."""

from __future__ import annotations

import json
import shutil
import time
import uuid
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any

from tdm.excel.atomic import (
    FileRevision,
    PreparedXlsx,
    WorkbookLike,
    assert_target_unchanged,
    capture_file_revision,
    commit_prepared_workbook,
    discard_prepared_workbook,
    flush_file,
    prepare_workbook,
    staged_xlsx_replacement,
    track_workbook_source,
    validate_xlsx,
)
from tdm.excel.errors import (
    WorkbookRecoveryRequiredError,
    WorkbookTransactionError,
)
from tdm.excel.paths import STAGING_SUFFIX, WorkbookPaths
from tdm.excel.metadata import atomic_write_json


@dataclass(frozen=True)
class WorkbookSave:
    workbook: WorkbookLike
    target: Path


@dataclass(frozen=True)
class RecoveryResult:
    transaction_id: str
    action: str


def _same_content(
    first: FileRevision | None,
    second: FileRevision | None,
) -> bool:
    if first is None or second is None:
        return first is second
    return first.size == second.size and first.digest == second.digest


def _revision_dict(revision: FileRevision | None) -> dict[str, Any] | None:
    return revision.to_dict() if revision is not None else None


def _manifest_entry(
    prepared: PreparedXlsx,
    rollback_path: Path | None,
) -> dict[str, Any]:
    expected = prepared.expected_revision
    if not isinstance(expected, (FileRevision, type(None))):
        raise TypeError("transaction entry must have a concrete expected revision")
    return {
        "target": str(prepared.target.absolute()),
        "staging": str(prepared.staging.absolute()),
        "rollback": str(rollback_path.absolute()) if rollback_path else None,
        "originalRevision": _revision_dict(expected),
        "stagedRevision": prepared.staged_revision.to_dict(),
    }


def _restore_entry(entry: dict[str, Any]) -> bool:
    target = Path(entry["target"])
    original = FileRevision.from_dict(entry.get("originalRevision"))
    staged = FileRevision.from_dict(entry.get("stagedRevision"))
    current = capture_file_revision(target)

    if _same_content(current, original):
        return False
    if not _same_content(current, staged):
        raise WorkbookRecoveryRequiredError(
            f"{target.name} 파일이 미완료 작업 이후 다시 변경되어 자동 복구하지 않았습니다."
        )

    rollback_value = entry.get("rollback")
    if rollback_value:
        rollback = Path(rollback_value)
        if not rollback.is_file():
            raise WorkbookRecoveryRequiredError(
                f"복구 파일이 없어 {target.name} 파일을 자동 복구할 수 없습니다."
            )
        with staged_xlsx_replacement(target) as staging:
            shutil.copy2(rollback, staging)
    else:
        target.unlink(missing_ok=True)
    return True


def _cleanup_transaction(directory: Path, entries: list[dict[str, Any]]) -> None:
    for entry in entries:
        for key in ("staging", "rollback"):
            value = entry.get(key)
            if not value:
                continue
            try:
                Path(value).unlink(missing_ok=True)
            except OSError:
                pass
    try:
        (directory / "manifest.json").unlink(missing_ok=True)
        directory.rmdir()
    except OSError:
        pass


def save_workbooks_transaction(
    saves: list[WorkbookSave],
    *,
    operation: str,
    paths: WorkbookPaths | None = None,
) -> None:
    """Stage every workbook, then commit all targets with rollback on failure."""
    if not saves:
        return

    workbook_paths = paths or WorkbookPaths.current()
    workbook_paths.ensure_directories()
    transaction_id = uuid.uuid4().hex
    transaction_dir = workbook_paths.transaction_dir / transaction_id
    transaction_dir.mkdir(parents=True, exist_ok=False)
    manifest_path = transaction_dir / "manifest.json"
    prepared_items: list[tuple[WorkbookSave, PreparedXlsx]] = []
    manifest_entries: list[dict[str, Any]] = []
    committed_entries: list[dict[str, Any]] = []

    try:
        for index, save_item in enumerate(saves):
            prepared = prepare_workbook(save_item.workbook, save_item.target)
            if not isinstance(prepared.expected_revision, (FileRevision, type(None))):
                prepared = replace(
                    prepared,
                    expected_revision=capture_file_revision(prepared.target),
                )

            rollback: Path | None = None
            if prepared.expected_revision is not None:
                rollback = transaction_dir / f"rollback-{index}.xlsx"
                shutil.copy2(prepared.target, rollback)
                validate_xlsx(rollback)
                flush_file(rollback)

            prepared_items.append((save_item, prepared))
            manifest_entries.append(_manifest_entry(prepared, rollback))

        manifest: dict[str, Any] = {
            "version": 1,
            "transactionId": transaction_id,
            "operation": operation,
            "createdAt": time.time(),
            "state": "prepared",
            "entries": manifest_entries,
        }
        atomic_write_json(manifest_path, manifest)

        for _, prepared in prepared_items:
            assert_target_unchanged(prepared.target, prepared.expected_revision)

        manifest["state"] = "committing"
        atomic_write_json(manifest_path, manifest)

        for index, (save_item, prepared) in enumerate(prepared_items):
            revision = commit_prepared_workbook(prepared)
            track_workbook_source(save_item.workbook, prepared.target, revision)
            committed_entries.append(manifest_entries[index])

        manifest["state"] = "committed"
        atomic_write_json(manifest_path, manifest)
    except Exception as exc:
        rollback_error: Exception | None = None
        for entry in reversed(committed_entries):
            try:
                _restore_entry(entry)
            except Exception as restore_exc:
                rollback_error = restore_exc
                break
        if rollback_error is not None:
            raise WorkbookTransactionError(
                f"'{operation}' 저장과 자동 복구에 모두 실패했습니다: {rollback_error}"
            ) from exc
        _cleanup_transaction(transaction_dir, manifest_entries)
        raise
    finally:
        for _, prepared in prepared_items:
            discard_prepared_workbook(prepared)

    _cleanup_transaction(transaction_dir, manifest_entries)


def recover_pending_transactions(
    paths: WorkbookPaths | None = None,
) -> list[RecoveryResult]:
    workbook_paths = paths or WorkbookPaths.current()
    transaction_root = workbook_paths.transaction_dir
    if not transaction_root.exists():
        return []

    results: list[RecoveryResult] = []
    for directory in sorted(path for path in transaction_root.iterdir() if path.is_dir()):
        manifest_path = directory / "manifest.json"
        if not manifest_path.is_file():
            continue
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            entries = list(manifest["entries"])
            transaction_id = str(manifest.get("transactionId") or directory.name)
        except (OSError, ValueError, KeyError, TypeError) as exc:
            raise WorkbookRecoveryRequiredError(
                f"손상된 저장 작업 기록을 확인해야 합니다: {manifest_path}"
            ) from exc

        all_committed = all(
            _same_content(
                capture_file_revision(Path(entry["target"])),
                FileRevision.from_dict(entry.get("stagedRevision")),
            )
            for entry in entries
        )
        if all_committed:
            _cleanup_transaction(directory, entries)
            results.append(RecoveryResult(transaction_id, "finalized"))
            continue

        changed = False
        for entry in reversed(entries):
            changed = _restore_entry(entry) or changed
        _cleanup_transaction(directory, entries)
        results.append(RecoveryResult(transaction_id, "rolled_back" if changed else "cleaned"))
    return results


def cleanup_stale_staging_files(
    root: str | Path,
    *,
    older_than_seconds: float = 24 * 60 * 60,
) -> list[Path]:
    base = Path(root)
    if not base.exists():
        return []

    referenced: set[Path] = set()
    transaction_root = base / ".tdm" / "transactions"
    if transaction_root.exists():
        for manifest_path in transaction_root.glob("*/manifest.json"):
            try:
                manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
                referenced.update(
                    Path(entry["staging"]).absolute()
                    for entry in manifest.get("entries", [])
                    if entry.get("staging")
                )
            except (OSError, ValueError, TypeError):
                continue

    cutoff = time.time() - max(0.0, older_than_seconds)
    removed: list[Path] = []
    for path in base.rglob(f"*{STAGING_SUFFIX}"):
        if path.absolute() in referenced:
            continue
        try:
            if path.stat().st_mtime > cutoff:
                continue
            path.unlink()
            removed.append(path)
        except OSError:
            pass
    return removed
