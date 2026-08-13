"""Shared workbook backup and retention policy."""

from __future__ import annotations

import os
import shutil
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

from tdm.excel.atomic import staged_xlsx_replacement
from tdm.excel.paths import WorkbookPaths, backup_filename


@dataclass(frozen=True)
class BackupRetentionPolicy:
    max_count: int = 30
    max_age_days: int = 90


DEFAULT_BACKUP_POLICY = BackupRetentionPolicy()


def _available_backup_path(directory: Path, stem: str, when: datetime) -> Path:
    candidate = directory / backup_filename(stem, when)
    suffix = 1
    while candidate.exists():
        timestamp = when.strftime("%Y%m%d%H%M%S")
        candidate = directory / f"{stem}({timestamp}-{suffix}).xlsx"
        suffix += 1
    return candidate


def prune_backups(
    directory: str | Path,
    stem: str,
    *,
    policy: BackupRetentionPolicy = DEFAULT_BACKUP_POLICY,
    now: datetime | None = None,
) -> list[Path]:
    backup_dir = Path(directory)
    if not backup_dir.exists():
        return []

    current_time = now or datetime.now()
    cutoff = current_time - timedelta(days=max(0, policy.max_age_days))
    candidates = sorted(
        backup_dir.glob(f"{stem}(*).xlsx"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    removed: list[Path] = []
    for index, path in enumerate(candidates):
        too_many = index >= max(1, policy.max_count)
        too_old = datetime.fromtimestamp(path.stat().st_mtime) < cutoff
        if not (too_many or too_old):
            continue
        try:
            path.unlink()
            removed.append(path)
        except OSError:
            pass
    return removed


def create_backup(
    source_path: str | Path,
    *,
    stem: str | None = None,
    backup_dir: str | Path | None = None,
    policy: BackupRetentionPolicy = DEFAULT_BACKUP_POLICY,
    now: datetime | None = None,
) -> Path:
    source = Path(source_path)
    if not source.is_file():
        raise FileNotFoundError(source)

    directory = Path(backup_dir) if backup_dir is not None else WorkbookPaths.current().backup_dir
    directory.mkdir(parents=True, exist_ok=True)
    backup_stem = stem or source.stem
    created_at = now or datetime.now()
    target = _available_backup_path(directory, backup_stem, created_at)

    with staged_xlsx_replacement(target) as staging:
        shutil.copyfile(source, staging)
        timestamp = created_at.timestamp()
        try:
            os.utime(staging, (timestamp, timestamp))
        except OSError:
            pass

    prune_backups(directory, backup_stem, policy=policy, now=created_at)
    return target
