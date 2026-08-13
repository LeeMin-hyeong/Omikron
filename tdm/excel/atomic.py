"""Atomic XLSX persistence helpers.

Workbooks are written to a unique file beside the destination and moved into
place only after the temporary XLSX can be opened as a ZIP archive. Keeping the
temporary file in the destination directory makes ``os.replace`` stay on the
same filesystem, including when the directory is an SMB-mounted NAS share.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import time
import zipfile
from contextlib import AbstractContextManager
from dataclasses import asdict, dataclass
from hashlib import sha256
from pathlib import Path
from types import TracebackType
from typing import Any, Protocol

from tdm.excel.errors import ConcurrentWorkbookChangeError, WorkbookBusyError
from tdm.excel.paths import STAGING_SUFFIX


class WorkbookLike(Protocol):
    def save(self, filename: str | os.PathLike[str]) -> None: ...


@dataclass(frozen=True)
class FileRevision:
    size: int
    modified_ns: int
    digest: str

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, value: dict[str, Any] | None) -> "FileRevision | None":
        if value is None:
            return None
        return cls(
            size=int(value["size"]),
            modified_ns=int(value["modified_ns"]),
            digest=str(value["digest"]),
        )


@dataclass(frozen=True)
class PreparedXlsx:
    target: Path
    staging: Path
    expected_revision: FileRevision | None | object
    staged_revision: FileRevision


@dataclass(frozen=True)
class ReplaceRetryPolicy:
    delays: tuple[float, ...] = (0.25, 0.5, 1.0)


_UNTRACKED = object()
_SOURCE_PATH_ATTRIBUTE = "_tdm_source_path"
_SOURCE_REVISION_ATTRIBUTE = "_tdm_source_revision"
_BUSY_ERROR_CODES = {5, 32, 33}


def capture_file_revision(path: str | os.PathLike[str]) -> FileRevision | None:
    target = Path(path)
    try:
        stat = target.stat()
    except FileNotFoundError:
        return None

    digest = sha256()
    with target.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return FileRevision(stat.st_size, stat.st_mtime_ns, digest.hexdigest())


def track_workbook_source(
    workbook: WorkbookLike,
    source_path: str | os.PathLike[str],
    revision: FileRevision | None,
) -> None:
    setattr(workbook, _SOURCE_PATH_ATTRIBUTE, str(Path(source_path).absolute()))
    setattr(workbook, _SOURCE_REVISION_ATTRIBUTE, revision)


def workbook_source_revision(
    workbook: WorkbookLike,
    target_path: str | os.PathLike[str],
) -> FileRevision | None | object:
    source = getattr(workbook, _SOURCE_PATH_ATTRIBUTE, None)
    if source is None:
        return _UNTRACKED
    if Path(source).absolute() != Path(target_path).absolute():
        return _UNTRACKED
    return getattr(workbook, _SOURCE_REVISION_ATTRIBUTE, _UNTRACKED)


def validate_xlsx(path: Path) -> None:
    with zipfile.ZipFile(path, "r") as archive:
        names = set(archive.namelist())
        required = {"[Content_Types].xml", "xl/workbook.xml"}
        if not required.issubset(names):
            raise zipfile.BadZipFile(f"필수 XLSX 항목이 없습니다: {path.name}")


def flush_file(path: Path) -> None:
    with path.open("rb+") as stream:
        stream.flush()
        os.fsync(stream.fileno())


def make_staging_path(target: Path) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    descriptor, name = tempfile.mkstemp(
        dir=target.parent,
        prefix=f".{target.stem}.",
        suffix=STAGING_SUFFIX,
    )
    os.close(descriptor)
    return Path(name)


def prepare_workbook(
    workbook: WorkbookLike,
    target_path: str | os.PathLike[str],
) -> PreparedXlsx:
    target = Path(target_path)
    staging = make_staging_path(target)
    try:
        workbook.save(staging)
        validate_xlsx(staging)
        flush_file(staging)
        return PreparedXlsx(
            target=target,
            staging=staging,
            expected_revision=workbook_source_revision(workbook, target),
            staged_revision=capture_file_revision(staging),
        )
    except Exception:
        staging.unlink(missing_ok=True)
        raise


def assert_target_unchanged(
    target: Path,
    expected_revision: FileRevision | None | object,
) -> None:
    if expected_revision is _UNTRACKED:
        return
    if capture_file_revision(target) != expected_revision:
        raise ConcurrentWorkbookChangeError(target)


def _replace_with_retry(
    source: Path,
    target: Path,
    retry_policy: ReplaceRetryPolicy,
) -> None:
    attempts = len(retry_policy.delays) + 1
    for attempt in range(attempts):
        try:
            os.replace(source, target)
            return
        except OSError as exc:
            error_code = int(getattr(exc, "winerror", 0) or 0)
            busy = isinstance(exc, PermissionError) or error_code in _BUSY_ERROR_CODES
            if not busy:
                raise
            if attempt >= len(retry_policy.delays):
                raise WorkbookBusyError(target) from exc
            time.sleep(retry_policy.delays[attempt])


def commit_prepared_workbook(
    prepared: PreparedXlsx,
    *,
    retry_policy: ReplaceRetryPolicy | None = None,
) -> FileRevision:
    assert_target_unchanged(prepared.target, prepared.expected_revision)
    _replace_with_retry(
        prepared.staging,
        prepared.target,
        retry_policy or ReplaceRetryPolicy(),
    )
    return capture_file_revision(prepared.target)


def discard_prepared_workbook(prepared: PreparedXlsx) -> None:
    try:
        prepared.staging.unlink(missing_ok=True)
    except OSError:
        pass


class _StagedXlsxReplacement(AbstractContextManager[Path]):
    def __init__(
        self,
        target_path: str | os.PathLike[str],
        *,
        copy_existing: bool,
    ) -> None:
        self.target = Path(target_path)
        self.copy_existing = copy_existing
        self.staging: Path | None = None
        self.expected_revision: FileRevision | None | object = _UNTRACKED

    def __enter__(self) -> Path:
        self.staging = make_staging_path(self.target)
        if self.copy_existing:
            self.expected_revision = capture_file_revision(self.target)
            try:
                shutil.copy2(self.target, self.staging)
            except Exception:
                self.staging.unlink(missing_ok=True)
                raise
        return self.staging

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> bool:
        staging = self.staging
        if staging is None:
            return False
        try:
            if exc_type is None:
                validate_xlsx(staging)
                flush_file(staging)
                assert_target_unchanged(self.target, self.expected_revision)
                _replace_with_retry(staging, self.target, ReplaceRetryPolicy())
        finally:
            try:
                staging.unlink(missing_ok=True)
            except OSError:
                pass
        return False


def staged_xlsx_replacement(
    target_path: str | os.PathLike[str],
    *,
    copy_existing: bool = False,
) -> AbstractContextManager[Path]:
    """Create a same-directory staging file and atomically replace on success."""
    return _StagedXlsxReplacement(target_path, copy_existing=copy_existing)


def atomic_save_workbook(
    workbook: WorkbookLike,
    target_path: str | os.PathLike[str],
    *,
    update_source: bool = True,
) -> None:
    """Save an openpyxl-compatible workbook without modifying the target in place."""
    prepared = prepare_workbook(workbook, target_path)
    try:
        revision = commit_prepared_workbook(prepared)
        if update_source:
            track_workbook_source(workbook, prepared.target, revision)
    finally:
        discard_prepared_workbook(prepared)
