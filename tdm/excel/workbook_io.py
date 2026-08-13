"""Common workbook open, close, and worksheet validation helpers."""

from __future__ import annotations

import zipfile
from contextlib import AbstractContextManager, closing
from pathlib import Path
from typing import TypeVar

import openpyxl as xl
from openpyxl.worksheet.worksheet import Worksheet

from tdm.domain.errors import NoMatchingSheetException, ReopenFileException
from tdm.excel.errors import ConcurrentWorkbookChangeError

WorkbookT = TypeVar("WorkbookT", bound=xl.Workbook)


def load_workbook(
    path: str | Path,
    *,
    display_name: str | None = None,
    data_only: bool = False,
    read_only: bool = False,
) -> xl.Workbook:
    """Load a workbook and attach its source path for conflict detection."""
    source = Path(path)
    label = display_name or source.stem
    revision_before = None
    if not read_only:
        from tdm.excel.atomic import capture_file_revision

        revision_before = capture_file_revision(source)
    try:
        workbook = xl.load_workbook(
            source,
            data_only=data_only,
            read_only=read_only,
        )
    except PermissionError as exc:
        raise ReopenFileException(
            f"{label} 파일에 접근할 수 없습니다.\n"
            "파일을 직접 연 후 닫으면 문제가 해결될 수 있습니다."
        ) from exc
    except zipfile.BadZipFile as exc:
        raise ReopenFileException(
            f"{label} 파일을 직접 연 후 닫으면 문제가 해결될 수 있습니다."
        ) from exc

    if not read_only:
        from tdm.excel.atomic import capture_file_revision, track_workbook_source

        revision_after = capture_file_revision(source)
        if revision_before != revision_after:
            workbook.close()
            raise ConcurrentWorkbookChangeError(source)
        track_workbook_source(workbook, source, revision_after)
    return workbook


def require_worksheet(workbook: xl.Workbook, sheet_name: str) -> Worksheet:
    try:
        return workbook[sheet_name]
    except KeyError as exc:
        raise NoMatchingSheetException(
            f"'{Path(sheet_name).stem}.xlsx'의 시트명을 '{sheet_name}'으로 변경해 주세요."
        ) from exc


def closing_workbook(workbook: WorkbookT) -> AbstractContextManager[WorkbookT]:
    """Return a standard context manager that always closes the workbook."""
    return closing(workbook)


def close_workbooks(*workbooks: xl.Workbook | None) -> None:
    for workbook in workbooks:
        if workbook is None:
            continue
        try:
            workbook.close()
        except Exception:
            pass
