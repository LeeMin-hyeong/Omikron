"""Microsoft Excel COM adapter isolated from workbook business rules."""

from __future__ import annotations

import os
from pathlib import Path

from tdm.domain.errors import ExcelRequiredException
from tdm.excel.atomic import staged_xlsx_replacement


def recalculate_workbook(file_path: str | Path) -> None:
    """Recalculate a copy with Excel and replace the source only on success."""
    # COM is an optional, Windows-only runtime dependency.  Import it only when
    # recalculation is actually requested so ordinary workbook operations and
    # tests do not require a working Excel/pywin32 installation.
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise ExcelRequiredException(
            "이 기능을 사용하려면 Microsoft Excel과 Excel 연동 구성요소가 필요합니다."
        ) from exc

    with staged_xlsx_replacement(file_path, copy_existing=True) as staging:
        pythoncom.CoInitialize()
        excel = None
        workbook = None
        try:
            try:
                excel = win32com.client.Dispatch("Excel.Application")
            except pythoncom.com_error as exc:
                raise ExcelRequiredException(
                    "이 기능을 사용하려면 Microsoft Excel이 설치되어 있어야 합니다."
                ) from exc

            workbook = excel.Workbooks.Open(os.path.abspath(staging))
            workbook.Save()
        finally:
            try:
                if workbook is not None:
                    workbook.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()
