"""Process entry points for long-running workbook and browser operations."""

import multiprocessing
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

import tdm.aisosik.browser
import tdm.excel.class_info
import tdm.excel.data_file
import tdm.excel.data_form
import tdm.excel.makeup_test
from tdm.domain.errors import (
    ChromeDriverVersionMismatchException,
    ExcelRequiredException,
    FileOpenException,
    NoMatchingSheetException,
)
from tdm.domain.progress import Progress
from tdm.excel.paths import WorkbookPaths
from tdm.excel.transaction import WorkbookSave, save_workbooks_transaction
from tdm.excel.workbook_io import close_workbooks
from tdm_host.runtime.files import cleanup_temp, decode_xlsx_upload_to_temp
from tdm_host.rpc.responses import failure_response


def _cancelled(cancel_event: Any, progress: Progress) -> bool:
    if not cancel_event.is_set():
        return False
    progress.emit_cb(
        {
            "ts": time.time(),
            "step": progress.step_no,
            "total": progress.total,
            "level": "warning",
            "status": "cancelled",
            "message": "작업이 취소되었습니다.",
        }
    )
    return True


def _progress_failure(progress: Progress, exc: BaseException, *, context: str) -> None:
    response = failure_response(exc, context=context)
    progress.error(
        response["error"],
        code=response["code"],
        detail=response.get("detail"),
    )


def update_class_job_process(
    job_id: str, q: multiprocessing.Queue, cancel_event: Any
) -> None:
    def emit(payload: dict) -> None:
        q.put(payload)

    progress = Progress(emit, total=5)
    progress.info("반 업데이트 준비중...")
    try:
        if _cancelled(cancel_event, progress):
            return
        tdm.excel.data_file.update_class(progress)
        if _cancelled(cancel_event, progress):
            return
        progress.step("반 정보 파일 최신화 중...")
        tdm.excel.class_info.update_class(progress)
        progress.done("반 업데이트가 완료되었습니다.")
    except ExcelRequiredException as exc:
        _progress_failure(progress, exc, context="update_class_job")
    except Exception as exc:
        _progress_failure(progress, exc, context="update_class_job")
    finally:
        tdm.excel.class_info.delete_temp()


def send_exam_message_job_process(
    job_id: str,
    q: multiprocessing.Queue,
    cancel_event: Any,
    *,
    filename: str,
    b64: str,
    makeup_test_date: Dict[str, Any],
) -> None:
    def emit(payload: dict) -> None:
        q.put(payload)

    progress = Progress(emit, total=3)
    emit(
        {
            "ts": time.time(),
            "step": 0,
            "total": 3,
            "level": "info",
            "status": "running",
            "message": "작업을 준비하고 있습니다.",
            "warnings": [],
        }
    )

    temp_file: Optional[Path] = None
    try:
        if _cancelled(cancel_event, progress):
            return
        temp_file = decode_xlsx_upload_to_temp(filename, b64)
        try:
            tdm.excel.data_form.data_validation(str(temp_file))
        except tdm.excel.data_form.DataValidationException as exc:
            _progress_failure(progress, exc, context="send_exam_message_job.validate")
            return
        progress.step("데이터 입력 양식 검증 완료")
        if _cancelled(cancel_event, progress):
            return

        for key, value in makeup_test_date.items():
            makeup_test_date[key] = datetime.strptime(value, "%Y-%m-%d")

        try:
            tdm.aisosik.browser.send_test_result_message(
                str(temp_file), makeup_test_date, progress
            )
        except ChromeDriverVersionMismatchException as exc:
            _progress_failure(progress, exc, context="send_exam_message_job.browser")
            return
        except Exception as exc:
            _progress_failure(progress, exc, context="send_exam_message_job.browser")
            return

        progress.step("작업 완료")
        progress.done("메시지 작성이 완료되었습니다. 전송 전 내용을 확인하세요.")
    except Exception as exc:
        _progress_failure(progress, exc, context="send_exam_message_job")
    finally:
        if temp_file:
            cleanup_temp(temp_file)


def save_exam_job_process(
    job_id: str,
    q: multiprocessing.Queue,
    cancel_event: Any,
    *,
    filename: str,
    b64: str,
    makeup_test_date: Dict[str, Any],
) -> None:
    def emit(payload: dict) -> None:
        q.put(payload)

    progress = Progress(emit, total=6)
    emit(
        {
            "ts": time.time(),
            "step": 0,
            "total": 6,
            "level": "info",
            "status": "running",
            "message": "작업을 준비하고 있습니다.",
            "warnings": [],
        }
    )

    temp_file: Optional[Path] = None
    datafile_workbook = None
    makeuptest_workbook = None
    try:
        if _cancelled(cancel_event, progress):
            return
        temp_file = decode_xlsx_upload_to_temp(filename, b64)
        try:
            tdm.excel.data_form.data_validation(str(temp_file))
        except tdm.excel.data_form.DataValidationException as exc:
            _progress_failure(progress, exc, context="save_exam_job.validate")
            return
        progress.step("데이터 입력 양식 검증 완료")
        if _cancelled(cancel_event, progress):
            return

        for key, value in makeup_test_date.items():
            makeup_test_date[key] = datetime.strptime(value, "%Y-%m-%d")

        try:
            datafile_workbook = tdm.excel.data_file.save_test_data(
                str(temp_file), progress
            )
            makeuptest_workbook = tdm.excel.makeup_test.save_makeup_test_list(
                str(temp_file), makeup_test_date, progress
            )
            progress.step("재시험 명단 입력 완료")
        except ExcelRequiredException as exc:
            _progress_failure(progress, exc, context="save_exam_job.prepare")
            return
        except NoMatchingSheetException as exc:
            _progress_failure(progress, exc, context="save_exam_job.prepare")
            return
        except tdm.excel.data_file.NoReservedColumnError as exc:
            _progress_failure(progress, exc, context="save_exam_job.prepare")
            return

        try:
            paths = WorkbookPaths.current()
            save_workbooks_transaction(
                [
                    WorkbookSave(datafile_workbook, paths.data_file),
                    WorkbookSave(makeuptest_workbook, paths.makeup_test),
                ],
                operation="시험 결과 저장",
                paths=paths,
            )
        except FileOpenException as exc:
            _progress_failure(progress, exc, context="save_exam_job.commit")
            return

        progress.step("파일 저장 완료")
        progress.done("데이터 저장을 완료하였습니다.")
    except Exception as exc:
        _progress_failure(progress, exc, context="save_exam_job")
        return
    finally:
        close_workbooks(datafile_workbook, makeuptest_workbook)
        tdm.excel.data_file.delete_temp()
        if temp_file:
            cleanup_temp(temp_file)
