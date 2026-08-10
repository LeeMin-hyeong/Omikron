"""Process entry points for long-running workbook and browser operations."""

import multiprocessing
import time
import traceback
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
from tdm_host.runtime.files import cleanup_temp, decode_xlsx_upload_to_temp


def update_class_job_process(job_id: str, q: multiprocessing.Queue) -> None:
    def emit(payload: dict) -> None:
        q.put(payload)

    progress = Progress(emit, total=5)
    progress.info("반 업데이트 준비중...")
    try:
        tdm.excel.data_file.update_class(progress)
        progress.step("반 정보 파일 최신화 중...")
        tdm.excel.class_info.update_class(progress)
        progress.done("반 업데이트가 완료되었습니다.")
    except ExcelRequiredException as exc:
        progress.error(str(exc))
    except Exception:
        progress.error("예상치 못한 오류가 발생했습니다.", detail=traceback.format_exc())
    finally:
        tdm.excel.class_info.delete_temp()


def send_exam_message_job_process(
    job_id: str,
    q: multiprocessing.Queue,
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
        temp_file = decode_xlsx_upload_to_temp(filename, b64)
        try:
            tdm.excel.data_form.data_validation(str(temp_file))
        except tdm.excel.data_form.DataValidationException as exc:
            progress.error(f"데이터 검증 오류가 발생하였습니다:\n{exc}")
            return
        progress.step("데이터 입력 양식 검증 완료")

        for key, value in makeup_test_date.items():
            makeup_test_date[key] = datetime.strptime(value, "%Y-%m-%d")

        try:
            tdm.aisosik.browser.send_test_result_message(
                str(temp_file), makeup_test_date, progress
            )
        except ChromeDriverVersionMismatchException as exc:
            progress.error(str(exc))
            return
        except Exception as exc:
            progress.error(f"메시지 작성 중 오류가 발생했습니다:\n {exc}")

        progress.step("작업 완료")
        progress.done("메시지 작성이 완료되었습니다. 전송 전 내용을 확인하세요.")
    except Exception:
        progress.error("예상치 못한 오류가 발생했습니다.", detail=traceback.format_exc())
    finally:
        if temp_file:
            cleanup_temp(temp_file)


def save_exam_job_process(
    job_id: str,
    q: multiprocessing.Queue,
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
            "total": 4,
            "level": "info",
            "status": "running",
            "message": "작업을 준비하고 있습니다.",
            "warnings": [],
        }
    )

    temp_file: Optional[Path] = None
    try:
        temp_file = decode_xlsx_upload_to_temp(filename, b64)
        try:
            tdm.excel.data_form.data_validation(str(temp_file))
        except tdm.excel.data_form.DataValidationException as exc:
            progress.error(f"데이터 검증 오류가 발생하였습니다:\n {exc}")
            return
        progress.step("데이터 입력 양식 검증 완료")

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
            progress.error(str(exc))
            return
        except NoMatchingSheetException as exc:
            progress.error(f"파일에서 목표 시트를 찾을 수 없습니다:\n {exc}")
            return
        except tdm.excel.data_file.NoReservedColumnError as exc:
            progress.error(f"파일에 필수 열이 없습니다:\n {exc}")
            return

        try:
            tdm.excel.data_file.save(datafile_workbook)
            tdm.excel.makeup_test.save(makeuptest_workbook)
        except FileOpenException as exc:
            progress.error(f"파일이 열려 있습니다:\n {exc}")
            return

        progress.step("파일 저장 완료")
        progress.done("데이터 저장을 완료하였습니다.")
    except Exception:
        progress.error("예상치 못한 오류가 발생했습니다.", detail=traceback.format_exc())
        return
    finally:
        tdm.excel.data_file.delete_temp()
        if temp_file:
            cleanup_temp(temp_file)
