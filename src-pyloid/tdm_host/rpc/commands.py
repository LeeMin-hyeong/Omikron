"""User-triggered file and data mutation RPC handlers."""

import base64
import os
import webbrowser
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

from pyloid.rpc import RPCContext

import tdm.aisosik.browser
import tdm.aisosik.reader
import tdm.config
import tdm.excel.class_info
import tdm.excel.data_file
import tdm.excel.data_form
import tdm.excel.makeup_test
import tdm.excel.student_info
from tdm.domain.errors import FileOpenException
from tdm.domain.progress import Progress
from tdm.excel.paths import WorkbookPaths
from tdm.excel.transaction import WorkbookSave, save_workbooks_transaction
from tdm.excel.workbook_io import close_workbooks
from tdm_host.rpc.error_codes import RpcErrorCode
from tdm_host.rpc.contracts import MoveStudentRequest, StudentRequest, TextRequest, UrlRequest
from tdm_host.rpc.responses import error_response, failure_response, success_response
from tdm_host.rpc.transport import server
from tdm_host.runtime.files import open_path_cross_platform as _open_path_cross_platform

@server.method()
async def change_data_dir(ctx:RPCContext):
    try:
        new_dir = ctx.pyloid.select_directory_dialog(tdm.config.DATA_DIR)
        if new_dir is None:
            return error_response(RpcErrorCode.CANCELLED, "폴더 선택을 취소했습니다.")
        abspath = os.path.abspath(new_dir)
        tdm.config.change_data_path(abspath)
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="change_data_dir")


@server.method()
async def change_data_file_name(ctx:RPCContext, new_filename:str) -> Dict[str, Any]:
    try:
        request = TextRequest.validate(new_filename, "new_filename")
        tdm.config.change_data_file_name(request.value)
        return success_response()
    except (FileExistsError, FileOpenException) as exc:
        return failure_response(exc, context="change_data_file_name")
    except Exception as exc:
        return failure_response(exc, context="change_data_file_name")


@server.method()
async def open_path(ctx: RPCContext, path: str) -> Dict[str, Any]:
    try:
        request = TextRequest.validate(path, "path")
        _open_path_cross_platform(request.value)
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="open_path")


@server.method()
async def open_url(ctx: RPCContext, url: str) -> Dict[str, Any]:
    try:
        request = UrlRequest.validate(url)
        opened = webbrowser.open(request.url, new=0, autoraise=True)
        if not opened:
            raise RuntimeError("브라우저를 열 수 없습니다.")
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="open_url")


@server.method()
async def make_class_info(ctx: RPCContext):
    try:
        tdm.excel.class_info.make_file()
        return success_response(path=str(WorkbookPaths.current().class_info))
    except Exception as exc:
        return failure_response(exc, context="make_class_info")


@server.method()
async def make_data_file(ctx: RPCContext):
    try:
        class_info = WorkbookPaths.current().class_info
        if not class_info.is_file():
            return error_response(RpcErrorCode.CLASS_INFO_MISSING, "반 정보.xlsx가 먼저 필요합니다.")

        if not tdm.config.DATA_FILE_NAME:
            return error_response(
                RpcErrorCode.DATA_FILE_NAME_MISSING,
                "데이터 파일 이름을 먼저 설정해 주세요.",
            )

        tdm.excel.data_file.make_file()
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="make_data_file")


@server.method()
async def make_student_info(ctx: RPCContext):
    try:
        tdm.excel.student_info.make_file()
        return success_response(path=str(WorkbookPaths.current().student_info))
    except Exception as exc:
        return failure_response(exc, context="make_student_info")


@server.method()
async def make_data_form(ctx: RPCContext):
    try:
        tdm.excel.data_form.make_file()
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="make_data_form")


@server.method()
async def reapply_conditional_format(ctx: RPCContext):
    try:
        warnings = tdm.excel.data_file.conditional_formatting()
        return success_response(warnings=warnings)
    except Exception as exc:
        return failure_response(exc, context="reapply_conditional_format")


@server.method()
async def update_student_info(ctx: RPCContext):
    try:
        tdm.excel.student_info.update_student()
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="update_student_info")


@server.method()
async def add_student(ctx: RPCContext, target_student_name, target_class_name):
    try:
        request = StudentRequest.validate(target_student_name, target_class_name)
        target_student_name = request.student_name
        target_class_name = request.class_name
        if not tdm.aisosik.reader.check_student_exists(target_student_name, target_class_name):
            return error_response(
                RpcErrorCode.AISOSIK_STUDENT_NOT_FOUND,
                f"아이소식에서 {target_class_name} 반의 {target_student_name} 학생을 찾을 수 없습니다.",
            )

        warnings = tdm.excel.data_file.add_student(target_student_name, target_class_name)

        tdm.excel.student_info.add_student(target_student_name)

        return success_response(warnings=warnings)
    except Exception as exc:
        return failure_response(exc, context="add_student")


@server.method()
async def remove_student(ctx: RPCContext, target_class_name, target_student_name):
    try:
        request = StudentRequest.validate(target_student_name, target_class_name)
        target_student_name = request.student_name
        target_class_name = request.class_name
        tdm.excel.data_file.delete_student(target_class_name, target_student_name)

        if not tdm.excel.data_file.check_student_exists(target_student_name):
            tdm.excel.student_info.delete_student(target_student_name)

        return success_response()
    except Exception as exc:
        return failure_response(exc, context="remove_student")


@server.method()
async def move_student(ctx: RPCContext, target_student_name, target_class_name, current_class_name):
    try:
        request = MoveStudentRequest.validate(
            target_student_name, target_class_name, current_class_name
        )
        target_student_name = request.student_name
        target_class_name = request.target_class_name
        current_class_name = request.current_class_name
        if not tdm.aisosik.reader.check_student_exists(target_student_name, target_class_name):
            return error_response(
                RpcErrorCode.AISOSIK_STUDENT_NOT_FOUND,
                f"아이소식에서 {target_class_name} 반의 {target_student_name} 학생을 찾을 수 없습니다.",
            )

        warnings = tdm.excel.data_file.move_student(
            target_student_name,
            target_class_name,
            current_class_name,
        )

        return success_response(warnings=warnings)
    except Exception as exc:
        return failure_response(exc, context="move_student")


@server.method()
async def change_class_info(ctx: RPCContext, target_class_name, target_teacher_name):
    try:
        tdm.excel.class_info.change_class_info(target_class_name, target_teacher_name)

        tdm.excel.data_file.change_class_info(target_class_name, target_teacher_name)

        return success_response()
    except Exception as exc:
        return failure_response(exc, context="change_class_info")


@server.method()
async def make_temp_class_info(ctx: RPCContext, new_class_list):
    try:
        filepath = tdm.excel.class_info.make_temp_file_for_update(new_class_list)
        return success_response(path=filepath)
    except Exception as exc:
        return failure_response(exc, context="make_temp_class_info")


@server.method()
async def update_class(ctx: RPCContext):
    try:
        tdm.excel.data_file.update_class()
        tdm.excel.class_info.update_class()
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="update_class")
    finally:
        try:
            tdm.excel.class_info.delete_temp()
        except OSError:
            pass


@server.method()
async def delete_class_info_temp(ctx: RPCContext):
    try:
        tdm.excel.class_info.delete_temp()
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="delete_class_info_temp")


@server.method()
async def save_individual_result(ctx: RPCContext, student_name:str, class_name:str, test_name:str, target_row:int, target_col:int, test_score:int|float, makeup_test_check:bool, makeup_test_date:dict):
    data_workbook = None
    makeup_workbook = None
    try:
        prog_warnings: list[str] = []
        def capture_progress(payload: dict) -> None:
            if payload.get("level") == "warning" and payload.get("message"):
                prog_warnings.append(str(payload["message"]))

        prog = Progress(capture_progress, total=3)

        for k, v in makeup_test_date.items():
            makeup_test_date[k] = datetime.strptime(v, "%Y-%m-%d")

        data_workbook, test_average = tdm.excel.data_file.prepare_individual_test_data(
            target_row,
            target_col,
            test_score,
        )

        if test_score < 80 and not makeup_test_check:
            makeup_workbook = tdm.excel.makeup_test.prepare_individual_makeup_test(
                student_name,
                class_name,
                test_name,
                test_score,
                makeup_test_date,
                prog,
            )

        paths = WorkbookPaths.current()
        saves = [WorkbookSave(data_workbook, paths.data_file)]
        if makeup_workbook is not None:
            saves.append(WorkbookSave(makeup_workbook, paths.makeup_test))
        save_workbooks_transaction(
            saves,
            operation="개인 시험 결과 저장",
            paths=paths,
        )

        tdm.aisosik.browser.send_individual_test_message(student_name, class_name, test_name, test_score, test_average, makeup_test_check, makeup_test_date, prog)

        return success_response(warnings=prog_warnings)
    except Exception as exc:
        return failure_response(exc, context="save_individual_result")
    finally:
        close_workbooks(data_workbook, makeup_workbook)


@server.method()
async def save_retest_result(ctx: RPCContext, target_row:int, makeup_test_score:str):
    try:
        tdm.excel.makeup_test.save_makeup_test_result(target_row, makeup_test_score)
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="save_retest_result")


@server.method()
async def change_data_file_name_by_select(ctx: RPCContext):
    try:
        selected_file = ctx.pyloid.open_file_dialog(f"{tdm.config.DATA_DIR}/data")
        if not selected_file:
            return error_response(RpcErrorCode.CANCELLED, "파일 선택을 취소했습니다.")

        new_filename = Path(selected_file).stem

        tdm.config.change_data_file_name_by_select(new_filename)
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="change_data_file_name_by_select")


@server.method()
async def open_file_picker(ctx: RPCContext):
    try:
        selected_file = ctx.pyloid.open_file_dialog(tdm.config.DATA_DIR)
        if not selected_file:
            return error_response(RpcErrorCode.CANCELLED, "파일 선택을 취소했습니다.")

        path_obj = Path(selected_file)
        if path_obj.suffix.lower() != ".xlsx":
            return error_response(
                RpcErrorCode.UNSUPPORTED_FILE_TYPE,
                "지원하지 않는 파일 형식입니다. .xlsx 파일만 선택해 주세요.",
            )

        file_b64 = base64.b64encode(path_obj.read_bytes()).decode()

        return success_response(path=str(path_obj), name=path_obj.name, b64=file_b64)
    except Exception as exc:
        return failure_response(exc, context="open_file_picker")
