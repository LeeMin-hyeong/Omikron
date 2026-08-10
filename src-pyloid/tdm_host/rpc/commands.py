"""User-triggered file and data mutation RPC handlers."""

import base64
import os
import traceback
import uuid
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
from tdm_host.jobs.registry import make_emit
from tdm_host.rpc.transport import server
from tdm_host.runtime.files import open_path_cross_platform as _open_path_cross_platform

@server.method()
async def change_data_dir(ctx:RPCContext):
    try:
        new_dir = ctx.pyloid.select_directory_dialog(tdm.config.DATA_DIR)
        if new_dir is None: return {"ok": False}
        abspath = os.path.abspath(new_dir)
        tdm.config.change_data_path(abspath)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def change_data_file_name(ctx:RPCContext, new_filename:str) -> Dict[str, Any]:
    try:
        tdm.config.change_data_file_name(new_filename)
        return {"ok": True}
    except FileExistsError as e:
        return {"ok": False, "error": str(e)}
    except FileOpenException as e:
        return {"ok": False, "error": str(e)}
    except Exception as e:
        return {"ok": False, "error": f"알 수 없는 에러가 발생하였습니다: {traceback.format_exc()}"}


@server.method()
async def open_path(ctx: RPCContext, path: str) -> Dict[str, Any]:
    try:
        _open_path_cross_platform(path)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": f"알 수 없는 에러가 발생하였습니다: {traceback.format_exc()}"}


@server.method()
async def open_url(ctx: RPCContext, url: str) -> Dict[str, Any]:
    try:
        if not url:
            raise ValueError("URL is empty.")
        if not url.startswith(("http://", "https://")):
            raise ValueError("지원하지 않는 URL 입니다.")
        opened = webbrowser.open(url, new=0, autoraise=True)
        if not opened:
            raise RuntimeError("브라우저를 열 수 없습니다.")
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def make_class_info(ctx: RPCContext):
    try:
        tdm.excel.class_info.make_file()
        return {"ok": True, "path": str(Path(tdm.config.DATA_DIR) / '반 정보.xlsx')}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def make_data_file(ctx: RPCContext):
    try:
        cwd = Path(tdm.config.DATA_DIR)
        class_info = cwd / "반 정보.xlsx"
        if not class_info.is_file():
            return {"ok": False, "error": "반 정보.xlsx가 먼저 필요합니다."}

        if not tdm.config.DATA_FILE_NAME:
            return {"ok": False, "error": "config.json의 dataFileName을 설정해 주세요."}

        tdm.excel.data_file.make_file()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def make_student_info(ctx: RPCContext):
    try:
        tdm.excel.student_info.make_file()
        return {"ok": True, "path": str(Path(tdm.config.DATA_DIR) / '학생 정보.xlsx')}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def make_data_form(ctx: RPCContext):
    try:
        tdm.excel.data_form.make_file()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def reapply_conditional_format(ctx: RPCContext):
    try:
        warnings = tdm.excel.data_file.conditional_formatting()
        return {"ok": True, "warnings": warnings}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def update_student_info(ctx: RPCContext):
    try:
        tdm.excel.student_info.update_student()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def add_student(ctx: RPCContext, target_student_name, target_class_name):
    try:
        if not tdm.aisosik.reader.check_student_exists(target_student_name, target_class_name):
            return {"ok": False, "error": f"아이소식에 {target_student_name} 학생이 {target_class_name} 반에 업데이트 되지 않아 중단되었습니다."}

        warnings = tdm.excel.data_file.add_student(target_student_name, target_class_name)

        tdm.excel.student_info.add_student(target_student_name)

        return {"ok": True, "warnings": warnings}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def remove_student(ctx: RPCContext, target_class_name, target_student_name):
    try:
        tdm.excel.data_file.delete_student(target_class_name, target_student_name)

        if not tdm.excel.data_file.check_student_exist(target_student_name):
            tdm.excel.student_info.delete_student(target_student_name)

        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def move_student(ctx: RPCContext, target_student_name, target_class_name, current_class_name):
    try:
        if not tdm.aisosik.reader.check_student_exists(target_student_name, target_class_name):
            return {"ok": False, "error": f"아이소식에 {target_student_name} 학생이 {target_class_name} 반에 업데이트 되지 않아 중단되었습니다."}

        tdm.excel.data_file.move_student(target_student_name, target_class_name, current_class_name)

        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def change_class_info(ctx: RPCContext, target_class_name, target_teacher_name):
    try:
        tdm.excel.class_info.change_class_info(target_class_name, target_teacher_name)

        tdm.excel.data_file.change_class_info(target_class_name, target_teacher_name)

        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def make_temp_class_info(ctx: RPCContext, new_class_list):
    try:
        filepath = tdm.excel.class_info.make_temp_file_for_update(new_class_list)
        return {"ok": True, "path": filepath}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def update_class(ctx: RPCContext):
    try:
        tdm.excel.data_file.update_class()
        tdm.excel.class_info.update_class()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}
    finally:
        try:
            tdm.excel.class_info.delete_temp()
        except:
            pass


@server.method()
async def delete_class_info_temp(ctx: RPCContext):
    try:
        tdm.excel.class_info.delete_temp()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def save_individual_result(ctx: RPCContext, student_name:str, class_name:str, test_name:str, target_row:int, target_col:int, test_score:int|float, makeup_test_check:bool, makeup_test_date:dict):
    try:
        job_id = str(uuid.uuid4())
        emit = make_emit(job_id)
        prog = Progress(emit, total=3)

        prog_warnings: list[str] = []
        _orig_warning = prog.warning

        def _capture_warning(msg: str):
            try:
                prog_warnings.append(str(msg))
            finally:
                # 원래 동작(실시간 이벤트 전송)도 유지
                _orig_warning(msg)

        prog.warning = _capture_warning  # type: ignore[attr-defined]

        for k, v in makeup_test_date.items():
            makeup_test_date[k] = datetime.strptime(v, "%Y-%m-%d")

        test_average = tdm.excel.data_file.save_individual_test_data(target_row, target_col, test_score)

        if test_score < 80 and not makeup_test_check:
            tdm.excel.makeup_test.save_individual_makeup_test(student_name, class_name, test_name, test_score, makeup_test_date, prog)

        tdm.aisosik.browser.send_individual_test_message(student_name, class_name, test_name, test_score, test_average, makeup_test_check, makeup_test_date, prog)

        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def save_retest_result(ctx: RPCContext, target_row:int, makeup_test_score:str):
    try:
        tdm.excel.makeup_test.save_makeup_test_result(target_row, makeup_test_score)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def change_data_file_name_by_select(ctx: RPCContext):
    try:
        selected_file = ctx.pyloid.open_file_dialog(f"{tdm.config.DATA_DIR}/data")
        if not selected_file:
            return {"ok": False}

        new_filename = Path(selected_file).stem

        tdm.config.change_data_file_name_by_select(new_filename)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def open_file_picker(ctx: RPCContext):
    try:
        selected_file = ctx.pyloid.open_file_dialog(tdm.config.DATA_DIR)
        if not selected_file:
            return {"ok": False}

        path_obj = Path(selected_file)
        if path_obj.suffix.lower() != ".xlsx":
            return {"ok": False, "error": "지원하지 않는 파일 형식입니다. .xlsx 파일만 선택해 주세요."}

        file_b64 = base64.b64encode(path_obj.read_bytes()).decode()

        return {"ok": True, "path": str(path_obj), "name": path_obj.name, "b64": file_b64}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}
