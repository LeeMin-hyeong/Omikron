import base64
from datetime import datetime
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
from pathlib import Path
from typing import Any, Dict, Optional
import webbrowser

from pyloid.rpc import PyloidRPC, RPCContext

from omikron.progress import Progress
import omikron.config

import omikron.classinfo
import omikron.chrome
import omikron.datafile
import omikron.dataform
import omikron.studentinfo
import omikron.makeuptest
from omikron.exception import NoMatchingSheetException, FileOpenException


####################################### 상태 관리 메서드 #######################################


server = PyloidRPC()

# 진행상태 저장소: job_id -> {step, status, message}
progress: dict[str, dict] = {}


def make_emit(job_id: str):
    def _emit(payload: dict):
        progress[job_id] = payload
        # (옵션) 추후 실시간 브로드캐스트가 필요하면 이 지점에서 처리

    return _emit


@server.method()
async def get_progress(ctx: RPCContext, job_id: str) -> Dict[str, Any]:
    """진행상태 조회 (프런트 폴링)"""
    default_payload = {
        "step": 0,
        "total": 0,
        "level": "info",
        "status": "unknown",
        "message": "",
        "ts": time.time(),
    }
    return progress.get(job_id, default_payload)


####################################### 파일 열기 #######################################


def _open_path_cross_platform(path: str):
    p = os.path.abspath(path)
    if os.name == "nt":
        os.startfile(p)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", p])
    else:
        subprocess.Popen(["xdg-open", p])


####################################### 임시 파일 #######################################


def _decode_upload_to_temp(filename: str, b64: str) -> Path:
    """업로드된 base64 데이터를 임시 파일로 저장"""
    tmp_root = Path(tempfile.mkdtemp(prefix="omikron_job_"))
    safe_name = Path(filename or "upload.bin").name
    tmp_path = tmp_root / safe_name
    try:
        data = base64.b64decode(b64)
        tmp_path.write_bytes(data)
        return tmp_path
    except Exception:
        shutil.rmtree(tmp_root, ignore_errors=True)
        raise


def _cleanup_temp(path: Path) -> None:
    """임시 파일/폴더 정리"""
    try:
        root = path if path.is_dir() else path.parent
        if path.is_file():
            try:
                path.unlink()
            except Exception:
                pass
        shutil.rmtree(root, ignore_errors=True)
    except Exception:
        pass


####################################### thread 작업 #######################################


def _send_exam_message_job(job_id: str, *, filename: str, b64: str, makeup_test_date: Dict[str, Any]) -> None:
    """Chrome 자동화를 통해 시험 결과 메시지 작성"""
    emit = make_emit(job_id)
    prog = Progress(emit, total=3)

    prog.info("작업을 준비하고 있습니다.")

    tmp_file: Optional[Path] = None
    try:
        tmp_file = _decode_upload_to_temp(filename, b64)

        try:
            omikron.dataform.data_validation(str(tmp_file))
        except omikron.dataform.DataValidationException as exc:
            prog.error(f"데이터 검증 오류: {exc}")
            return
        prog.step("데이터 입력 양식 검증 완료")

        for k, v in makeup_test_date.items():
            makeup_test_date[k] = datetime.strptime(v, "%Y-%m-%d")

        ok = omikron.chrome.send_test_result_message(str(tmp_file), makeup_test_date, prog)
        if not ok:
            prog.error("메시지 작성 중 오류가 발생했습니다.")
            return
        prog.step("Chrome 자동화를 실행하여 메시지를 작성합니다.")

        prog.step("작업 완료")

        prog.done("메시지 작성이 완료되었습니다. 전송 전 내용을 확인하세요.")
    except Exception as exc:
        prog.error(f"예상치 못한 오류가 발생했습니다: {exc}")
    finally:
        if tmp_file:
            _cleanup_temp(tmp_file)


def _save_exam_job(job_id: str, *, filename: str, b64: str, makeup_test_date: Dict[str, Any]) -> None:
    emit = make_emit(job_id)
    prog = Progress(emit, total=3)

    tmp_file: Optional[Path] = None
    try:
        tmp_file = _decode_upload_to_temp(filename, b64)

        try:
            omikron.dataform.data_validation(str(tmp_file))
        except omikron.dataform.DataValidationException as exc:
            prog.error(f"데이터 검증 오류가 발생하였습니다:\n {exc}")
            return
        prog.step("데이터 입력 양식 검증 완료")

        for k, v in makeup_test_date.items():
            makeup_test_date[k] = datetime.strptime(v, "%Y-%m-%d")

        try:
            datafile_wb = omikron.datafile.save_test_data(str(tmp_file), prog)
            makeuptest_wb = omikron.makeuptest.save_makeup_test_list(str(tmp_file), makeup_test_date, prog)
            prog.step("재시험 명단 입력 완료")
        except NoMatchingSheetException as e:
            prog.error(f"파일에서 목표 시트를 찾을 수 없습니다:\n {e}")
            return
        except omikron.datafile.NoReservedColumnError as e:
            prog.error(f"파일에 필수 열이 없습니다:\n {e}")
            return

        try:
            omikron.datafile.save(datafile_wb)
            omikron.makeuptest.save(makeuptest_wb)
        except FileOpenException as e:
            prog.error(f"파일이 열려 있습니다:\n {e}")
            return

        prog.step("파일 저장 완료")

        prog.done("데이터 저장을 완료하였습니다.")
        omikron.datafile.delete_temp()
    except Exception as exc:
        prog.error(f"예상치 못한 오류가 발생했습니다:\n {exc}")
        return
    finally:
        if tmp_file:
            _cleanup_temp(tmp_file)


####################################### 데이터 요청 API #######################################

@server.method()
async def check_data_files(ctx: RPCContext) -> Dict[str, Any]:
    """
    실행 디렉터리에 '반 정보.xlsx', '학생 정보.xlsx' 존재 여부와
    config.json의 dataFileName으로 './data/<name>.xlsx' 존재 여부 확인
    """
    cwd = Path(omikron.config.DATA_DIR)
    class_info = cwd / "반 정보.xlsx"
    student_info = cwd / "학생 정보.xlsx"
    data_file_name = omikron.config.DATA_FILE_NAME
    data_file = cwd / "data" / f"{data_file_name}.xlsx" if data_file_name else None

    has_class = class_info.is_file()
    has_student = student_info.is_file()
    has_data = bool(data_file_name) and data_file and data_file.is_file()

    missing = []
    if not has_class:
        missing.append("반 정보.xlsx")
    if not has_data:
        missing.append(f"data/{data_file_name}.xlsx")
    if not has_student:
        missing.append("학생 정보.xlsx")
    if not data_file_name:
        missing.append("config.json: dataFileName 설정 필요")

    ok = has_class and has_data and has_student
    return {
        "ok": ok,
        "has_class": has_class,
        "has_data": has_data,
        "has_student": has_student,
        "data_file_name": data_file_name,
        "cwd": str(cwd),
        "data_dir": omikron.config.DATA_DIR,
        "missing": missing,
    }


@server.method()
async def get_datafile_data(ctx: RPCContext, mocktest = False) -> Dict[Any, Any]:
    try:
        return {"ok": True, "data": omikron.datafile.get_data_sorted_dict(mocktest)}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def get_aisosic_data(ctx: RPCContext):
    try:
        return {"ok": True, "data": omikron.chrome.get_class_names()}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def get_makeuptest_data(ctx: RPCContext):
    try:
        return {"ok": True, "data": omikron.makeuptest.get_studnet_test_index_dict()}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def get_class_list(ctx: RPCContext):
    try:
        return {"ok": True, "data": omikron.classinfo.get_class_names()}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def get_class_info(ctx: RPCContext, class_name:str):
    try:
        return {"ok": True, "data": omikron.classinfo.get_class_info(class_name)}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def get_new_class_list(ctx: RPCContext):
    try:
        return {"ok": True, "data": omikron.classinfo.get_new_class_names()}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def is_cell_empty(ctx: RPCContext, row:int, col:int):
    try:
        empty, value = omikron.datafile.is_cell_empty(row, col)
        return {"ok": True, "empty": empty, "value": value}
    except Exception as e:
            return {"ok": False, "error": str(e)}


####################################### 작업 API #######################################

@server.method()
async def change_data_dir(ctx:RPCContext):
    try:
        new_dir = ctx.pyloid.select_directory_dialog(omikron.config.DATA_DIR)
        if new_dir is None: return {"ok": False}
        abspath = os.path.abspath(new_dir)
        omikron.config.change_data_path(abspath)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def change_data_file_name(ctx:RPCContext, new_filename:str) -> Dict[str, Any]:
    try:
        omikron.config.change_data_file_name(new_filename)
        return {"ok": True}
    except FileOpenException as e:
        return {"ok": False, "error": str(e)}
    except Exception as e:
        return {"ok": False, "error": f"알 수 없는 에러가 발생하였습니다: {str(e)}"}


@server.method()
async def open_path(ctx: RPCContext, path: str) -> Dict[str, Any]:
    try:
        _open_path_cross_platform(path)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": f"알 수 없는 에러가 발생하였습니다: {str(e)}"}


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
        return {"ok": False, "error": str(e)}


@server.method()
async def start_send_exam_message(ctx: RPCContext, filename: str, b64: str, makeup_test_date: Dict[str, Any]) -> Dict[str, Any]:
    job_id = str(uuid.uuid4())

    make_emit(job_id)({
        "ts": time.time(),
        "step": 0,
        "total": 3,
        "level": "info",
        "status": "running",
        "message": "작업 대기 중...",
    })

    thread = threading.Thread(
        target=_send_exam_message_job,
        kwargs={"job_id": job_id, "filename": filename, "b64": b64, "makeup_test_date": makeup_test_date},
        daemon=True,
    )
    thread.start()

    return {"job_id": job_id}


@server.method()
async def start_save_exam(ctx: RPCContext, filename: str, b64: str, makeup_test_date: Dict[str, Any]) -> Dict[str, Any]:
    job_id = str(uuid.uuid4())

    make_emit(job_id)({
        "ts": time.time(),
        "step": 0,
        "total": 4,
        "level": "info",
        "status": "running",
        "message": "작업 대기 중...",
    })

    thread = threading.Thread(
        target=_save_exam_job,
        kwargs={"job_id": job_id, "filename": filename, "b64": b64, "makeup_test_date": makeup_test_date},
        daemon=True,
    )
    thread.start()

    return {"job_id": job_id}


@server.method()
async def make_class_info(ctx: RPCContext):
    try:
        omikron.classinfo.make_file()
        return {"ok": True, "path": str(Path(omikron.config.DATA_DIR) / '반 정보.xlsx')}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def make_data_file(ctx: RPCContext):
    try:
        cwd = Path(omikron.config.DATA_DIR)
        class_info = cwd / "반 정보.xlsx"
        if not class_info.is_file():
            return {"ok": False, "error": "반 정보.xlsx가 먼저 필요합니다."}

        if not omikron.config.DATA_FILE_NAME:
            return {"ok": False, "error": "config.json의 dataFileName을 설정해 주세요."}

        omikron.datafile.make_file()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def make_student_info(ctx: RPCContext):
    try:
        omikron.studentinfo.make_file()
        return {"ok": True, "path": str(Path(omikron.config.DATA_DIR) / '학생 정보.xlsx')}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def make_data_form(ctx: RPCContext):
    try:
        omikron.dataform.make_file()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def reapply_conditional_format(ctx: RPCContext):
    try:
        warnings = omikron.datafile.conditional_formatting()
        return {"ok": True, "warnings": warnings}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def update_student_info(ctx: RPCContext):
    try:
        omikron.studentinfo.update_student()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def add_student(ctx: RPCContext, target_student_name, target_class_name):
    try:
        if not omikron.chrome.check_student_exists(target_student_name, target_class_name):
            return {"ok": False, "error": f"아이소식에 {target_student_name} 학생이 {target_class_name} 반에 업데이트 되지 않아 중단되었습니다."}

        warnings = omikron.datafile.add_student(target_student_name, target_class_name)

        omikron.studentinfo.add_student(target_student_name)

        return {"ok": True, "warnings": warnings}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def remove_student(ctx: RPCContext, target_student_name):
    try:
        omikron.datafile.delete_student(target_student_name)

        omikron.studentinfo.delete_student(target_student_name)

        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def move_student(ctx: RPCContext, target_student_name, target_class_name, current_class_name):
    try:
        if not omikron.chrome.check_student_exists(target_student_name, target_class_name):
            return {"ok": False, "error": f"아이소식에 {target_student_name} 학생이 {target_class_name} 반에 업데이트 되지 않아 중단되었습니다."}

        omikron.datafile.move_student(target_student_name, target_class_name, current_class_name)

        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def change_class_info(ctx: RPCContext, target_class_name, target_teacher_name):
    try:
        omikron.classinfo.change_class_info(target_class_name, target_teacher_name)

        omikron.datafile.change_class_info(target_class_name, target_teacher_name)

        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def make_temp_class_info(ctx: RPCContext, new_class_list):
    try:
        filepath = omikron.classinfo.make_temp_file_for_update(new_class_list)
        return {"ok": True, "path": filepath}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def update_class(ctx: RPCContext):
    try:
        omikron.datafile.update_class()
        omikron.classinfo.update_class()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def delete_class_info_temp(ctx: RPCContext):
    try:
        omikron.classinfo.delete_temp()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


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

        test_average = omikron.datafile.save_individual_test_data(target_row, target_col, test_score)

        if test_score < 80 and not makeup_test_check:
            omikron.makeuptest.save_individual_makeup_test(student_name, class_name, test_name, test_score, makeup_test_date, prog)

        omikron.chrome.send_individual_test_message(student_name, class_name, test_name, test_score, test_average, makeup_test_check, makeup_test_date, prog)

        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def save_retest_result(ctx: RPCContext, target_row:int, makeup_test_score:str):
    try:
        omikron.makeuptest.save_makeup_test_result(target_row, makeup_test_score)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def change_data_file_name_by_select(ctx: RPCContext):
    try:
        selected_file = ctx.pyloid.open_file_dialog(f"{omikron.config.DATA_DIR}/data")
        if not selected_file:
            return {"ok": False}

        new_filename = Path(selected_file).stem

        omikron.config.change_data_file_name_by_select(new_filename)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def open_file_picker(ctx: RPCContext):
    try:
        selected_file = ctx.pyloid.open_file_dialog(omikron.config.DATA_DIR)
        if not selected_file:
            return {"ok": False}

        path_obj = Path(selected_file)
        file_b64 = base64.b64encode(path_obj.read_bytes()).decode()

        return {"ok": True, "path": str(path_obj), "name": path_obj.name, "b64": file_b64}
    except Exception as e:
        return {"ok": False, "error": str(e)}
