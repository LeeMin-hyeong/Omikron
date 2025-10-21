import base64
import json
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


####################################### config.json 관리 #######################################

def _open_path_cross_platform(path: str):
    p = os.path.abspath(path)
    if os.name == "nt":
        os.startfile(p)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", p])
    else:
        subprocess.Popen(["xdg-open", p])


def _read_config(cwd: Path) -> Dict[str, Any]:
    cfg_path = cwd / "config.json"
    if cfg_path.is_file():
        try:
            return json.loads(cfg_path.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _write_config(cwd: Path, cfg: Dict[str, Any]):
    (cwd / "config.json").write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def _ensure_parent(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)


####################################### 기존 작업 스레드 #######################################



####################################### 파일 작업 #######################################

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

####################################### 파일 작업 #######################################

def _send_exam_message_job(job_id: str, *, filename: str, b64: str) -> None:
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

        # TODO: 현재는 임시로 빈 dict 사용. 추후 UI에서 재시험 일정을 받도록 개선.
        makeup_test_date: Dict[str, Any] = {}

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


def _make_class_info_job(job_id: str):
    emit = make_emit(job_id)
    prog = Progress(emit, total=3)  # 대략 단계 수 추정

    try:
        omikron.classinfo.make_file()
        prog.done("반 정보 파일 생성 성공")
    except Exception as e:
        prog.error(f"예외 발생: {e}")


def _save_exam_job(job_id: str, *, filename: str, b64: str) -> None:
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

        # TODO: 현재는 임시로 빈 dict 사용. 추후 UI에서 재시험 일정을 받도록 개선.
        makeup_test_date: Dict[str, Any] = {}

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


def _make_data_file_job(job_id: str):
    emit = make_emit(job_id)
    prog = Progress(emit, total=3)  # 대략 단계 수 추정

    try:
        omikron.datafile.make_file()
        prog.done("데이터 파일 생성 성공")
    except Exception as e:
        prog.error(f"예외 발생: {e}")

####################################### 데이터 요청 API #######################################

@server.method()
async def check_data_files(ctx: RPCContext) -> Dict[str, Any]:
    """
    실행 디렉터리에 '반 정보.xlsx', '학생 정보.xlsx' 존재 여부와
    config.json의 dataFileName으로 './data/<name>.xlsx' 존재 여부 확인
    """
    cwd = Path(os.getcwd())
    class_info = cwd / "반 정보.xlsx"
    student_info = cwd / "학생 정보.xlsx"
    cfg = _read_config(cwd)
    data_file_name = str(cfg.get("dataFileName") or "").strip()
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
        "data_dir": str(cwd / "data"),
        "missing": missing,
    }


@server.method()
async def get_datafile_data(ctx: RPCContext) -> Dict[Any, Any]:
    return omikron.datafile.get_data_sorted_dict()


@server.method()
async def get_aisosic_data(ctx: RPCContext):
    return omikron.classinfo.check_updated_class()


@server.method()
async def get_makeuptest_data(ctx: RPCContext):
    return omikron.makeuptest.get_studnet_test_index_dict()


@server.method()
async def get_class_list(ctx: RPCContext):
    return omikron.classinfo.get_class_names()


@server.method()
async def get_class_info(ctx: RPCContext, class_name:str):
    return omikron.classinfo.get_class_info(class_name)

####################################### 작업 API #######################################

@server.method()
async def change_data_file_name(ctx:RPCContext, new_filename:str) -> Dict[str, Any]:
    omikron.config.change_data_file_name(new_filename)
    return {"ok": True}

@server.method()
async def start_make_class_info(ctx: RPCContext) -> Dict[str, Any]:
    """반 정보.xlsx 생성"""
    try:
        cwd = Path(os.getcwd())
        _ensure_parent(cwd)

        job_id = str(uuid.uuid4())
        t = threading.Thread(target=lambda: _make_class_info_job(job_id), daemon=True)
        t.start()
        return {"job_id": job_id}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def start_make_data_file(ctx: RPCContext) -> Dict[str, Any]:
    """
    데이터 파일 생성. name이 전달되면 config.json의 dataFileName을 갱신하고 생성
    (반 정보.xlsx 선행 필요)
    """
    try:
        cwd = Path(os.getcwd())
        class_info = cwd / "반 정보.xlsx"
        if not class_info.is_file():
            return {"ok": False, "error": "반 정보.xlsx가 먼저 필요합니다."}

        cfg = _read_config(cwd)

        if not (cfg.get("dataFileName") or "").strip():
            return {"ok": False, "error": "config.json의 dataFileName을 설정해 주세요."}

        _ensure_parent(cwd / "data")

        job_id = str(uuid.uuid4())
        t = threading.Thread(target=lambda: _make_data_file_job(job_id), daemon=True)
        t.start()
        return {"job_id": job_id}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def start_make_student_info(ctx: RPCContext) -> Dict[str, Any]:
    """학생 정보.xlsx 생성 (반 정보.xlsx 필요)"""
    try:
        cwd = Path(os.getcwd())
        class_info = cwd / "반 정보.xlsx"
        if not class_info.is_file():
            return {"ok": False, "error": "반 정보.xlsx가 먼저 필요합니다."}

        omikron.studentinfo.make_file()
        return {"ok": True, "path": str(cwd / '학생 정보.xlsx')}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def open_path(ctx: RPCContext, path: str) -> Dict[str, Any]:
    try:
        _open_path_cross_platform(path)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@server.method()
async def start_send_exam_message(ctx: RPCContext, filename: str, b64: str) -> Dict[str, Any]:
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
        kwargs={"job_id": job_id, "filename": filename, "b64": b64},
        daemon=True,
    )
    thread.start()

    return {"job_id": job_id}


@server.method()
async def start_save_exam(ctx: RPCContext, filename: str, b64: str) -> Dict[str, Any]:
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
        kwargs={"job_id": job_id, "filename": filename, "b64": b64},
        daemon=True,
    )
    thread.start()

    return {"job_id": job_id}


@server.method()
async def make_data_form(ctx: RPCContext):
    omikron.dataform.make_file()
    return {"ok": True}


@server.method()
async def reapply_conditional_format(ctx: RPCContext):
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

    omikron.datafile.conditional_formatting(prog)
    return {"ok": True}


@server.method()
async def update_student_info(ctx: RPCContext):
    omikron.studentinfo.update_student()
    return {"ok": True}
