import json
from pathlib import Path
import time
from typing import Any, Dict, Optional
import os, base64, tempfile, threading, uuid
import openpyxl
from pyloid.rpc import PyloidRPC, RPCContext
from omikron.datafile import save_test_data
import omikron.dataform
import omikron.classinfo
import omikron.datafile
import omikron.studentinfo
import omikron.chrome
from omikron.progress import Progress
import sys

####################################### 상태 관리 및 RPC 메서드 ####################################### 

server = PyloidRPC()

# 진행상태 저장소: job_id -> {step, status, message}
progress: dict[str, dict] = {}

def make_emit(job_id: str):
    def _emit(payload: dict):
        progress[job_id] = payload
        # (옵션) 즉시 푸시도 함께:
        # try: ctx.pyloid.emit("omikron:progress", { "job_id": job_id, **payload })
        # except: pass
    return _emit

@server.method()
async def get_progress(ctx: RPCContext, job_id: str) -> Dict[str, Any]:
    """진행상태 조회 (프런트가 주기적으로 호출)"""
    # 표준 포맷: step, total, level, status, message, ts
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

####################################### tasks #######################################

def _run_save_exam(job_id: str, filename: str, b64: str):
    try:
        reporter = Progress(emit=make_emit(job_id), total=8)  # 대략 단계 수

        suffix = os.path.splitext(filename)[1] or ".xlsx"
        fd, temp_path = tempfile.mkstemp(prefix="omk_", suffix=suffix)
        os.close(fd)
        with open(temp_path, "wb") as f:
            f.write(base64.b64decode(b64))

        if not omikron.dataform.data_validation(temp_path):
            return

        reporter.info("업로드 수신 완료", inc=True)
        ok, _wb = save_test_data(temp_path, progress=reporter)
        # if not ok:
        #     # save_test_data 내부에서 이미 error를 보냈을 수 있지만, 보장용:
        #     reporter.error("처리 실패")
        #     return
        reporter.done("작업 성공")
    except Exception as e:
        reporter.error(f"예외 발생: {e}")

def _run_make_data_file(job_id: str):
    try:
        reporter = Progress(emit=make_emit(job_id), total=6)  # 대략 단계 수
        omikron.datafile.make_file(progress=reporter)
        reporter.done("데이터 파일 생성 성공")
    except Exception as e:
        reporter.error(f"예외 발생: {e}")

def _run_make_class_info(job_id: str):
    try:
        reporter = Progress(emit=make_emit(job_id), total=1)  # 대략 단계 수
        omikron.classinfo.make_file(progress=reporter)
        reporter.done("반 정보 파일 생성 성공")
    except Exception as e:
        reporter.error(f"예외 발생: {e}")

####################################### api calls #######################################

@server.method()
async def start_save_exam(ctx: RPCContext, filename: str, b64: str) -> Dict[str, Any]:
    job_id = str(uuid.uuid4())
    t = threading.Thread(target=_run_save_exam, args=(job_id, filename, b64), daemon=True)
    t.start()
    return {"job_id": job_id}

@server.method()
async def check_data_files(ctx: RPCContext) -> Dict[str, Any]:
    """
    실행 디렉터리에 '반 정보.xlsx', '학생 정보.xlsx' 존재 여부와
    config.json의 dataFileName으로 './data/<name>.xlsx' 존재 여부를 확인.
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
async def start_make_class_info(ctx: RPCContext) -> Dict[str, Any]:
    """반 정보.xlsx 생성"""
    try:
        path = Path(os.getcwd()) / "반 정보.xlsx"
        omikron.classinfo.make_file()
        return {"ok": True, "path": str(path)}
    except Exception as e:
        return {"ok": False, "error": str(e)}

@server.method()
async def start_make_data_file(ctx: RPCContext, name: Optional[str] = None) -> dict:
    """
    데이터파일 생성. name이 전달되면 config.json의 dataFileName을 갱신하고 생성.
    (반 정보.xlsx 선행 필요)
    """
    try:
        cwd = Path(os.getcwd())
        class_info = cwd / "반 정보.xlsx"
        if not class_info.is_file():
            return {"ok": False, "error": "반 정보.xlsx가 먼저 필요합니다."}

        cfg = _read_config(cwd)

        if name and name.strip():
            cfg["dataFileName"] = name.strip()
            _write_config(cwd, cfg)

        final_name = (cfg.get("dataFileName") or "").strip()
        if not final_name:
            return {"ok": False, "error": "config.json의 dataFileName을 설정해 주세요."}

        data_path = cwd / "data" / f"{final_name}.xlsx"

        job_id = str(uuid.uuid4())
        t = threading.Thread(target=lambda: _run_make_data_file(job_id), daemon=True)
        t.start()
        return {"job_id": job_id}
    except Exception as e:
        return {"ok": False, "error": str(e)}

@server.method()
async def start_make_student_info(ctx: RPCContext) -> Dict[str, Any]:
    """학생 정보.xlsx 생성 (선행: 반 정보.xlsx 존재)"""
    try:
        cwd = Path(os.getcwd())
        class_info = cwd / "반 정보.xlsx"
        if not class_info.is_file():
            return {"ok": False, "error": "반 정보.xlsx가 먼저 필요합니다."}

        omikron.studentinfo.make_file()
        return {"ok": True, "path": str("")}
    except Exception as e:
        return {"ok": False, "error": str(e)}

@server.method()
async def open_path(ctx: RPCContext, path: str) -> Dict[str, Any]:
    try:
        _open_path_cross_platform(path)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}