import base64
from datetime import datetime
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import multiprocessing
from queue import Empty
import time
import traceback
import uuid
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.error import URLError, HTTPError
from urllib.request import Request, urlopen
import webbrowser
from bs4 import BeautifulSoup

from pyloid.rpc import PyloidRPC, RPCContext

from tdm.progress import Progress
import tdm.config

import tdm.classinfo
import tdm.chrome
import tdm.datafile
import tdm.dataform
import tdm.studentinfo
import tdm.makeuptest
from tdm.exception import NoMatchingSheetException, FileOpenException, ExcelRequiredException, ChromeDriverVersionMismatchException


####################################### 상태 관리 메서드 #######################################


server = PyloidRPC()

# 진행상태 저장소: job_id -> {step, status, message}
progress: dict[str, dict] = {}
job_threads: dict[str, threading.Thread] = {}
job_processes: dict[str, multiprocessing.Process] = {}
progress_queues: dict[str, multiprocessing.Queue] = {}
progress_listeners: dict[str, threading.Thread] = {}
job_process_started_at: dict[str, float] = {}
job_process_seen_payload: dict[str, bool] = {}


def make_emit(job_id: str):
    def _emit(payload: dict):
        prev = progress.get(job_id, {})
        warnings = list(prev.get("warnings", []))
        if payload.get("level") == "warning":
            msg = payload.get("message")
            if msg:
                msg_str = str(msg)
                if not warnings or warnings[-1] != msg_str:
                    warnings.append(msg_str)
        payload = {**payload, "warnings": warnings}
        progress[job_id] = payload
        # (옵션) 추후 실시간 브로드캐스트가 필요하면 이 지점에서 처리

    return _emit


def _queue_listener(job_id: str, q: multiprocessing.Queue, proc: multiprocessing.Process) -> None:
    while True:
        try:
            payload = q.get(timeout=0.5)
        except Empty:
            if not proc.is_alive():
                if proc.pid is None:
                    started_at = job_process_started_at.get(job_id, 0)
                    if started_at and (time.time() - started_at) < 5.0:
                        continue
                break
            continue
        if payload is None:
            break
        job_process_seen_payload[job_id] = True
        make_emit(job_id)(payload)
    progress_queues.pop(job_id, None)
    progress_listeners.pop(job_id, None)
    job_process_seen_payload.pop(job_id, None)
    job_process_started_at.pop(job_id, None)


@server.method()
async def get_progress(ctx: RPCContext, job_id: str) -> Dict[str, Any]:
    """진행상태 조회 (프런트 폴링)"""
    default_payload = {
        "step": 0,
        "total": 0,
        "level": "info",
        "status": "unknown",
        "message": "",
        "error": "",
        "detail": "",
        "warnings": [],
        "ts": time.time(),
    }
    payload = progress.get(job_id, default_payload)
    thread = job_threads.get(job_id)
    if thread and not thread.is_alive():
        status = payload.get("status")
        if status in ("running", "unknown"):
            payload = {
                **payload,
                "status": "done",
                "level": "success",
                "message": payload.get("message") or "작업이 완료되었습니다.",
                "ts": time.time(),
            }
            progress[job_id] = payload
        job_threads.pop(job_id, None)
    proc = job_processes.get(job_id)
    if proc and not proc.is_alive():
        status = payload.get("status")
        if status in ("running", "unknown"):
            started_at = job_process_started_at.get(job_id, 0)
            seen_payload = job_process_seen_payload.get(job_id, False)
            if not seen_payload and (time.time() - started_at) < 2.0:
                return payload
            if proc.exitcode not in (0, None):
                payload = {
                    **payload,
                    "status": "error",
                    "level": "error",
                    "message": payload.get("message") or "update_class process failed.",
                    "ts": time.time(),
                }
                progress[job_id] = payload
                job_processes.pop(job_id, None)
                return payload
            payload = {
                **payload,
                "status": "done",
                "level": "success",
                "message": payload.get("message") or "작업이 완료되었습니다.",
                "ts": time.time(),
            }
            progress[job_id] = payload
        job_processes.pop(job_id, None)

    return payload


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
    safe_name = Path(filename or "upload.bin").name
    if Path(safe_name).suffix.lower() != ".xlsx":
        raise ValueError("지원하지 않는 파일 형식입니다. .xlsx 파일만 사용할 수 있습니다.")

    tmp_root = Path(tempfile.mkdtemp(prefix="tdm_job_"))
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


def _validate_script_page_url(url: str) -> Optional[str]:
    """Validate that any <center> text equals 'Script'."""
    try:
        req = Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
            },
        )
        with urlopen(req, timeout=10) as resp:
            raw = resp.read()

        html = raw.decode("utf-8", errors="replace")
        soup = BeautifulSoup(html, "html.parser")
        marker_ok = any(c.get_text(strip=True) == "Script" for c in soup.select("center"))
        if not marker_ok:
            return "URL 페이지가 스크립트 페이지가 아닙니다. /html/body/div[1]/h1/center 값이 'Script'여야 합니다."
        return None
    except (HTTPError, URLError, TimeoutError):
        return "URL에 접속할 수 없습니다. 주소와 네트워크 상태를 확인해 주세요."
    except Exception:
        return "URL 검증 중 오류가 발생했습니다."


####################################### thread 작업 #######################################


def _update_class_job_process(job_id: str, q: multiprocessing.Queue) -> None:
    def _emit(payload: dict):
        q.put(payload)

    prog = Progress(_emit, total=5)

    prog.info("반 업데이트 준비중...")
    try:
        tdm.datafile.update_class(prog)
        prog.step("반 정보 파일 최신화 중...")
        tdm.classinfo.update_class(prog)
        prog.done("반 업데이트가 완료되었습니다.")
    except ExcelRequiredException as exc:
        prog.error(str(exc))
    except Exception:
        prog.error("예상치 못한 오류가 발생했습니다.", detail=traceback.format_exc())
    finally:
        tdm.classinfo.delete_temp()


def _send_exam_message_job_process(
    job_id: str,
    q: multiprocessing.Queue,
    *,
    filename: str,
    b64: str,
    makeup_test_date: Dict[str, Any],
) -> None:
    def _emit(payload: dict):
        q.put(payload)

    prog = Progress(_emit, total=3)

    _emit({
        "ts": time.time(),
        "step": 0,
        "total": 3,
        "level": "info",
        "status": "running",
        "message": "작업을 준비하고 있습니다.",
        "warnings": [],
    })

    tmp_file: Optional[Path] = None
    try:
        tmp_file = _decode_upload_to_temp(filename, b64)

        try:
            tdm.dataform.data_validation(str(tmp_file))
        except tdm.dataform.DataValidationException as exc:
            prog.error(f"데이터 검증 오류가 발생하였습니다:\n{exc}")
            return
        prog.step("데이터 입력 양식 검증 완료")

        for k, v in makeup_test_date.items():
            makeup_test_date[k] = datetime.strptime(v, "%Y-%m-%d")

        try:
            tdm.chrome.send_test_result_message(str(tmp_file), makeup_test_date, prog)
        except ChromeDriverVersionMismatchException as e:
            prog.error(str(e))
            return
        except Exception as e:
            prog.error(f"메시지 작성 중 오류가 발생했습니다:\n {e}")

        prog.step("작업 완료")

        prog.done("메시지 작성이 완료되었습니다. 전송 전 내용을 확인하세요.")
    except Exception:
        prog.error("예상치 못한 오류가 발생했습니다.", detail=traceback.format_exc())
    finally:
        if tmp_file:
            _cleanup_temp(tmp_file)


def _save_exam_job_process(
    job_id: str,
    q: multiprocessing.Queue,
    *,
    filename: str,
    b64: str,
    makeup_test_date: Dict[str, Any],
) -> None:
    def _emit(payload: dict):
        q.put(payload)

    prog = Progress(_emit, total=3)

    _emit({
        "ts": time.time(),
        "step": 0,
        "total": 4,
        "level": "info",
        "status": "running",
        "message": "작업을 준비하고 있습니다.",
        "warnings": [],
    })

    tmp_file: Optional[Path] = None
    try:
        tmp_file = _decode_upload_to_temp(filename, b64)

        try:
            tdm.dataform.data_validation(str(tmp_file))
        except tdm.dataform.DataValidationException as exc:
            prog.error(f"데이터 검증 오류가 발생하였습니다:\n {exc}")
            return
        prog.step("데이터 입력 양식 검증 완료")

        for k, v in makeup_test_date.items():
            makeup_test_date[k] = datetime.strptime(v, "%Y-%m-%d")

        try:
            datafile_wb = tdm.datafile.save_test_data(str(tmp_file), prog)
            makeuptest_wb = tdm.makeuptest.save_makeup_test_list(str(tmp_file), makeup_test_date, prog)
            prog.step("재시험 명단 입력 완료")
        except ExcelRequiredException as e:
            prog.error(str(e))
            return
        except NoMatchingSheetException as e:
            prog.error(f"파일에서 목표 시트를 찾을 수 없습니다:\n {e}")
            return
        except tdm.datafile.NoReservedColumnError as e:
            prog.error(f"파일에 필수 열이 없습니다:\n {e}")
            return

        try:
            tdm.datafile.save(datafile_wb)
            tdm.makeuptest.save(makeuptest_wb)
        except FileOpenException as e:
            prog.error(f"파일이 열려 있습니다:\n {e}")
            return

        prog.step("파일 저장 완료")

        prog.done("데이터 저장을 완료하였습니다.")
    except Exception:
        prog.error("예상치 못한 오류가 발생했습니다.", detail=traceback.format_exc())
        return
    finally:
        tdm.datafile.delete_temp()
        if tmp_file:
            _cleanup_temp(tmp_file)


####################################### 데이터 요청 API #######################################

@server.method()
async def check_data_files(ctx: RPCContext) -> Dict[str, Any]:
    """
    실행 디렉터리에 '반 정보.xlsx', '학생 정보.xlsx' 존재 여부와
    config.json의 dataFileName으로 './data/<name>.xlsx' 존재 여부 확인
    """
    cwd = Path(tdm.config.DATA_DIR)
    class_info = cwd / "반 정보.xlsx"
    student_info = cwd / "학생 정보.xlsx"
    data_file_name = tdm.config.DATA_FILE_NAME
    data_file = cwd / "data" / f"{data_file_name}.xlsx" if data_file_name else None

    has_class = class_info.is_file()
    has_student = student_info.is_file()
    has_data = bool(data_file_name) and data_file and data_file.is_file()
    data_dir_valid = tdm.config.DATA_DIR_VALID

    missing = []
    if not data_dir_valid:
        missing.append("데이터 저장 위치가 유효하지 않습니다.")
    if not has_class:
        missing.append("반 정보.xlsx")
    if not has_data:
        missing.append(f"data/{data_file_name}.xlsx")
    if not has_student:
        missing.append("학생 정보.xlsx")
    if not data_file_name:
        missing.append("config.json: dataFileName 설정 필요")

    ok = has_class and has_data and has_student and data_dir_valid
    return {
        "ok": ok,
        "has_class": has_class,
        "has_data": has_data,
        "has_student": has_student,
        "data_dir_valid": data_dir_valid,
        "data_file_name": data_file_name,
        "cwd": str(cwd),
        "data_dir": tdm.config.DATA_DIR,
        "missing": missing,
    }


@server.method()
async def get_config_status(ctx: RPCContext) -> Dict[str, Any]:
    return {
        "ok": True,
        "exists": tdm.config.CONFIG_EXISTS,
        "ready": tdm.config.is_initialized(),
        "termsAccepted": tdm.config.is_terms_accepted(),
        "config": {
            "url": tdm.config.URL,
            "dataDir": tdm.config.DATA_DIR if tdm.config.DATA_DIR_VALID else "",
            "dataFileName": tdm.config.DATA_FILE_NAME,
            "dailyTest": tdm.config.TEST_RESULT_MESSAGE,
            "makeupTest": tdm.config.MAKEUP_TEST_NO_SCHEDULE_MESSAGE,
            "makeupTestDate": tdm.config.MAKEUP_TEST_SCHEDULE_MESSAGE,
        },
    }


@server.method()
async def select_data_dir(ctx: RPCContext) -> Dict[str, Any]:
    try:
        selected = ctx.pyloid.select_directory_dialog(tdm.config.DATA_DIR or str(Path.cwd()))
        if not selected:
            return {"ok": False}
        return {"ok": True, "path": os.path.abspath(selected)}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def save_initial_config(
    ctx: RPCContext,
    url: str,
    data_dir: str,
    data_file_name: str,
    daily_test_message: str,
    makeup_test_message: str,
    makeup_test_date_message: str,
) -> Dict[str, Any]:
    try:
        url = (url or "").strip()
        data_dir = (data_dir or "").strip()
        data_file_name = (data_file_name or "").strip()
        daily_test_message = daily_test_message or ""
        makeup_test_message = makeup_test_message or ""
        makeup_test_date_message = makeup_test_date_message or ""

        if not url:
            return {"ok": False, "error": "아이소식 URL이 작성되지 않았습니다."}
        if not url.startswith(("http://", "https://")):
            return {"ok": False, "error": "URL이 유효하지 않습니다."}
        if not data_dir:
            return {"ok": False, "error": "데이터 저장 위치가 "}
        if not data_file_name:
            return {"ok": False, "error": "데이터 파일 이름이 지정되지 않았습니다."}
        if not daily_test_message.strip():
            return {"ok": False, "error": "테스트 결과 메시지 템플릿이 작성되지 않았습니다."}
        if not makeup_test_message.strip():
            return {"ok": False, "error": "재시험 안내 문구를 입력해 주세요."}
        if not makeup_test_date_message.strip():
            return {"ok": False, "error": "재시험 일정 안내 문구를 입력해 주세요.?"}

        tdm.config.initialize_config(
            url=url,
            data_dir=os.path.abspath(data_dir),
            data_file_name=data_file_name,
            daily_test_message=daily_test_message,
            makeup_test_message=makeup_test_message,
            makeup_test_date_message=makeup_test_date_message,
        )

        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def update_message_templates(
    ctx: RPCContext,
    url: str,
    daily_test_message: str,
    makeup_test_message: str,
    makeup_test_date_message: str,
) -> Dict[str, Any]:
    try:
        url = (url or "").strip()
        daily_test_message = daily_test_message or ""
        makeup_test_message = makeup_test_message or ""
        makeup_test_date_message = makeup_test_date_message or ""

        if not url:
            return {"ok": False, "error": "아이소식 URL이 작성되지 않았습니다."}
        if not url.startswith(("http://", "https://")):
            return {"ok": False, "error": "URL이 유효하지 않습니다."}
        if not daily_test_message.strip():
            return {"ok": False, "error": "테스트 결과 메시지 템플릿이 작성되지 않았습니다."}
        if not makeup_test_message.strip():
            return {"ok": False, "error": "재시험 안내 문구를 입력해 주세요."}
        if not makeup_test_date_message.strip():
            return {"ok": False, "error": "재시험 일정 안내 문구를 입력해 주세요."}

        tdm.config.update_message_templates(
            url=url,
            daily_test_message=daily_test_message,
            makeup_test_message=makeup_test_message,
            makeup_test_date_message=makeup_test_date_message,
        )
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def validate_script_url(ctx: RPCContext, url: str) -> Dict[str, Any]:
    try:
        url = (url or "").strip()
        if not url:
            return {"ok": False, "error": "아이소식 URL이 작성되지 않았습니다."}
        if not url.startswith(("http://", "https://")):
            return {"ok": False, "error": "URL이 유효하지 않습니다."}

        warning = _validate_script_page_url(url) is not None
        return {"ok": True, "warning": warning}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_terms_text(ctx: RPCContext) -> Dict[str, Any]:
    try:
        fallback_text = (
            "이용약관\n\n"
            "1) 본 프로그램은 등록된 라이선스 사용자만 사용할 수 있습니다.\n"
            "2) 사용자 데이터의 백업 및 보안 관리는 사용자 책임입니다.\n"
            "3) 프로그램의 무단 복제/배포/역공학을 금지합니다.\n"
            "4) 서비스 제공자는 시스템/네트워크/외부 서비스 이슈로 인한 장애를 보장하지 않습니다.\n"
            "5) 자세한 약관은 LICENSE 파일을 우선하며, 파일이 없을 경우 본 안내문이 적용됩니다."
        )

        license_path = Path("./LICENSE")
        if not license_path.exists():
            return {"ok": True, "title": "이용약관", "text": fallback_text}

        raw = license_path.read_bytes()
        text = None
        for enc in ("utf-8", "utf-8-sig", "cp949"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            text = raw.decode("latin-1", errors="replace")

        return {"ok": True, "title": "이용약관", "text": text}
    except Exception as e:
        return {"ok": True, "title": "이용약관", "text": fallback_text}


@server.method()
async def accept_terms(ctx: RPCContext) -> Dict[str, Any]:
    try:
        tdm.config.accept_terms()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_startup_notice(ctx: RPCContext) -> Dict[str, Any]:
    try:
        # Works without notice.json by using these defaults.
        raw: Dict[str, Any] = {
            "enabled": True,
            "source": "github_release",
            "repo": "LeeMin-hyeong/TestDataManagement",
            "include_prerelease": False,
            "title_prefix": "업데이트 안내",
        }

        notice_path = Path("./notice.json")
        if notice_path.exists():
            try:
                loaded = json.loads(notice_path.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    raw.update(loaded)
            except Exception:
                # Ignore invalid local override and keep defaults.
                pass

        enabled = bool(raw.get("enabled", True))
        if not enabled:
            return {"ok": True, "enabled": False, "title": "공지사항", "message": ""}

        source = str(raw.get("source", "static")).strip().lower()
        if source != "github_release":
            title = str(raw.get("title", "공지사항"))
            message = str(raw.get("message", "")).strip()
            notice_id = hashlib.sha256(f"{title}\n{message}".encode("utf-8")).hexdigest()
            seen_id = tdm.config.get_notice_seen_id()
            should_show = bool(message) and (notice_id != seen_id)
            return {
                "ok": True,
                "enabled": should_show,
                "title": title,
                "message": message,
                "noticeId": notice_id,
            }

        repo = str(raw.get("repo", "LeeMin-hyeong/TestDataManagement")).strip()
        include_prerelease = bool(raw.get("include_prerelease", False))
        if "/" not in repo:
            return {"ok": False, "error": "repo must be 'owner/repo' format."}

        if include_prerelease:
            api_url = f"https://api.github.com/repos/{repo}/releases?per_page=5"
        else:
            api_url = f"https://api.github.com/repos/{repo}/releases/latest"

        req = Request(
            api_url,
            headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": "tdm-notice-fetcher",
            },
        )
        try:
            with urlopen(req, timeout=8) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
        except (HTTPError, URLError) as e:
            return {"ok": False, "error": f"GitHub releases fetch failed: {e}"}

        release = None
        if isinstance(payload, dict):
            release = payload
        elif isinstance(payload, list):
            for item in payload:
                if not isinstance(item, dict):
                    continue
                if item.get("draft"):
                    continue
                if (not include_prerelease) and item.get("prerelease"):
                    continue
                release = item
                break

        if not isinstance(release, dict):
            return {"ok": True, "enabled": False, "title": "공지사항", "message": ""}

        name = str(release.get("name") or release.get("tag_name") or "Latest Release")
        body = str(release.get("body") or "").strip()
        url = str(release.get("html_url") or "")
        published = str(release.get("published_at") or "")
        notice_id = str(release.get("id") or release.get("tag_name") or url or name).strip()

        title_prefix = str(raw.get("title_prefix", "업데이트 안내")).strip() or "업데이트 안내"
        title = f"{title_prefix}: {name}"

        message_parts = []
        if published:
            message_parts.append(f"배포일: {published}")
        if body:
            message_parts.append(body)
        if url:
            message_parts.append(f"릴리즈 링크: {url}")
        message = "\n\n".join(message_parts).strip()

        seen_id = tdm.config.get_notice_seen_id()
        should_show = bool(message) and bool(notice_id) and (notice_id != seen_id)
        return {
            "ok": True,
            "enabled": should_show,
            "title": title,
            "message": message,
            "noticeId": notice_id,
        }
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def mark_notice_seen(ctx: RPCContext, notice_id: str) -> Dict[str, Any]:
    try:
        notice_id = (notice_id or "").strip()
        if not notice_id:
            return {"ok": False, "error": "notice_id is required."}
        tdm.config.set_notice_seen_id(notice_id)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def quit_app(ctx: RPCContext) -> Dict[str, Any]:
    try:
        ctx.pyloid.quit()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_startup_messages(ctx: RPCContext) -> Dict[str, Any]:
    return {
        "ok": True,
        "termsTitle": "?댁슜?쎄?",
        "termsMessage": (
            "蹂??꾨줈洹몃옩? ?깅줉???쇱씠?좎뒪 ?ъ슜?먮쭔 ?댁슜?????덉뒿?덈떎.\n"
            "?꾨줈洹몃옩 ?ъ슜?쇰줈 ?명븳 ?곗씠???먯떎??諛⑹??섍린 ?꾪빐 ?ъ슜 ??諛깆뾽??沅뚯옣?⑸땲??\n"
            "?숈썝 ?댁쁺 ?뺤콉 諛?愿??踰뺣졊??以?섑븯???ъ슜??二쇱꽭??"
        ),
        "noticeTitle": "怨듭??ы빆",
        "noticeMessage": "?꾩옱 ?깅줉??怨듭??ы빆???놁뒿?덈떎.",
    }


@server.method()
async def get_datafile_data(ctx: RPCContext, mocktest = False) -> Dict[Any, Any]:
    try:
        return {"ok": True, "data": tdm.datafile.get_data_sorted_dict(mocktest)}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_aisosic_data(ctx: RPCContext):
    try:
        return {"ok": True, "data": tdm.chrome.get_class_names()}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_aisosic_student_data(ctx: RPCContext):
    try:
        return {"ok": True, "data": tdm.chrome.get_class_student_dict()}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def check_aisosic_difference(ctx: RPCContext):
    try:
        aisosic = tdm.chrome.get_class_student_dict()
        datafile_raw = tdm.datafile.get_data_sorted_dict()
        if isinstance(datafile_raw, (list, tuple)) and len(datafile_raw) >= 1:
            datafile = datafile_raw[0]
        else:
            datafile = datafile_raw

        aisosic = aisosic or {}
        datafile = datafile or {}

        same = True
        for class_name, student_dict in datafile.items():
            datafile_students = set((student_dict or {}).keys())
            aisosic_students = set(aisosic.get(class_name) or [])
            if datafile_students != aisosic_students:
                same = False
                break
        return {"ok": True, "data": same}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_makeuptest_data(ctx: RPCContext):
    try:
        return {"ok": True, "data": tdm.makeuptest.get_studnet_test_index_dict()}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_class_list(ctx: RPCContext):
    try:
        return {"ok": True, "data": tdm.classinfo.get_class_names()}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_class_info(ctx: RPCContext, class_name:str):
    try:
        return {"ok": True, "data": tdm.classinfo.get_class_info(class_name)}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def get_new_class_list(ctx: RPCContext):
    try:
        return {"ok": True, "data": tdm.classinfo.get_new_class_names()}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def is_cell_empty(ctx: RPCContext, row:int, col:int):
    try:
        empty, value = tdm.datafile.is_cell_empty(row, col)
        return {"ok": True, "empty": empty, "value": value}
    except Exception as e:
            return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


####################################### 작업 API #######################################

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
async def start_send_exam_message(ctx: RPCContext, filename: str, b64: str, makeup_test_date: Dict[str, Any]) -> Dict[str, Any]:
    job_id = str(uuid.uuid4())

    make_emit(job_id)({
        "ts": time.time(),
        "step": 0,
        "total": 3,
        "level": "info",
        "status": "running",
        "message": "작업 대기 중...",
        "warnings": [],
    })

    ctx_mp = multiprocessing.get_context("spawn")
    q = ctx_mp.Queue()
    proc = ctx_mp.Process(
        target=_send_exam_message_job_process,
        kwargs={
            "job_id": job_id,
            "q": q,
            "filename": filename,
            "b64": b64,
            "makeup_test_date": makeup_test_date,
        },
        daemon=True,
    )
    progress_queues[job_id] = q
    job_processes[job_id] = proc
    job_process_started_at[job_id] = time.time()
    job_process_seen_payload[job_id] = False
    listener = threading.Thread(
        target=_queue_listener,
        args=(job_id, q, proc),
        daemon=True,
    )
    progress_listeners[job_id] = listener
    listener.start()
    try:
        proc.start()
    except Exception:
        make_emit(job_id)({
            "ts": time.time(),
            "step": 0,
            "total": 0,
            "level": "error",
            "status": "error",
            "message": "send_exam_message process failed to start.",
            "warnings": [],
        })

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
        "warnings": [],
    })

    ctx_mp = multiprocessing.get_context("spawn")
    q = ctx_mp.Queue()
    proc = ctx_mp.Process(
        target=_save_exam_job_process,
        kwargs={
            "job_id": job_id,
            "q": q,
            "filename": filename,
            "b64": b64,
            "makeup_test_date": makeup_test_date,
        },
        daemon=True,
    )
    progress_queues[job_id] = q
    job_processes[job_id] = proc
    job_process_started_at[job_id] = time.time()
    job_process_seen_payload[job_id] = False
    listener = threading.Thread(
        target=_queue_listener,
        args=(job_id, q, proc),
        daemon=True,
    )
    progress_listeners[job_id] = listener
    listener.start()
    try:
        proc.start()
    except Exception:
        make_emit(job_id)({
            "ts": time.time(),
            "step": 0,
            "total": 0,
            "level": "error",
            "status": "error",
            "message": "save_exam process failed to start.",
            "warnings": [],
        })

    return {"job_id": job_id}


@server.method()
async def start_update_class(ctx: RPCContext) -> Dict[str, Any]:
    job_id = str(uuid.uuid4())

    make_emit(job_id)({
        "ts": time.time(),
        "step": 0,
        "total": 6,
        "level": "info",
        "status": "running",
        "message": "반 업데이트 준비중...",
        "warnings": [],
    })

    ctx_mp = multiprocessing.get_context("spawn")
    q = ctx_mp.Queue()
    proc = ctx_mp.Process(
        target=_update_class_job_process,
        kwargs={"job_id": job_id, "q": q},
        daemon=True,
    )
    progress_queues[job_id] = q
    job_processes[job_id] = proc
    job_process_started_at[job_id] = time.time()
    job_process_seen_payload[job_id] = False
    listener = threading.Thread(
        target=_queue_listener,
        args=(job_id, q, proc),
        daemon=True,
    )
    progress_listeners[job_id] = listener
    listener.start()
    try:
        proc.start()
    except Exception:
        make_emit(job_id)({
            "ts": time.time(),
            "step": 0,
            "total": 0,
            "level": "error",
            "status": "error",
            "message": "update_class process failed to start.",
            "warnings": [],
        })

    return {"job_id": job_id}


@server.method()
async def make_class_info(ctx: RPCContext):
    try:
        tdm.classinfo.make_file()
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

        tdm.datafile.make_file()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def make_student_info(ctx: RPCContext):
    try:
        tdm.studentinfo.make_file()
        return {"ok": True, "path": str(Path(tdm.config.DATA_DIR) / '학생 정보.xlsx')}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def make_data_form(ctx: RPCContext):
    try:
        tdm.dataform.make_file()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def reapply_conditional_format(ctx: RPCContext):
    try:
        warnings = tdm.datafile.conditional_formatting()
        return {"ok": True, "warnings": warnings}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def update_student_info(ctx: RPCContext):
    try:
        tdm.studentinfo.update_student()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def add_student(ctx: RPCContext, target_student_name, target_class_name):
    try:
        if not tdm.chrome.check_student_exists(target_student_name, target_class_name):
            return {"ok": False, "error": f"아이소식에 {target_student_name} 학생이 {target_class_name} 반에 업데이트 되지 않아 중단되었습니다."}

        warnings = tdm.datafile.add_student(target_student_name, target_class_name)

        tdm.studentinfo.add_student(target_student_name)

        return {"ok": True, "warnings": warnings}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def remove_student(ctx: RPCContext, target_class_name, target_student_name):
    try:
        tdm.datafile.delete_student(target_class_name, target_student_name)

        if not tdm.datafile.check_student_exist(target_student_name):
            tdm.studentinfo.delete_student(target_student_name)

        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def move_student(ctx: RPCContext, target_student_name, target_class_name, current_class_name):
    try:
        if not tdm.chrome.check_student_exists(target_student_name, target_class_name):
            return {"ok": False, "error": f"아이소식에 {target_student_name} 학생이 {target_class_name} 반에 업데이트 되지 않아 중단되었습니다."}

        tdm.datafile.move_student(target_student_name, target_class_name, current_class_name)

        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def change_class_info(ctx: RPCContext, target_class_name, target_teacher_name):
    try:
        tdm.classinfo.change_class_info(target_class_name, target_teacher_name)

        tdm.datafile.change_class_info(target_class_name, target_teacher_name)

        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def make_temp_class_info(ctx: RPCContext, new_class_list):
    try:
        filepath = tdm.classinfo.make_temp_file_for_update(new_class_list)
        return {"ok": True, "path": filepath}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def update_class(ctx: RPCContext):
    try:
        tdm.datafile.update_class()
        tdm.classinfo.update_class()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}
    finally:
        try:
            tdm.classinfo.delete_temp()
        except:
            pass


@server.method()
async def delete_class_info_temp(ctx: RPCContext):
    try:
        tdm.classinfo.delete_temp()
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

        test_average = tdm.datafile.save_individual_test_data(target_row, target_col, test_score)

        if test_score < 80 and not makeup_test_check:
            tdm.makeuptest.save_individual_makeup_test(student_name, class_name, test_name, test_score, makeup_test_date, prog)

        tdm.chrome.send_individual_test_message(student_name, class_name, test_name, test_score, test_average, makeup_test_check, makeup_test_date, prog)

        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e), "detail": traceback.format_exc()}


@server.method()
async def save_retest_result(ctx: RPCContext, target_row:int, makeup_test_score:str):
    try:
        tdm.makeuptest.save_makeup_test_result(target_row, makeup_test_score)
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
