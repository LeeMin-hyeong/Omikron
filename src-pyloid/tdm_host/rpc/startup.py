"""Startup, configuration, notice, and terms RPC handlers."""

import hashlib
import json
import os
from pathlib import Path
from typing import Any, Dict
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from pyloid.rpc import RPCContext

import tdm.config
from tdm.excel.paths import WorkbookPaths
from tdm.excel.storage_health import verify_directory_writable
from tdm.excel.transaction import (
    cleanup_stale_staging_files,
    recover_pending_transactions,
)
from tdm_host.rpc.error_codes import RpcErrorCode
from tdm_host.rpc.responses import error_response, failure_response, success_response
from tdm_host.rpc.transport import server
from tdm_host.runtime.validation import validate_script_page_url as _validate_script_page_url

@server.method()
async def check_data_files(ctx: RPCContext) -> Dict[str, Any]:
    """
    실행 디렉터리에 '반 정보.xlsx', '학생 정보.xlsx' 존재 여부와
    config.json의 dataFileName으로 './data/<name>.xlsx' 존재 여부 확인
    """
    cwd = Path(tdm.config.DATA_DIR)
    paths = WorkbookPaths.current()
    class_info = paths.class_info
    student_info = paths.student_info
    data_file_name = tdm.config.DATA_FILE_NAME
    data_file = paths.data_file if data_file_name else None

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
        "recovery_actions": [],
        "storage_error": "",
    }


@server.method()
async def verify_storage_health(ctx: RPCContext) -> Dict[str, Any]:
    """Run write verification and transaction recovery once during startup."""
    if not tdm.config.DATA_DIR_VALID:
        return error_response(RpcErrorCode.DATA_DIR_INVALID, "데이터 저장 위치가 유효하지 않습니다.")
    try:
        paths = WorkbookPaths.current()
        verify_directory_writable(paths.data_dir)
        recovery_actions = [
            {"transactionId": result.transaction_id, "action": result.action}
            for result in recover_pending_transactions(paths)
        ]
        cleanup_stale_staging_files(paths.root)
        return success_response(recovery_actions=recovery_actions)
    except Exception as exc:
        return failure_response(exc, context="verify_storage_health")


@server.method()
async def get_config_status(ctx: RPCContext) -> Dict[str, Any]:
    return success_response(data={
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
    })


@server.method()
async def select_data_dir(ctx: RPCContext) -> Dict[str, Any]:
    try:
        selected = ctx.pyloid.select_directory_dialog(tdm.config.DATA_DIR or str(Path.cwd()))
        if not selected:
            return error_response(RpcErrorCode.CANCELLED, "폴더 선택을 취소했습니다.")
        return success_response(path=os.path.abspath(selected))
    except Exception as exc:
        return failure_response(exc, context="select_data_dir")


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
            return error_response(RpcErrorCode.URL_REQUIRED, "아이소식 URL이 작성되지 않았습니다.")
        if not url.startswith(("http://", "https://")):
            return error_response(RpcErrorCode.URL_INVALID, "URL이 유효하지 않습니다.")
        if not data_dir:
            return error_response(RpcErrorCode.DATA_DIR_REQUIRED, "데이터 저장 위치를 선택해 주세요.")
        if not data_file_name:
            return error_response(RpcErrorCode.DATA_FILE_NAME_REQUIRED, "데이터 파일 이름이 지정되지 않았습니다.")
        if not daily_test_message.strip():
            return error_response(RpcErrorCode.DAILY_MESSAGE_REQUIRED, "테스트 결과 메시지 템플릿이 작성되지 않았습니다.")
        if not makeup_test_message.strip():
            return error_response(RpcErrorCode.MAKEUP_MESSAGE_REQUIRED, "재시험 안내 문구를 입력해 주세요.")
        if not makeup_test_date_message.strip():
            return error_response(RpcErrorCode.MAKEUP_DATE_MESSAGE_REQUIRED, "재시험 일정 안내 문구를 입력해 주세요.")

        verify_directory_writable(Path(data_dir))

        tdm.config.initialize_config(
            url=url,
            data_dir=os.path.abspath(data_dir),
            data_file_name=data_file_name,
            daily_test_message=daily_test_message,
            makeup_test_message=makeup_test_message,
            makeup_test_date_message=makeup_test_date_message,
        )

        return success_response()
    except Exception as exc:
        return failure_response(exc, context="save_initial_config")


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
            return error_response(RpcErrorCode.URL_REQUIRED, "아이소식 URL이 작성되지 않았습니다.")
        if not url.startswith(("http://", "https://")):
            return error_response(RpcErrorCode.URL_INVALID, "URL이 유효하지 않습니다.")
        if not daily_test_message.strip():
            return error_response(RpcErrorCode.DAILY_MESSAGE_REQUIRED, "테스트 결과 메시지 템플릿이 작성되지 않았습니다.")
        if not makeup_test_message.strip():
            return error_response(RpcErrorCode.MAKEUP_MESSAGE_REQUIRED, "재시험 안내 문구를 입력해 주세요.")
        if not makeup_test_date_message.strip():
            return error_response(RpcErrorCode.MAKEUP_DATE_MESSAGE_REQUIRED, "재시험 일정 안내 문구를 입력해 주세요.")

        tdm.config.update_message_templates(
            url=url,
            daily_test_message=daily_test_message,
            makeup_test_message=makeup_test_message,
            makeup_test_date_message=makeup_test_date_message,
        )
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="update_message_templates")


@server.method()
async def validate_script_url(ctx: RPCContext, url: str) -> Dict[str, Any]:
    try:
        url = (url or "").strip()
        if not url:
            return error_response(RpcErrorCode.URL_REQUIRED, "아이소식 URL이 작성되지 않았습니다.")
        if not url.startswith(("http://", "https://")):
            return error_response(RpcErrorCode.URL_INVALID, "URL이 유효하지 않습니다.")

        warning = _validate_script_page_url(url) is not None
        return success_response(warning=warning)
    except Exception as exc:
        return failure_response(exc, context="validate_script_url")


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
            return success_response(data={"title": "이용약관", "text": fallback_text})

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

        return success_response(data={"title": "이용약관", "text": text})
    except OSError:
        return success_response(data={"title": "이용약관", "text": fallback_text})


@server.method()
async def accept_terms(ctx: RPCContext) -> Dict[str, Any]:
    try:
        tdm.config.accept_terms()
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="accept_terms")


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
            except (OSError, ValueError, UnicodeError):
                # Ignore invalid local override and keep defaults.
                pass

        enabled = bool(raw.get("enabled", True))
        if not enabled:
            return success_response(data={"enabled": False, "title": "공지사항", "message": ""})

        source = str(raw.get("source", "static")).strip().lower()
        if source != "github_release":
            title = str(raw.get("title", "공지사항"))
            message = str(raw.get("message", "")).strip()
            notice_id = hashlib.sha256(f"{title}\n{message}".encode("utf-8")).hexdigest()
            seen_id = tdm.config.get_notice_seen_id()
            should_show = bool(message) and (notice_id != seen_id)
            return success_response(data={
                "enabled": should_show,
                "title": title,
                "message": message,
                "noticeId": notice_id,
            })

        repo = str(raw.get("repo", "LeeMin-hyeong/TestDataManagement")).strip()
        include_prerelease = bool(raw.get("include_prerelease", False))
        if "/" not in repo:
            return error_response(
                RpcErrorCode.NOTICE_REPOSITORY_INVALID,
                "공지사항 저장소 설정이 올바르지 않습니다.",
            )

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
        except (HTTPError, URLError) as exc:
            return failure_response(exc, context="get_startup_notice.fetch")

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
            return success_response(data={"enabled": False, "title": "공지사항", "message": ""})

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
        return success_response(data={
            "enabled": should_show,
            "title": title,
            "message": message,
            "noticeId": notice_id,
        })
    except Exception as exc:
        return failure_response(exc, context="get_startup_notice")


@server.method()
async def mark_notice_seen(ctx: RPCContext, notice_id: str) -> Dict[str, Any]:
    try:
        notice_id = (notice_id or "").strip()
        if not notice_id:
            return error_response(RpcErrorCode.NOTICE_ID_REQUIRED, "공지사항 식별자가 필요합니다.")
        tdm.config.set_notice_seen_id(notice_id)
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="mark_notice_seen")


@server.method()
async def quit_app(ctx: RPCContext) -> Dict[str, Any]:
    try:
        ctx.pyloid.quit()
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="quit_app")


@server.method()
async def confirm_app_exit(ctx: RPCContext) -> Dict[str, Any]:
    try:
        from tdm_host.jobs.registry import registry

        registry.confirm_shutdown()
        ctx.pyloid.quit()
        return success_response()
    except Exception as exc:
        return failure_response(exc, context="confirm_app_exit")


@server.method()
async def get_startup_messages(ctx: RPCContext) -> Dict[str, Any]:
    return success_response(data={
        "termsTitle": "이용약관",
        "termsMessage": (
            "본 프로그램은 등록된 라이선스 사용자만 이용할 수 있습니다.\n"
            "데이터 손실을 방지하기 위해 사용 전 백업을 권장합니다.\n"
            "학원 운영 정책 및 관련 법령을 준수하여 사용해 주세요."
        ),
        "noticeTitle": "공지사항",
        "noticeMessage": "현재 등록된 공지사항이 없습니다.",
    })
