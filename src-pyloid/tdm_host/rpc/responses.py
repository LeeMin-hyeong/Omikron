"""Stable success and failure payloads shared by RPC handlers."""

from __future__ import annotations

from typing import Any

from tdm.domain.errors import TDMError
from tdm_host.runtime.diagnostics import diagnostic_detail
from tdm_host.rpc.error_codes import RpcErrorCode


_MISSING = object()


def success_response(data: Any = _MISSING, **payload: Any) -> dict[str, Any]:
    """Build the canonical success envelope while accepting legacy keyword payloads."""
    if data is not _MISSING and payload:
        raise ValueError("success_response accepts either data or keyword payloads")
    normalized = payload if data is _MISSING else data
    return {"ok": True, "data": normalized}


def error_response(
    code: str | RpcErrorCode,
    error: str,
    *,
    detail: str | None = None,
    **payload: Any,
) -> dict[str, Any]:
    response: dict[str, Any] = {
        "ok": False,
        "code": str(code),
        "error": error,
        **payload,
    }
    if detail:
        response["detail"] = detail
    return response


def _classify_exception(exc: BaseException) -> tuple[str | RpcErrorCode, str]:
    if isinstance(exc, TDMError):
        return exc.code, exc.display_message()
    if isinstance(exc, FileExistsError):
        return RpcErrorCode.FILE_ALREADY_EXISTS, "같은 이름의 파일이 이미 존재합니다."
    if isinstance(exc, FileNotFoundError):
        return RpcErrorCode.FILE_NOT_FOUND, "필요한 파일을 찾을 수 없습니다."
    if isinstance(exc, PermissionError):
        return RpcErrorCode.STORAGE_PERMISSION_DENIED, "저장 위치에 접근할 권한이 없습니다."
    if isinstance(exc, (ConnectionError, TimeoutError)):
        return RpcErrorCode.STORAGE_UNAVAILABLE, "저장 위치 또는 네트워크에 연결할 수 없습니다."
    if isinstance(exc, OSError):
        return RpcErrorCode.ENVIRONMENT_IO, "파일 또는 저장 위치에 접근할 수 없습니다."
    if isinstance(exc, ValueError):
        return RpcErrorCode.INVALID_INPUT, "입력값이 올바르지 않습니다."
    return RpcErrorCode.INTERNAL_ERROR, "예상하지 못한 오류가 발생했습니다."


def failure_response(
    exc: BaseException,
    *,
    context: str = "RPC",
) -> dict[str, Any]:
    code, message = _classify_exception(exc)
    return error_response(
        code,
        message,
        detail=diagnostic_detail(exc, context=context),
    )
