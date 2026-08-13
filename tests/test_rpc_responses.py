from __future__ import annotations

import logging

from tdm.domain.progress import Progress
from tdm.excel.errors import (
    ConcurrentWorkbookChangeError,
    WorkbookBusyError,
    WorkbookTransactionError,
)
from tdm_host.rpc.responses import error_response, failure_response, success_response


def _reset_diagnostic_logger() -> None:
    logger = logging.getLogger("tdm")
    for handler in list(logger.handlers):
        handler.close()
        logger.removeHandler(handler)


def test_expected_exception_has_stable_code_without_traceback(tmp_path, monkeypatch):
    _reset_diagnostic_logger()
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))

    try:
        raise WorkbookBusyError(tmp_path / "데이터.xlsx")
    except WorkbookBusyError as exc:
        response = failure_response(exc, context="test.busy")

    assert response["ok"] is False
    assert response["code"] == "WORKBOOK_BUSY"
    assert "파일을 닫은 뒤" in response["error"]
    assert response["detail"].startswith("오류 ID: ")
    assert "Traceback" not in response["detail"]
    _reset_diagnostic_logger()


def test_internal_exception_does_not_expose_diagnostic_message(tmp_path, monkeypatch):
    _reset_diagnostic_logger()
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))

    secret = "internal implementation detail"
    try:
        raise RuntimeError(secret)
    except RuntimeError as exc:
        response = failure_response(exc, context="test.internal")

    assert response["code"] == "INTERNAL_ERROR"
    assert secret not in response["error"]
    assert secret not in response["detail"]
    assert secret in (tmp_path / "TDM" / "logs" / "tdm.log").read_text(
        encoding="utf-8"
    )
    _reset_diagnostic_logger()


def test_conflict_and_explicit_responses_follow_contract(tmp_path):
    conflict = ConcurrentWorkbookChangeError(tmp_path / "데이터.xlsx")
    assert conflict.code == "WORKBOOK_CONFLICT"
    assert success_response(data=1) == {"ok": True, "data": 1}
    assert error_response("INVALID_INPUT", "확인해 주세요.") == {
        "ok": False,
        "code": "INVALID_INPUT",
        "error": "확인해 주세요.",
    }


def test_environment_exception_hides_technical_message(tmp_path, monkeypatch):
    _reset_diagnostic_logger()
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    response = failure_response(
        WorkbookTransactionError("rollback path C:/secret failed"),
        context="test.transaction",
    )

    assert response["code"] == "WORKBOOK_TRANSACTION"
    assert "C:/secret" not in response["error"]
    assert "C:/secret" not in response["detail"]
    _reset_diagnostic_logger()


def test_progress_error_carries_error_code():
    events: list[dict] = []
    progress = Progress(events.append, total=1)

    progress.error("저장 실패", code="WORKBOOK_BUSY", detail="오류 ID: ABC")

    assert events[-1]["code"] == "WORKBOOK_BUSY"
    assert events[-1]["detail"] == "오류 ID: ABC"
