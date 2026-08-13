"""Storage-level errors shared by every workbook module."""

from __future__ import annotations

from pathlib import Path

from tdm.domain.errors import (
    ConflictError,
    EnvironmentFailure,
    FileOpenException,
    ReopenFileException,
)


class WorkbookBusyError(FileOpenException):
    """The destination cannot be replaced because another program is using it."""

    code = "WORKBOOK_BUSY"

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        super().__init__(
            f"{self.path.name} 파일이 열려 있거나 사용 중입니다. "
            "파일을 닫은 뒤 다시 시도해 주세요."
        )


class ConcurrentWorkbookChangeError(ConflictError):
    """The source changed after this process loaded it."""

    code = "WORKBOOK_CONFLICT"

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        super().__init__(
            f"{self.path.name} 파일이 작업 중 다른 위치에서 변경되었습니다. "
            "최신 파일을 다시 불러온 뒤 작업해 주세요."
        )


class InvalidWorkbookError(ReopenFileException):
    """A workbook is missing required XLSX structures or cannot be read."""


class WorkbookTransactionError(EnvironmentFailure):
    """A multi-workbook transaction could not be committed or recovered."""

    code = "WORKBOOK_TRANSACTION"
    user_message = "여러 파일을 함께 저장하는 중 문제가 발생해 변경을 취소했습니다."


class WorkbookRecoveryRequiredError(WorkbookTransactionError):
    """Recovery stopped because a target contains an unknown newer revision."""

    code = "WORKBOOK_RECOVERY_REQUIRED"
    user_message = "미완료 저장 작업을 자동으로 복구할 수 없습니다. 진단 로그를 확인해 주세요."
