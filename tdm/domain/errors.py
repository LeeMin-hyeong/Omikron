"""Application exception hierarchy and stable user-facing error codes."""

from __future__ import annotations


class TDMError(Exception):
    """Base class for expected application failures."""

    code = "TDM_ERROR"
    user_message = "작업을 완료하지 못했습니다."
    expose_message = True

    def display_message(self) -> str:
        if self.expose_message and str(self):
            return str(self)
        return self.user_message


class UserActionableError(TDMError):
    """The user can resolve the problem by correcting input or file state."""


class EnvironmentFailure(TDMError):
    """The local PC, Excel, network, or storage environment is unavailable."""

    user_message = "실행 환경 또는 저장 위치 문제로 작업을 완료하지 못했습니다."
    expose_message = False


class ConflictError(TDMError):
    """The source changed while an operation was being prepared."""


class InvalidOperationError(UserActionableError):
    code = "INVALID_OPERATION"
    user_message = "현재 상태에서는 요청한 작업을 수행할 수 없습니다."


class JobAlreadyRunningError(ConflictError):
    code = "JOB_ALREADY_RUNNING"
    user_message = "다른 작업이 진행 중입니다. 완료된 후 다시 시도해 주세요."


class NoMatchingSheetException(UserActionableError):
    code = "WORKBOOK_SHEET_MISSING"
    user_message = "필수 시트를 찾을 수 없습니다. 시트 이름을 확인해 주세요."


class FileOpenException(UserActionableError):
    code = "FILE_OPEN"
    user_message = "파일이 열려 있습니다. 파일을 닫고 다시 시도해 주세요."


class ReopenFileException(UserActionableError):
    code = "WORKBOOK_INVALID"
    user_message = "Excel 파일을 읽을 수 없습니다. 파일을 열었다 닫은 뒤 다시 시도해 주세요."


class ExcelRequiredException(EnvironmentFailure):
    code = "EXCEL_REQUIRED"
    user_message = "이 기능을 사용하려면 Microsoft Excel이 필요합니다."


class ChromeDriverVersionMismatchException(EnvironmentFailure):
    code = "CHROME_DRIVER_MISMATCH"
    user_message = "Chrome과 브라우저 연동 구성요소의 버전이 맞지 않습니다."
