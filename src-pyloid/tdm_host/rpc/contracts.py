"""Small runtime-validated request DTOs for RPC methods."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any


def _required_text(value: Any, field: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field} 값이 필요합니다.")
    return value.strip()


@dataclass(frozen=True)
class TextRequest:
    value: str

    @classmethod
    def validate(cls, value: Any, field: str) -> TextRequest:
        return cls(_required_text(value, field))


@dataclass(frozen=True)
class UrlRequest:
    url: str

    @classmethod
    def validate(cls, url: Any) -> UrlRequest:
        clean_url = _required_text(url, "url")
        if not clean_url.startswith(("http://", "https://")):
            raise ValueError("url은 http:// 또는 https://로 시작해야 합니다.")
        return cls(clean_url)


@dataclass(frozen=True)
class StudentRequest:
    student_name: str
    class_name: str

    @classmethod
    def validate(cls, student_name: Any, class_name: Any) -> StudentRequest:
        return cls(
            _required_text(student_name, "target_student_name"),
            _required_text(class_name, "target_class_name"),
        )


@dataclass(frozen=True)
class MoveStudentRequest:
    student_name: str
    target_class_name: str
    current_class_name: str

    @classmethod
    def validate(
        cls, student_name: Any, target_class_name: Any, current_class_name: Any
    ) -> MoveStudentRequest:
        return cls(
            _required_text(student_name, "target_student_name"),
            _required_text(target_class_name, "target_class_name"),
            _required_text(current_class_name, "current_class_name"),
        )


@dataclass(frozen=True)
class CellRequest:
    row: int
    col: int

    @classmethod
    def validate(cls, row: Any, col: Any) -> CellRequest:
        if not isinstance(row, int) or row < 1:
            raise ValueError("row는 1 이상의 정수여야 합니다.")
        if not isinstance(col, int) or col < 1:
            raise ValueError("col은 1 이상의 정수여야 합니다.")
        return cls(row, col)


@dataclass(frozen=True)
class JobUploadRequest:
    filename: str
    b64: str
    makeup_test_date: dict[str, str]

    @classmethod
    def validate(
        cls, filename: Any, b64: Any, makeup_test_date: Any
    ) -> JobUploadRequest:
        clean_filename = _required_text(filename, "filename")
        if not clean_filename.lower().endswith(".xlsx"):
            raise ValueError("filename은 .xlsx 파일이어야 합니다.")
        clean_b64 = _required_text(b64, "b64")
        if not isinstance(makeup_test_date, dict):
            raise ValueError("makeup_test_date는 객체여야 합니다.")

        clean_dates: dict[str, str] = {}
        for key, value in makeup_test_date.items():
            clean_key = _required_text(key, "makeup_test_date key")
            clean_value = _required_text(value, f"makeup_test_date.{clean_key}")
            date.fromisoformat(clean_value)
            clean_dates[clean_key] = clean_value
        return cls(clean_filename, clean_b64, clean_dates)


@dataclass(frozen=True)
class JobIdRequest:
    job_id: str

    @classmethod
    def validate(cls, job_id: Any) -> JobIdRequest:
        return cls(_required_text(job_id, "jobId"))


@dataclass(frozen=True)
class JobBatchRequest:
    jobs: list[dict[str, Any]]

    @classmethod
    def validate(cls, jobs: Any) -> JobBatchRequest:
        if not isinstance(jobs, list) or len(jobs) > 100:
            raise ValueError("jobs는 최대 100개의 배열이어야 합니다.")
        validated: list[dict[str, Any]] = []
        for item in jobs:
            if not isinstance(item, dict):
                raise ValueError("jobs 항목은 객체여야 합니다.")
            job_id = JobIdRequest.validate(item.get("jobId") or item.get("job_id")).job_id
            revision = item.get("revision", 0)
            if not isinstance(revision, int) or revision < 0:
                raise ValueError("revision은 0 이상의 정수여야 합니다.")
            validated.append({"jobId": job_id, "revision": revision})
        return cls(validated)
