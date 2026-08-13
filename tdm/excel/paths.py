"""Central path and filename policy for TDM workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import tdm.config
from tdm.domain.models import ClassInfo, DataFile, MakeupTestList, StudentInfo


BACKUP_TIMESTAMP_FORMAT = "%Y%m%d%H%M%S"
STAGING_SUFFIX = ".tmp.xlsx"
TRANSACTION_DIR_NAME = ".tdm"


def backup_filename(stem: str, when: datetime | None = None) -> str:
    timestamp = (when or datetime.now()).strftime(BACKUP_TIMESTAMP_FORMAT)
    return f"{stem}({timestamp}).xlsx"


@dataclass(frozen=True)
class WorkbookPaths:
    root: Path
    data_file_name: str

    @classmethod
    def current(cls) -> "WorkbookPaths":
        return cls(Path(tdm.config.DATA_DIR), tdm.config.DATA_FILE_NAME)

    @property
    def data_dir(self) -> Path:
        return self.root / "data"

    @property
    def backup_dir(self) -> Path:
        return self.data_dir / "backup"

    @property
    def metadata_dir(self) -> Path:
        return self.root / TRANSACTION_DIR_NAME

    @property
    def transaction_dir(self) -> Path:
        return self.metadata_dir / "transactions"

    @property
    def data_file(self) -> Path:
        return self.data_dir / f"{self.data_file_name}.xlsx"

    @property
    def data_temp(self) -> Path:
        return self.data_dir / f"{DataFile.TEMP_FILE_NAME}.xlsx"

    @property
    def previous_data(self) -> Path:
        return self.data_dir / f"{DataFile.PRE_DATA_FILE_NAME}.xlsx"

    @property
    def class_info(self) -> Path:
        return self.root / f"{ClassInfo.DEFAULT_NAME}.xlsx"

    @property
    def class_info_temp(self) -> Path:
        return self.root / f"{ClassInfo.TEMP_FILE_NAME}.xlsx"

    @property
    def class_info_temp_revision(self) -> Path:
        return self.class_info_temp.with_suffix(".xlsx.revision.json")

    @property
    def student_info(self) -> Path:
        return self.root / f"{StudentInfo.DEFAULT_NAME}.xlsx"

    @property
    def makeup_test(self) -> Path:
        return self.data_dir / f"{MakeupTestList.DEFAULT_NAME}.xlsx"

    def backup_path(self, stem: str, when: datetime | None = None) -> Path:
        return self.backup_dir / backup_filename(stem, when)

    def daily_form_path(
        self,
        when: datetime | None = None,
        sequence: int | None = None,
    ) -> Path:
        date_text = (when or datetime.now()).strftime("%m.%d")
        suffix = f" ({sequence})" if sequence is not None else ""
        return self.root / f"데일리테스트 기록 양식({date_text}){suffix}.xlsx"

    def next_daily_form_path(self, when: datetime | None = None) -> Path:
        first = self.daily_form_path(when)
        if not first.exists():
            return first
        sequence = 1
        while self.daily_form_path(when, sequence).exists():
            sequence += 1
        return self.daily_form_path(when, sequence)

    def ensure_directories(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        self.transaction_dir.mkdir(parents=True, exist_ok=True)
