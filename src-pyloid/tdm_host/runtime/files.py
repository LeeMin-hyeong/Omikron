"""Filesystem helpers used by RPC handlers and background workers."""

import base64
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def open_path_cross_platform(path: str) -> None:
    target = os.path.abspath(path)
    if os.name == "nt":
        os.startfile(target)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", target])
    else:
        subprocess.Popen(["xdg-open", target])


def decode_xlsx_upload_to_temp(filename: str, encoded: str) -> Path:
    """Decode an uploaded XLSX payload into an isolated temporary directory."""
    safe_name = Path(filename or "upload.bin").name
    if Path(safe_name).suffix.lower() != ".xlsx":
        raise ValueError("지원하지 않는 파일 형식입니다. .xlsx 파일만 사용할 수 있습니다.")

    temp_root = Path(tempfile.mkdtemp(prefix="tdm_job_"))
    temp_path = temp_root / safe_name
    try:
        temp_path.write_bytes(base64.b64decode(encoded))
        return temp_path
    except Exception:
        shutil.rmtree(temp_root, ignore_errors=True)
        raise


def cleanup_temp(path: Path) -> None:
    """Remove a temporary upload and its containing job directory."""
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
