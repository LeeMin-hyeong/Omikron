"""Non-destructive write capability checks for the configured data directory."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path


def verify_directory_writable(directory: str | Path) -> None:
    root = Path(directory)
    root.mkdir(parents=True, exist_ok=True)
    descriptor, first_name = tempfile.mkstemp(dir=root, prefix=".tdm-probe-")
    first = Path(first_name)
    second = first.with_suffix(".renamed")
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(b"tdm-storage-probe")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(first, second)
        if second.read_bytes() != b"tdm-storage-probe":
            raise OSError(f"저장소 쓰기 검증 결과가 일치하지 않습니다: {root}")
    finally:
        first.unlink(missing_ok=True)
        second.unlink(missing_ok=True)

