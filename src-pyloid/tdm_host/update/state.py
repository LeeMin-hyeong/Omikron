"""Small, atomic update-state helpers shared by the main application."""

from __future__ import annotations

import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


_TRANSACTION_ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]{1,64}$")


def application_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd().resolve()


def update_root(root: Path | None = None) -> Path:
    return (root or application_root()) / ".update"


def atomic_write_json(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    try:
        with temporary.open("w", encoding="utf-8", newline="\n") as stream:
            json.dump(value, stream, ensure_ascii=False, indent=2, sort_keys=True)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def write_update_state(root: Path, **values: Any) -> Path:
    state_path = update_root(root) / "update-state.json"
    payload = {
        "schemaVersion": 1,
        "updatedAt": datetime.now(timezone.utc).isoformat(),
        **values,
    }
    atomic_write_json(state_path, payload)
    return state_path


def health_marker_path(root: Path, transaction_id: str) -> Path:
    if not _TRANSACTION_ID_PATTERN.fullmatch(transaction_id):
        raise ValueError("Invalid update transaction ID")
    return update_root(root) / "health" / f"{transaction_id}.startup-ok"


def mark_startup_healthy(root: Path, transaction_id: str) -> Path:
    marker = health_marker_path(root, transaction_id)
    atomic_write_json(
        marker,
        {
            "schemaVersion": 1,
            "transactionId": transaction_id,
            "createdAt": datetime.now(timezone.utc).isoformat(),
            "pid": os.getpid(),
        },
    )
    return marker


def post_update_transaction_id(arguments: list[str] | None = None) -> str | None:
    args = list(sys.argv[1:] if arguments is None else arguments)
    try:
        index = args.index("--post-update")
        value = args[index + 1].strip()
    except (ValueError, IndexError):
        return None
    if not _TRANSACTION_ID_PATTERN.fullmatch(value):
        return None
    return value
