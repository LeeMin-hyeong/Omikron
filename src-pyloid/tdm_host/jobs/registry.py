"""In-memory lifecycle registry for background jobs."""

import multiprocessing
import threading
import time
from queue import Empty
from typing import Any, Dict


progress: dict[str, dict] = {}
job_threads: dict[str, threading.Thread] = {}
job_processes: dict[str, multiprocessing.Process] = {}
progress_queues: dict[str, multiprocessing.Queue] = {}
progress_listeners: dict[str, threading.Thread] = {}
job_process_started_at: dict[str, float] = {}
job_process_seen_payload: dict[str, bool] = {}


def make_emit(job_id: str):
    def emit(payload: dict) -> None:
        previous = progress.get(job_id, {})
        warnings = list(previous.get("warnings", []))
        if payload.get("level") == "warning":
            message = payload.get("message")
            if message:
                message_text = str(message)
                if not warnings or warnings[-1] != message_text:
                    warnings.append(message_text)
        progress[job_id] = {**payload, "warnings": warnings}

    return emit


def queue_listener(
    job_id: str, queue: multiprocessing.Queue, process: multiprocessing.Process
) -> None:
    while True:
        try:
            payload = queue.get(timeout=0.5)
        except Empty:
            if not process.is_alive():
                if process.pid is None:
                    started_at = job_process_started_at.get(job_id, 0)
                    if started_at and (time.time() - started_at) < 5.0:
                        continue
                break
            continue
        if payload is None:
            break
        job_process_seen_payload[job_id] = True
        make_emit(job_id)(payload)
    progress_queues.pop(job_id, None)
    progress_listeners.pop(job_id, None)
    job_process_seen_payload.pop(job_id, None)
    job_process_started_at.pop(job_id, None)


def get_progress_payload(job_id: str) -> Dict[str, Any]:
    default_payload = {
        "step": 0,
        "total": 0,
        "level": "info",
        "status": "unknown",
        "message": "",
        "error": "",
        "detail": "",
        "warnings": [],
        "ts": time.time(),
    }
    payload = progress.get(job_id, default_payload)
    thread = job_threads.get(job_id)
    if thread and not thread.is_alive():
        status = payload.get("status")
        if status in ("running", "unknown"):
            payload = {
                **payload,
                "status": "done",
                "level": "success",
                "message": payload.get("message") or "작업이 완료되었습니다.",
                "ts": time.time(),
            }
            progress[job_id] = payload
        job_threads.pop(job_id, None)

    process = job_processes.get(job_id)
    if process and not process.is_alive():
        status = payload.get("status")
        if status in ("running", "unknown"):
            started_at = job_process_started_at.get(job_id, 0)
            seen_payload = job_process_seen_payload.get(job_id, False)
            if not seen_payload and (time.time() - started_at) < 2.0:
                return payload
            if process.exitcode not in (0, None):
                payload = {
                    **payload,
                    "status": "error",
                    "level": "error",
                    "message": payload.get("message") or "update_class process failed.",
                    "ts": time.time(),
                }
                progress[job_id] = payload
                job_processes.pop(job_id, None)
                return payload
            payload = {
                **payload,
                "status": "done",
                "level": "success",
                "message": payload.get("message") or "작업이 완료되었습니다.",
                "ts": time.time(),
            }
            progress[job_id] = payload
        job_processes.pop(job_id, None)

    return payload
