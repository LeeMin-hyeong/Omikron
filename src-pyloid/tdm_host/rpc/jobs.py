"""Background-job RPC handlers."""

import multiprocessing
import threading
import time
import uuid
from typing import Any, Dict

from pyloid.rpc import RPCContext

from tdm_host.jobs.registry import (
    get_progress_payload,
    job_process_seen_payload,
    job_process_started_at,
    job_processes,
    make_emit,
    progress_listeners,
    progress_queues,
    queue_listener as _queue_listener,
)
from tdm_host.jobs.workers import (
    save_exam_job_process as _save_exam_job_process,
    send_exam_message_job_process as _send_exam_message_job_process,
    update_class_job_process as _update_class_job_process,
)
from tdm_host.rpc.transport import server

@server.method()
async def get_progress(ctx: RPCContext, job_id: str) -> Dict[str, Any]:
    """진행상태 조회 (프런트 폴링)"""
    return get_progress_payload(job_id)


####################################### 데이터 요청 API #######################################


@server.method()
async def start_send_exam_message(ctx: RPCContext, filename: str, b64: str, makeup_test_date: Dict[str, Any]) -> Dict[str, Any]:
    job_id = str(uuid.uuid4())

    make_emit(job_id)({
        "ts": time.time(),
        "step": 0,
        "total": 3,
        "level": "info",
        "status": "running",
        "message": "작업 대기 중...",
        "warnings": [],
    })

    ctx_mp = multiprocessing.get_context("spawn")
    q = ctx_mp.Queue()
    proc = ctx_mp.Process(
        target=_send_exam_message_job_process,
        kwargs={
            "job_id": job_id,
            "q": q,
            "filename": filename,
            "b64": b64,
            "makeup_test_date": makeup_test_date,
        },
        daemon=True,
    )
    progress_queues[job_id] = q
    job_processes[job_id] = proc
    job_process_started_at[job_id] = time.time()
    job_process_seen_payload[job_id] = False
    listener = threading.Thread(
        target=_queue_listener,
        args=(job_id, q, proc),
        daemon=True,
    )
    progress_listeners[job_id] = listener
    listener.start()
    try:
        proc.start()
    except Exception:
        make_emit(job_id)({
            "ts": time.time(),
            "step": 0,
            "total": 0,
            "level": "error",
            "status": "error",
            "message": "send_exam_message process failed to start.",
            "warnings": [],
        })

    return {"job_id": job_id}


@server.method()
async def start_save_exam(ctx: RPCContext, filename: str, b64: str, makeup_test_date: Dict[str, Any]) -> Dict[str, Any]:
    job_id = str(uuid.uuid4())

    make_emit(job_id)({
        "ts": time.time(),
        "step": 0,
        "total": 4,
        "level": "info",
        "status": "running",
        "message": "작업 대기 중...",
        "warnings": [],
    })

    ctx_mp = multiprocessing.get_context("spawn")
    q = ctx_mp.Queue()
    proc = ctx_mp.Process(
        target=_save_exam_job_process,
        kwargs={
            "job_id": job_id,
            "q": q,
            "filename": filename,
            "b64": b64,
            "makeup_test_date": makeup_test_date,
        },
        daemon=True,
    )
    progress_queues[job_id] = q
    job_processes[job_id] = proc
    job_process_started_at[job_id] = time.time()
    job_process_seen_payload[job_id] = False
    listener = threading.Thread(
        target=_queue_listener,
        args=(job_id, q, proc),
        daemon=True,
    )
    progress_listeners[job_id] = listener
    listener.start()
    try:
        proc.start()
    except Exception:
        make_emit(job_id)({
            "ts": time.time(),
            "step": 0,
            "total": 0,
            "level": "error",
            "status": "error",
            "message": "save_exam process failed to start.",
            "warnings": [],
        })

    return {"job_id": job_id}


@server.method()
async def start_update_class(ctx: RPCContext) -> Dict[str, Any]:
    job_id = str(uuid.uuid4())

    make_emit(job_id)({
        "ts": time.time(),
        "step": 0,
        "total": 6,
        "level": "info",
        "status": "running",
        "message": "반 업데이트 준비중...",
        "warnings": [],
    })

    ctx_mp = multiprocessing.get_context("spawn")
    q = ctx_mp.Queue()
    proc = ctx_mp.Process(
        target=_update_class_job_process,
        kwargs={"job_id": job_id, "q": q},
        daemon=True,
    )
    progress_queues[job_id] = q
    job_processes[job_id] = proc
    job_process_started_at[job_id] = time.time()
    job_process_seen_payload[job_id] = False
    listener = threading.Thread(
        target=_queue_listener,
        args=(job_id, q, proc),
        daemon=True,
    )
    progress_listeners[job_id] = listener
    listener.start()
    try:
        proc.start()
    except Exception:
        make_emit(job_id)({
            "ts": time.time(),
            "step": 0,
            "total": 0,
            "level": "error",
            "status": "error",
            "message": "update_class process failed to start.",
            "warnings": [],
        })

    return {"job_id": job_id}
