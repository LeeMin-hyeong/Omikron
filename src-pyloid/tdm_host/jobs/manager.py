"""Common process creation and shutdown for background jobs."""

from __future__ import annotations

import multiprocessing
import threading
import time
import uuid
from collections.abc import Callable
from typing import Any

from tdm_host.jobs.registry import make_emit, queue_listener, registry
from tdm_host.rpc.responses import failure_response


class JobManager:
    def start(
        self,
        *,
        job_type: str,
        target: Callable[..., None],
        total: int,
        message: str,
        kwargs: dict[str, Any],
        exclusive_group: str | None = None,
    ) -> dict[str, Any]:
        job_id = str(uuid.uuid4())
        registry.reserve(
            job_id,
            job_type,
            {
                "ts": time.time(),
                "step": 0,
                "total": total,
                "level": "info",
                "status": "running",
                "message": message,
                "warnings": [],
            },
            exclusive_group=exclusive_group,
        )
        context = multiprocessing.get_context("spawn")
        queue = context.Queue()
        cancel_event = context.Event()
        process = context.Process(
            target=target,
            kwargs={"job_id": job_id, "q": queue, "cancel_event": cancel_event, **kwargs},
            daemon=True,
            name=f"tdm-{job_type}-{job_id[:8]}",
        )
        listener = threading.Thread(
            target=queue_listener,
            args=(job_id, queue, process),
            daemon=True,
            name=f"job-listener-{job_type}-{job_id[:8]}",
        )
        registry.attach_process(job_id, process, queue, listener, cancel_event)
        try:
            process.start()
            listener.start()
        except Exception as exc:
            response = failure_response(exc, context=f"{job_type}.start")
            make_emit(job_id)(
                {
                    "ts": time.time(),
                    "level": "error",
                    "status": "error",
                    "message": response["error"],
                    "error": response["error"],
                    "code": response["code"],
                    "detail": response.get("detail"),
                }
            )
            try:
                if process.is_alive():
                    process.terminate()
            except (AssertionError, OSError, ValueError):
                pass
            registry.release_resources(job_id)
            return response
        return {
            "ok": True,
            "data": {
                "jobId": job_id,
                "jobType": job_type,
                "status": "running",
            },
        }

    def shutdown(self, *, timeout: float = 3.0) -> None:
        registry.shutdown(timeout=timeout)


job_manager = JobManager()
