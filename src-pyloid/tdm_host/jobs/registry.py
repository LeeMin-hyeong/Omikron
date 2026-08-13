"""Thread-safe lifecycle registry for background jobs."""

from __future__ import annotations

import multiprocessing
import threading
import time
from dataclasses import dataclass
from queue import Empty
from typing import Any

from tdm.domain.errors import InvalidOperationError, JobAlreadyRunningError


TERMINAL_STATUSES = {"done", "error", "cancelled"}
DEFAULT_FINISHED_TTL_SECONDS = 20 * 60


@dataclass
class JobResources:
    process: multiprocessing.Process | None = None
    queue: multiprocessing.Queue | None = None
    listener: threading.Thread | None = None
    cancel_event: Any = None


class JobRegistry:
    def __init__(self, *, finished_ttl_seconds: float = DEFAULT_FINISHED_TTL_SECONDS) -> None:
        self._lock = threading.RLock()
        self._states: dict[str, dict[str, Any]] = {}
        self._resources: dict[str, JobResources] = {}
        self._finished_ttl_seconds = finished_ttl_seconds
        self._shutting_down = False
        self._shutdown_confirmed = False

    def create(self, job_id: str, job_type: str, payload: dict[str, Any]) -> None:
        now = time.time()
        state = {
            **payload,
            "jobId": job_id,
            "jobType": job_type,
            "revision": 1,
            "createdAt": now,
            "updatedAt": now,
            "finishedAt": None,
            "cancellationRequested": False,
        }
        with self._lock:
            self._states[job_id] = state
            self._resources[job_id] = JobResources()
            self.cleanup_expired(now=now)

    def reserve(
        self,
        job_id: str,
        job_type: str,
        payload: dict[str, Any],
        *,
        exclusive_group: str | None = None,
    ) -> None:
        """Atomically register a job after checking its exclusive group."""
        with self._lock:
            if self._shutting_down:
                raise InvalidOperationError("프로그램이 종료 중이라 새 작업을 시작할 수 없습니다.")
            if exclusive_group is not None:
                conflict = any(
                    state.get("exclusiveGroup") == exclusive_group
                    and (
                        state.get("status") not in TERMINAL_STATUSES
                        or self._resource_is_alive(job_key)
                    )
                    for job_key, state in self._states.items()
                )
                if conflict:
                    raise JobAlreadyRunningError(
                        "같은 종류의 파일 작업이 이미 진행 중입니다. 완료 후 다시 시도해 주세요."
                    )
            self.create(
                job_id,
                job_type,
                {**payload, "exclusiveGroup": exclusive_group},
            )

    def _resource_is_alive(self, job_id: str) -> bool:
        resources = self._resources.get(job_id)
        if resources is None or resources.process is None:
            return False
        try:
            return resources.process.is_alive()
        except (AssertionError, OSError, ValueError):
            return False

    def confirm_shutdown(self) -> None:
        with self._lock:
            self._shutdown_confirmed = True

    def is_shutdown_confirmed(self) -> bool:
        with self._lock:
            return self._shutdown_confirmed

    def emit(self, job_id: str, payload: dict[str, Any]) -> None:
        now = time.time()
        with self._lock:
            previous = self._states.get(job_id)
            if previous is None:
                self.create(job_id, "unknown", payload)
                return

            warnings = list(previous.get("warnings", []))
            if payload.get("level") == "warning":
                message = payload.get("message")
                if message:
                    message_text = str(message)
                    if not warnings or warnings[-1] != message_text:
                        warnings.append(message_text)
            elif "warnings" in payload:
                for warning in payload.get("warnings") or []:
                    warning_text = str(warning)
                    if warning_text not in warnings:
                        warnings.append(warning_text)

            next_state = {**previous, **payload, "warnings": warnings}
            comparable_keys = set(next_state) - {"revision", "updatedAt", "ts"}
            changed = any(previous.get(key) != next_state.get(key) for key in comparable_keys)
            if not changed:
                return
            next_state["revision"] = int(previous.get("revision", 0)) + 1
            next_state["updatedAt"] = now
            if next_state.get("status") in TERMINAL_STATUSES:
                next_state["finishedAt"] = previous.get("finishedAt") or now
            self._states[job_id] = next_state

    def attach_process(
        self,
        job_id: str,
        process: multiprocessing.Process,
        queue: multiprocessing.Queue,
        listener: threading.Thread,
        cancel_event: Any,
    ) -> None:
        with self._lock:
            self._resources[job_id] = JobResources(process, queue, listener, cancel_event)

    def get(self, job_id: str) -> dict[str, Any]:
        with self._lock:
            state = self._states.get(job_id)
            if state is None:
                return {
                    "jobId": job_id,
                    "revision": 0,
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
            return dict(state)

    def get_many(self, requests: list[dict[str, Any]]) -> list[dict[str, Any]]:
        self.cleanup_expired()
        result: list[dict[str, Any]] = []
        for request in requests:
            job_id = str(request.get("jobId") or request.get("job_id") or "")
            known_revision = int(request.get("revision") or 0)
            state = self.get(job_id)
            revision = int(state.get("revision") or 0)
            if revision == known_revision:
                result.append({"jobId": job_id, "revision": revision, "changed": False})
            else:
                result.append(
                    {"jobId": job_id, "revision": revision, "changed": True, "state": state}
                )
        return result

    def acknowledge(self, job_id: str) -> bool:
        with self._lock:
            state = self._states.get(job_id)
            if state is None:
                return False
            if state.get("status") not in TERMINAL_STATUSES:
                return False
            del self._states[job_id]
            return True

    def request_cancel(self, job_id: str) -> str:
        """Record a cooperative cancellation request without killing the process."""
        with self._lock:
            state = self._states.get(job_id)
            if state is None:
                return "not_found"
            if state.get("status") in TERMINAL_STATUSES:
                return "not_cancellable"
            resources = self._resources.get(job_id)
            if resources is None or resources.cancel_event is None:
                return "not_cancellable"
            resources.cancel_event.set()
            state["cancellationRequested"] = True
            state["revision"] = int(state.get("revision", 0)) + 1
            state["updatedAt"] = time.time()
            return "requested"

    def release_resources(self, job_id: str) -> None:
        with self._lock:
            resources = self._resources.pop(job_id, None)
        if resources is None:
            return
        process = resources.process
        if process is not None:
            try:
                process.join(timeout=1.0)
            except (AssertionError, OSError, ValueError):
                pass
            try:
                process.close()
            except (AttributeError, OSError, ValueError):
                pass
        queue = resources.queue
        if queue is not None:
            try:
                queue.close()
                queue.cancel_join_thread()
            except (AttributeError, OSError, ValueError):
                pass

    def shutdown(self, *, timeout: float = 3.0) -> None:
        """Stop accepting jobs and reap all child-process resources."""
        with self._lock:
            if self._shutting_down:
                return
            self._shutting_down = True
            resources = list(self._resources.items())

        timeout = max(0.0, timeout)
        graceful_deadline = time.monotonic() + (timeout * 0.7)
        for _, item in resources:
            process = item.process
            if process is None:
                continue
            try:
                process.join(timeout=max(0.0, graceful_deadline - time.monotonic()))
            except (AssertionError, OSError, ValueError):
                pass

        for job_id, item in resources:
            process = item.process
            if process is None:
                continue
            try:
                if process.is_alive():
                    self.emit(
                        job_id,
                        {
                            "status": "cancelled",
                            "level": "warning",
                            "message": "프로그램 종료로 작업이 중단되었습니다.",
                            "code": "APP_SHUTDOWN",
                            "ts": time.time(),
                        },
                    )
                    process.terminate()
            except (AssertionError, OSError, ValueError):
                pass

        deadline = time.monotonic() + (timeout * 0.3)
        for _, item in resources:
            process = item.process
            if process is None:
                continue
            try:
                process.join(timeout=max(0.0, deadline - time.monotonic()))
            except (AssertionError, OSError, ValueError):
                pass

        for job_id, item in resources:
            listener = item.listener
            if listener is not None and listener is not threading.current_thread():
                listener.join(timeout=max(0.0, deadline - time.monotonic()))
            self.release_resources(job_id)

    def cleanup_expired(self, *, now: float | None = None) -> None:
        current = time.time() if now is None else now
        with self._lock:
            expired = [
                job_id
                for job_id, state in self._states.items()
                if state.get("finishedAt") is not None
                and current - float(state["finishedAt"]) >= self._finished_ttl_seconds
            ]
            for job_id in expired:
                self._states.pop(job_id, None)
                self._resources.pop(job_id, None)


registry = JobRegistry()


def make_emit(job_id: str):
    return lambda payload: registry.emit(job_id, payload)


def queue_listener(
    job_id: str, queue: multiprocessing.Queue, process: multiprocessing.Process
) -> None:
    saw_terminal = False
    process_exit_seen_at: float | None = None
    try:
        while True:
            try:
                payload = queue.get(timeout=0.5)
            except Empty:
                if process.is_alive():
                    process_exit_seen_at = None
                    continue
                if process_exit_seen_at is None:
                    process_exit_seen_at = time.monotonic()
                    continue
                # multiprocessing.Queue의 feeder가 자식 종료보다 늦게 마지막
                # 이벤트를 전달할 수 있으므로 terminal 이벤트를 잠시 더 기다린다.
                if time.monotonic() - process_exit_seen_at < 1.0:
                    continue
                break
            except (EOFError, OSError, ValueError):
                break
            if payload is None:
                break
            registry.emit(job_id, payload)
            if payload.get("status") in TERMINAL_STATUSES:
                saw_terminal = True
        if not saw_terminal:
            current = registry.get(job_id)
            if current.get("status") not in TERMINAL_STATUSES:
                if process.exitcode in (0, None):
                    registry.emit(
                        job_id,
                        {
                            "status": "done",
                            "level": "success",
                            "message": current.get("message") or "작업이 완료되었습니다.",
                            "ts": time.time(),
                        },
                    )
                else:
                    registry.emit(
                        job_id,
                        {
                            "status": "error",
                            "level": "error",
                            "message": "작업 프로세스가 비정상 종료되었습니다.",
                            "error": "작업 프로세스가 비정상 종료되었습니다.",
                            "code": "JOB_PROCESS_EXIT",
                            "detail": f"exitcode={process.exitcode}",
                            "ts": time.time(),
                        },
                    )
    finally:
        registry.release_resources(job_id)


def get_progress_payload(job_id: str) -> dict[str, Any]:
    """Compatibility endpoint for older clients."""
    return registry.get(job_id)
