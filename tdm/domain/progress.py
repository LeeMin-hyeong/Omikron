# src-pyloid/progress.py
from __future__ import annotations
from typing import Callable, Literal, Optional
import time

Level = Literal["info", "success", "warning", "error"]
Status = Literal["running", "done", "error"]

class Progress:
    """
    .info/.warning/.error 등으로 호출.
    emit(payload: dict) 콜백으로 외부(RPC progress dict / 이벤트 버스)로 전달.
    """
    def __init__(self, emit: Callable[[dict], None], total: Optional[int] = None):
        self.emit_cb = emit
        self.step_no = 0
        self.total = total
        self.phase_step: Optional[int] = None
        self.phase_total: Optional[int] = None

    def _post(
        self,
        message: str,
        level: Level = "info",
        status: Status = "running",
        inc: bool = False,
        *,
        error: Optional[str] = None,
        detail: Optional[str] = None,
    ):
        if inc:
            self.step_no += 1
        payload = {
            "ts": time.time(),
            "step": self.step_no,
            "total": self.total,
            "phase_step": self.phase_step,
            "phase_total": self.phase_total,
            "level": level,
            "status": status,
            "message": message,
        }
        if error is not None:
            payload["error"] = error
        if detail is not None:
            payload["detail"] = detail
        self.emit_cb(payload)

    def info(self, msg: str, *, inc: bool = False):    self._post(msg, "info",    "running", inc)
    def success(self, msg: str, *, inc: bool = False): self._post(msg, "success", "running", inc)
    def warning(self, msg: str, *, inc: bool = False): self._post(msg, "warning", "running", inc)
    def error(self, msg: str, *, detail: Optional[str] = None, inc: bool = False):
        self._post(msg, "error", "error", inc, error=msg, detail=detail)
    def step(self, msg: str):
        self.phase_step = None
        self.phase_total = None
        self._post(msg, "info", "running", True)
    def phase(self, step: int, total: int, msg: str, level: Level = "info"):
        self.phase_step = step
        self.phase_total = total
        self._post(msg, level, "running", False)
    def done(self, msg: str = "완료"):
        self.phase_step = None
        self.phase_total = None
        self._post(msg, "success", "done", False)
