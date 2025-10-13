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

    def _post(self, message: str, level: Level = "info",
              status: Status = "running", inc: bool = False):
        if inc:
            self.step_no += 1
        payload = {
            "ts": time.time(),
            "step": self.step_no,
            "total": self.total,
            "level": level,
            "status": status,
            "message": message,
        }
        self.emit_cb(payload)

    def info(self, msg: str, *, inc: bool = False):    self._post(msg, "info",    "running", inc)
    def success(self, msg: str, *, inc: bool = False): self._post(msg, "success", "running", inc)
    def warning(self, msg: str, *, inc: bool = False): self._post(msg, "warning", "running", inc)
    def error(self, msg: str, *, inc: bool = False):   self._post(msg, "error",   "error",   inc)
    def step(self, msg: str):                          self._post(msg, "info",    "running", True)
    def done(self, msg: str = "완료"):                 self._post(msg, "success", "done",    False)
