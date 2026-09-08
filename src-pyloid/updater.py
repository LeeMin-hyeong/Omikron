"""Windows updater: background file work, main-thread dialogs, recoverable installs."""

import ctypes
import os
from pathlib import Path
from queue import Empty, Queue
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import messagebox, ttk

import updater_core as core

ROOT = (Path(sys.executable).parent if getattr(sys, "frozen", False)
        else Path(__file__).parent).resolve()


def resource_path(relative):
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent)) / relative


def log(message):
    core.LOGGER.info(message)


def show_dialog(parent, kind, title, message):
    """Keep errors visible even when Tk itself cannot create a dialog."""
    try:
        getattr(messagebox, kind)(title, message, parent=parent)
    except Exception:
        core.LOGGER.exception("Tk dialog failed")
        if os.name != "nt":
            raise
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        user32.MessageBoxW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_uint]
        user32.MessageBoxW.restype = ctypes.c_int
        icon = 0x10 if kind == "showerror" else 0x30 if kind == "showwarning" else 0x40
        if not user32.MessageBoxW(None, message, title, icon | 0x10000 | 0x40000):
            raise ctypes.WinError(ctypes.get_last_error())


def error_details(exc):
    message = core.describe_error(exc)
    if isinstance(exc, core.InstallError) and not exc.rollback_ok:
        message += ("\n\n프로그램을 실행하지 않았습니다. 백업 폴더를 삭제하지 말고 관리자에게 문의해 주세요."
                    f"\n백업 위치: {exc.backup_dir}")
    message += f"\n\n로그 파일: {core.LOG_PATH}" if core.LOG_PATH else "\n\n로그 파일을 저장하지 못했습니다. 이 오류 내용을 보관해 주세요."
    return message


# ===================== 스플래시 UI =====================
class Updater:
    def __init__(self):
        self.root = tk.Tk()
        self.events = Queue()
        self.closed = False
        self.exit_code = 0
        self.worker = None
        self.fatal_ui_error = False
        self.root.report_callback_exception = self._callback_error
        self.root.protocol("WM_DELETE_WINDOW", self._request_close)

        # === 윈도우 프레임 제거 & 항상 위 ===
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)

        # === 투명 배경 + 둥근 흰 카드 ===
        self.trans_color = "#00FF00"   # 완전 투명으로 사용할 색
        self.card_color  = "#F5F5F5"   # 카드(사각형) 색

        self.root.configure(bg=self.trans_color)
        # Windows에서 trans_color를 완전 투명으로
        self.root.wm_attributes("-transparentcolor", self.trans_color)

        # 창 크기 & 중앙 배치
        self.width, self.height = 360, 260
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = int((sw - self.width) / 2)
        y = int((sh - self.height) / 2)
        self.root.geometry(f"{self.width}x{self.height}+{x}+{y}")

        # === 드래그 이동 지원 ===
        self._add_drag_support()

        # === 캔버스 (투명 배경) ===
        self.canvas = tk.Canvas(
            self.root,
            width=self.width,
            height=self.height,
            bg=self.trans_color,
            highlightthickness=0,
            bd=0,
        )
        self.canvas.pack(fill="both", expand=True)

        # 둥근 흰 사각형(카드) 그리기
        self._draw_rounded_card(
            8, 8,
            self.width - 8,
            self.height - 8,
            radius=8,
            fill=self.card_color
        )

        # === 로고 ===
        self.logo_img = None
        logo_path = resource_path("src/assets/tdm.png")
        if logo_path.exists():
            try:
                self.logo_img = tk.PhotoImage(file=str(logo_path))
            except Exception as e:
                log(f"로고 로드 실패: {e}")

        if self.logo_img is not None:
            self.canvas.create_image(
                self.width // 2,
                95,
                image=self.logo_img
            )
        else:
            self.canvas.create_text(
                self.width // 2,
                80,
                text="tdm",
                fill="#000000",
                font=("Segoe UI", 20, "bold")
            )

        # === 상태 텍스트 + 프로그래스바 담을 프레임 ===
        self.bottom_frame = tk.Frame(
            self.root,
            bg=self.card_color,
            bd=0,
            highlightthickness=0,
        )

        # 상태 텍스트
        self.status_var = tk.StringVar(value="테스트 데이터 관리 프로그램 업데이트 체커")
        self.status_label = tk.Label(
            self.bottom_frame,
            textvariable=self.status_var,
            fg="#000000",
            bg=self.card_color,
            font=("Segoe UI", 11, "bold"),
            borderwidth=0,
            highlightthickness=0,
            justify="center",
        )
        self.status_label.pack(pady=(0, 6))

        # 프로그래스바
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress = ttk.Progressbar(
            self.bottom_frame,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            variable=self.progress_var,
            length=220,
        )
        self.progress.pack()

        # Progressbar 스타일 (연한 회색 바탕 + 빨간 채움)
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "tdm.Horizontal.TProgressbar",
            troughcolor="#383838",
            bordercolor="#383838",
            background="#ed1b24",
            lightcolor="#ed1b24",
            darkcolor="#ed1b24",
        )
        self.progress.configure(style="tdm.Horizontal.TProgressbar")

        # 카드 안, 아래쪽에 frame 올리기
        self.canvas.create_window(
            self.width // 2,
            self.height - 60,
            window=self.bottom_frame
        )

        # 100ms 후 업데이트 쓰레드 시작
        self.root.after(100, self._drain_events)
        self.root.after(100, self.start_update_thread)

    # ------------------ UI 보조 메서드 ------------------
    def _draw_rounded_card(self, x1, y1, x2, y2, radius=20, fill="#FFFFFF"):
        """캔버스에 둥근 모서리 흰 사각형 하나만 그리기 (outline 없음)."""
        r = radius
        # 중앙 직사각형
        self.canvas.create_rectangle(
            x1 + r, y1,
            x2 - r, y2,
            fill=fill, outline=fill
        )
        self.canvas.create_rectangle(
            x1, y1 + r,
            x2, y2 - r,
            fill=fill, outline=fill
        )
        # 네 모서리 원호
        self.canvas.create_oval(
            x1, y1, x1 + 2 * r, y1 + 2 * r,
            fill=fill, outline=fill
        )
        self.canvas.create_oval(
            x2 - 2 * r, y1, x2, y1 + 2 * r,
            fill=fill, outline=fill
        )
        self.canvas.create_oval(
            x1, y2 - 2 * r, x1 + 2 * r, y2,
            fill=fill, outline=fill
        )
        self.canvas.create_oval(
            x2 - 2 * r, y2 - 2 * r, x2, y2,
            fill=fill, outline=fill
        )

    def _add_drag_support(self):
        """헤더가 없으니까, 아무 데나 드래그해서 창 이동."""
        self._drag_x = 0
        self._drag_y = 0

        def on_button_press(event):
            self._drag_x = event.x_root
            self._drag_y = event.y_root

        def on_move(event):
            dx = event.x_root - self._drag_x
            dy = event.y_root - self._drag_y
            self._drag_x = event.x_root
            self._drag_y = event.y_root
            x = self.root.winfo_x() + dx
            y = self.root.winfo_y() + dy
            self.root.geometry(f"+{x}+{y}")

        self.root.bind("<ButtonPress-1>", on_button_press)
        self.root.bind("<B1-Motion>", on_move)

    def set_status(self, text: str, progress: float | None = None):
        # Only Queue operations occur on a worker thread; even root.after is Tk.
        self.events.put(("status", text, progress))

    def _drain_events(self):
        try:
            if self.fatal_ui_error:
                if not self.worker or not self.worker.is_alive():
                    self._close(1)
                return
            while not self.closed:
                try:
                    event = self.events.get_nowait()
                except Empty:
                    break
                kind, *args = event
                if kind == "status":
                    self.status_var.set(args[0])
                    if args[1] is not None:
                        self.progress_var.set(args[1])
                elif kind == "update_failed":
                    exc, can_launch = args
                    message = error_details(exc)
                    if can_launch:
                        message += "\n\n확인을 누르면 업데이트 이전 프로그램의 실행을 시도합니다."
                    show_dialog(self.root, "showwarning" if can_launch else "showerror",
                                "tdm 업데이트 실패", message)
                    if can_launch:
                        self._start_worker(self._launch_existing)
                    else:
                        self._close(1)
                elif kind == "launch_failed":
                    show_dialog(self.root, "showerror", "tdm 실행 실패", error_details(args[0]))
                    self._close(1)
                elif kind == "info_close":
                    show_dialog(self.root, "showinfo", "tdm", args[0])
                    self._close(0)
                elif kind == "cleanup_warning":
                    show_dialog(self.root, "showwarning", "tdm 임시 파일 정리 안내", args[0])
                elif kind == "finished":
                    warnings = args[0]
                    if warnings:
                        show_dialog(self.root, "showwarning", "tdm 업데이트 안내", "\n\n".join(warnings))
                    self._close(0)
        except Exception as exc:
            self._callback_error(type(exc), exc, exc.__traceback__)
        finally:
            if not self.closed:
                self.root.after(100, self._drain_events)

    def _callback_error(self, exc_type, exc, traceback):
        core.LOGGER.error("Updater UI callback failed", exc_info=(exc_type, exc, traceback))
        self.fatal_ui_error = True
        show_dialog(self.root, "showerror", "tdm 업데이트 창 오류", error_details(exc))
        # A UI callback failure must not kill a worker in the middle of a rename.
        if not self.worker or not self.worker.is_alive():
            self._close(1)

    def _request_close(self):
        if (not self.worker or not self.worker.is_alive()) and self.events.empty():
            self._close(0)
            return
        show_dialog(self.root, "showinfo", "tdm", "프로그램 파일을 보호하기 위해 업데이트와 실행 확인이 끝날 때까지 기다려 주세요.")

    def _close(self, code):
        self.exit_code = code
        self.closed = True
        self.root.destroy()

    def _start_worker(self, target):
        # Non-daemon: closing the UI must not interrupt a pending file transaction.
        try:
            self.worker = threading.Thread(target=target, daemon=False)
            self.worker.start()
        except Exception as exc:
            core.LOGGER.exception("Could not start updater worker")
            self.events.put(("launch_failed", core.UpdateError(f"업데이트 작업을 시작하지 못했습니다.\n{core.describe_error(exc)}")))

    def start_update_thread(self):
        self._start_worker(self.run_update_flow)

    def run_update_flow(self):
        transaction = None
        process = None
        temporary = None
        can_launch = False
        launching = False
        warnings = []
        terminal_event = None
        try:
            self.set_status("프로그램 실행 상태 확인 중…", 5)
            if core.is_main_running(ROOT):
                terminal_event = ("info_close", "이 폴더의 tdm이 이미 실행 중입니다. 기존 프로그램 창이나 작업 표시줄의 tdm 아이콘을 확인해 주세요.")
                return

            # Interrupted installs must be recovered before even reading version.txt.
            self.set_status("이전 업데이트 상태 확인 중…", 10)
            warnings.extend(core.recover_pending(ROOT))
            can_launch = True
            current = core.read_local_version(ROOT)
            self.set_status("서버에서 최신 버전 확인 중…", 20)
            latest, asset, checksum = core.fetch_latest_zip_asset()
            if core.cmp_semver(current, latest) < 0:
                working_root = core.safe_child(ROOT, ".update_tmp")
                working_root.mkdir(parents=True, exist_ok=True)
                temporary = tempfile.TemporaryDirectory(prefix="run-", dir=working_root)
                working = Path(temporary.name)
                self.set_status("업데이트 파일 다운로드 중…", 40)
                archive = core.download_asset(asset, working / "update.zip")
                if checksum:
                    self.set_status("파일 무결성 확인 중…", 55)
                    core.verify_sha256(archive, core.gh_get(checksum["browser_download_url"]).decode("utf-8-sig"))
                else:
                    core.LOGGER.info("Release has no optional SHA256 asset")
                self.set_status("업데이트 파일 압축 해제 중…", 65)
                staging = working / "staging"
                core.safe_extract_zip(archive, staging)
                new_root = core.safe_child(staging, "tdm-win")
                core.validate_payload(new_root)
                if core.is_main_running(ROOT):
                    can_launch = False
                    raise core.UpdateError("업데이트 준비 중 tdm이 실행되었습니다. 프로그램을 종료한 후 다시 실행해 주세요.")
                self.set_status("업데이트 적용 중…", 80)
                transaction = core.Transaction(ROOT, new_root, latest)
                transaction.install()

            self.set_status("프로그램 시작 확인 중…", 95)
            launching = True
            process = core.launch_main(ROOT)
            if transaction:
                # Only commit after process creation and the early-exit observation.
                try:
                    warning = transaction.commit()
                    if warning:
                        warnings.append(warning)
                except Exception as exc:
                    # Main is live: never roll back files it may have loaded.
                    core.LOGGER.exception("Could not finalize update while main is running")
                    warnings.append("프로그램을 실행했으나 업데이트 완료 기록을 저장하지 못했습니다. "
                                    "프로그램을 종료한 뒤 업데이터를 다시 실행하면 이전 상태로 복구합니다.\n"
                                    + error_details(exc) + f"\n백업 위치: {transaction.directory}")
            terminal_event = ("finished", warnings)
        except Exception as exc:
            core.LOGGER.exception("Update or launch failed")
            if transaction and transaction.record and transaction.record["phase"] == "pending" and process is None:
                try:
                    # Do not overwrite a concurrent main process or a live child.
                    if core.is_main_running(ROOT):
                        raise core.UpdateError("main.exe가 실행 중이어서 지금 파일을 복구할 수 없습니다.")
                    transaction.rollback()
                    exc = core.InstallError(f"{core.describe_error(exc)}\n업데이트 이전 상태로 복구했습니다.",
                                            rollback_ok=True, backup_dir=transaction.directory)
                except Exception as recovery:
                    core.LOGGER.exception("Rollback failed")
                    exc = core.InstallError(f"{core.describe_error(exc)}\n복구 실패: {core.describe_error(recovery)}",
                                            rollback_ok=False, backup_dir=transaction.directory)
            if isinstance(exc, core.InstallError) and not exc.rollback_ok:
                can_launch = False
            # No installed update means this can also be a normal launch failure.
            if process is None and transaction is None and launching:
                terminal_event = ("launch_failed", exc)
            else:
                terminal_event = ("update_failed", exc, can_launch)
        finally:
            if temporary:
                try:
                    # TemporaryDirectory owns this specific run directory only.
                    temporary.cleanup()
                except OSError as cleanup_error:
                    core.LOGGER.warning("Staging cleanup failed: %s", temporary.name, exc_info=True)
                    cleanup_message = ("업데이트 임시 파일 정리를 완료하지 못했습니다.\n"
                                       f"{temporary.name}\n{error_details(cleanup_error)}")
                    if terminal_event and terminal_event[0] == "finished":
                        warnings.append(cleanup_message)
                    else:
                        self.events.put(("cleanup_warning", cleanup_message))
            if terminal_event:
                self.events.put(terminal_event)

    def _launch_existing(self):
        try:
            self.set_status("기존 프로그램 시작 확인 중…", 95)
            if core.is_main_running(ROOT):
                self.events.put(("info_close", "tdm이 이미 실행 중입니다. 기존 프로그램 창을 확인해 주세요."))
                return
            core.launch_main(ROOT)
            self.events.put(("finished", []))
        except Exception as exc:
            core.LOGGER.exception("Existing main launch failed")
            self.events.put(("launch_failed", exc))

    def run(self):
        self.root.mainloop()
        return self.exit_code


def main():
    lock = core.UpdaterLock(ROOT)
    try:
        core.configure_logging()
        log(f"Updater started: {ROOT}")
        lock.acquire()
        return Updater().run()
    except Exception as exc:
        core.LOGGER.exception("Updater startup failed")
        show_dialog(None, "showerror", "tdm 업데이터 실행 실패", error_details(exc))
        return 1
    finally:
        lock.close()


if __name__ == "__main__":
    sys.exit(main())
