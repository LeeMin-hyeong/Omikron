import multiprocessing
import ctypes
from pyloid.tray import (
    TrayEvent,
)
from pyloid.utils import (
    get_production_path,
    is_production,
)
from pyloid.serve import pyloid_serve
from pyloid import Pyloid
from tdm_host.rpc.server import server
from license import verify_license_or_exit

WIDTH, HEIGHT = 1400, 830


def _enable_dpi_awareness():
    # Make Windows return non-virtualized metrics/work-area values under display scaling.
    if not hasattr(ctypes, "windll"):
        return

    user32 = ctypes.windll.user32
    shcore = getattr(ctypes.windll, "shcore", None)

    try:
        # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        user32.SetProcessDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        user32.SetProcessDpiAwarenessContext.restype = ctypes.c_bool
        if user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4 & 0xFFFFFFFFFFFFFFFF)):
            return
    except Exception:
        pass

    try:
        if shcore is not None:
            # PROCESS_PER_MONITOR_DPI_AWARE = 2
            shcore.SetProcessDpiAwareness.argtypes = [ctypes.c_int]
            shcore.SetProcessDpiAwareness.restype = ctypes.c_long
            if shcore.SetProcessDpiAwareness(2) in (0, 0x00000005):  # S_OK / E_ACCESSDENIED
                return
    except Exception:
        pass

    try:
        user32.SetProcessDPIAware()
    except Exception:
        pass


def _get_screen_size():
    try:
        user32 = ctypes.windll.user32
        return int(user32.GetSystemMetrics(0)), int(user32.GetSystemMetrics(1))
    except Exception:
        try:
            import tkinter as tk

            root = tk.Tk()
            root.withdraw()
            width = int(root.winfo_screenwidth())
            height = int(root.winfo_screenheight())
            root.destroy()
            return width, height
        except Exception:
            return WIDTH, HEIGHT


def _get_work_area():
    try:
        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long),
            ]

        rect = RECT()
        SPI_GETWORKAREA = 0x0030
        ok = ctypes.windll.user32.SystemParametersInfoW(
            SPI_GETWORKAREA, 0, ctypes.byref(rect), 0
        )
        if ok:
            return int(rect.left), int(rect.top), int(rect.right), int(rect.bottom)
    except Exception:
        pass

    sw, sh = _get_screen_size()
    return 0, 0, sw, sh


def _get_window_geometry():
    left, top, right, bottom = _get_work_area()
    work_w = max(1, right - left)
    work_h = max(1, bottom - top)
    # When the screen is smaller than the base window, use at most 90% of the
    # available work area (taskbar excluded) while preserving aspect ratio.
    avail_w = max(1, int(work_w * 0.9))
    avail_h = max(1, int(work_h * 0.9))

    # Preserve the original 1400:830 aspect ratio exactly (integer-rounded).
    if avail_w >= WIDTH and avail_h >= HEIGHT:
        window_w, window_h = WIDTH, HEIGHT
    elif avail_w * HEIGHT <= avail_h * WIDTH:
        window_w = avail_w
        window_h = max(1, round(window_w * HEIGHT / WIDTH))
    else:
        window_h = avail_h
        window_w = max(1, round(window_h * WIDTH / HEIGHT))
    window_x = left + (work_w - window_w) // 2
    window_y = top + (work_h - window_h) // 2

    return window_w, window_h, window_x, window_y


def _recenter_window_to_work_area(window, fallback_w, fallback_h):
    left, top, right, bottom = _get_work_area()
    work_w = max(1, right - left)
    work_h = max(1, bottom - top)

    actual_w = fallback_w
    actual_h = fallback_h
    try:
        size = window.get_size()
        actual_w = int(size.get("width", fallback_w))
        actual_h = int(size.get("height", fallback_h))
    except Exception:
        pass

    x = left + max(0, (work_w - actual_w) // 2)
    y = top + max(0, (work_h - actual_h) // 2)

    # Clamp to work area in case frame/shadow size makes the actual window larger than expected.
    x = min(max(x, left), max(left, right - actual_w))
    y = min(max(y, top), max(top, bottom - actual_h))

    try:
        window.set_position(x, y)
    except Exception:
        pass


def main():
    verify_license_or_exit()
    _enable_dpi_awareness()
    app = Pyloid(app_name="tdm", single_instance=True, server=server)
    window_width, window_height, window_x, window_y = _get_window_geometry()

    app.set_icon(get_production_path("src-pyloid/icons/tdm_icon.ico"))
    app.set_tray_icon(get_production_path("src-pyloid/icons/tdm_icon.ico"))

    ############################## Tray ################################
    def on_double_click():
        app.show_and_focus_main_window()

    app.set_tray_actions(
        {
            TrayEvent.DoubleClick: on_double_click,
        }
    )
    app.set_tray_menu_items(
        [
            {"label": "Show Window", "callback": app.show_and_focus_main_window},
            {"label": "Exit", "callback": app.quit},
        ]
    )
    ####################################################################

    if is_production():
        url = pyloid_serve(directory=get_production_path("dist-front"))
        window = app.create_window(
            title="테스트 데이터 관리",
            width=window_width,
            height=window_height,
            x=window_x,
            y=window_y,
            transparent=True,
        )
        window.load_url(url)
    else:
        window = app.create_window(
            title="테스트 데이터 관리-dev",
            dev_tools=True,
            width=window_width,
            height=window_height,
            x=window_x,
            y=window_y,
            transparent=True,
        )
        window.load_url("http://localhost:5173")

    _recenter_window_to_work_area(window, window_width, window_height)
    window.set_resizable(False)
    window.show_and_focus()

    app.run()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
