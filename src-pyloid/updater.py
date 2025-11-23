import json
import os
import shutil
import sys
import time
import zipfile
import threading
import hashlib
import subprocess
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError
import tkinter as tk
from tkinter import ttk

# ===================== 사용자 설정 =====================
GITHUB_OWNER = "LeeMin-hyeong"
GITHUB_REPO  = "Omikron"
ASSET_NAME_CONTAINS = "omikron-win.zip"   # 릴리스 ZIP 자산 이름 일부
SHA256_SUFFIX       = ".sha256"          # (선택) 체크섬 파일 suffix
MAIN_EXE_NAME       = "main.exe"
LOCAL_VERSION_FILE  = "version.txt"      # 루트 버전 파일
ZIP_VERSION_FILE    = "version.txt"      # zip 내부 버전 파일 (omikron/version.txt)
LAUNCH_ARGS         = []                 # main.exe 실행 시 전달할 인자
GITHUB_TOKEN        = None               # 필요시 GitHub PAT
HTTP_TIMEOUT        = 30

# ===================== 경로 처리 =====================
if getattr(sys, 'frozen', False):
    ROOT = Path(sys.executable).parent
else:
    ROOT = Path(__file__).parent

def resource_path(relative: str) -> Path:
    """PyInstaller 환경에서도 자원 경로를 올바르게 찾기 위한 헬퍼."""
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative
    return Path(f"{relative}")

MAIN_EXE_PATH  = ROOT / MAIN_EXE_NAME
LOCAL_VER_PATH = ROOT / LOCAL_VERSION_FILE

TMP_DIR       = ROOT / ".update_tmp"
DOWNLOAD_DIR  = TMP_DIR / "downloads"
STAGING_DIR   = TMP_DIR / "staging"
BACKUP_ROOT   = TMP_DIR / "backup"

for d in (TMP_DIR, DOWNLOAD_DIR, STAGING_DIR, BACKUP_ROOT):
    d.mkdir(parents=True, exist_ok=True)

# 루트에 배치될 대상들
PAYLOAD_FILES = [MAIN_EXE_NAME]   # main.exe
PAYLOAD_DIRS  = ["_internal"]     # _internal/

# ===================== 로깅 =====================
def log(msg: str):
    print(f"[updater] {msg}", flush=True)

# ===================== GitHub 통신 =====================
def gh_get(url: str) -> bytes:
    headers = {"User-Agent": "omikron-updater"}
    if GITHUB_TOKEN:
        headers["Authorization"] = f"Bearer {GITHUB_TOKEN}"
    req = Request(url, headers=headers)
    return urlopen(req, timeout=HTTP_TIMEOUT).read()

def fetch_latest_zip_asset():
    """latest 릴리스 정보에서 ZIP 자산과 SHA256 자산 찾기."""
    url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
    data = gh_get(url)
    j = json.loads(data.decode("utf-8"))

    tag = (j.get("tag_name") or j.get("name") or "").lstrip("vV")
    asset = None
    sha_asset = None

    for a in j.get("assets", []):
        name = a.get("name", "")
        if ASSET_NAME_CONTAINS in name and name.endswith(".zip"):
            asset = a
        if name.endswith(SHA256_SUFFIX) and ASSET_NAME_CONTAINS in name.replace(SHA256_SUFFIX, ""):
            sha_asset = a

    return tag, asset, sha_asset

def download_asset(asset, dst: Path) -> Path:
    headers = {"User-Agent": "omikron-updater"}
    if GITHUB_TOKEN:
        headers["Authorization"] = f"Bearer {GITHUB_TOKEN}"
    req = Request(asset["browser_download_url"], headers=headers)
    with urlopen(req, timeout=HTTP_TIMEOUT) as r, open(dst, "wb") as f:
        shutil.copyfileobj(r, f)
    return dst

# ===================== 버전 / 체크섬 =====================
def read_local_version() -> str:
    if LOCAL_VER_PATH.exists():
        return LOCAL_VER_PATH.read_text(encoding="utf-8").strip()
    return "0.0.0"

def parse_semver(v: str):
    return [int(x) for x in v.strip().split(".") if x.isdigit()]

def cmp_semver(a: str, b: str) -> int:
    pa, pb = parse_semver(a), parse_semver(b)
    la, lb = len(pa), len(pb)
    if la < lb:
        pa += [0] * (lb - la)
    elif lb < la:
        pb += [0] * (la - lb)
    return (pa > pb) - (pa < pb)

def sha256_file(p: Path) -> str:
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()

def verify_sha256(file_path: Path, sha_text: str):
    want = sha_text.strip().split()[0].lower()
    got = sha256_file(file_path).lower()
    if want != got:
        raise RuntimeError(f"SHA256 mismatch: expected {want}, got {got}")

# ===================== ZIP / 설치 =====================
def is_main_running() -> bool:
    """Windows에서 main.exe가 이미 떠 있는지 확인."""
    if os.name != "nt":
        return False
    try:
        out = subprocess.check_output(
            ["tasklist", "/FI", f"IMAGENAME eq {MAIN_EXE_NAME}"]
        ).decode("cp949", errors="ignore")
        return MAIN_EXE_NAME.lower() in out.lower()
    except Exception as e:
        log(f"실행 여부 확인 실패(무시): {e}")
        return False

def safe_extract_zip(zip_path: Path, dest_dir: Path):
    with zipfile.ZipFile(zip_path, "r") as zf:
        for m in zf.infolist():
            member = Path(m.filename)
            if member.is_absolute() or ".." in member.parts:
                raise RuntimeError(f"Unsafe path in zip: {m.filename}")
        zf.extractall(dest_dir)

def find_new_root(staging_root: Path) -> Path:
    """staging_root/omikron 을 새 버전 루트로 사용."""
    omikron_dir = staging_root / "omikron"
    if omikron_dir.exists() and omikron_dir.is_dir():
        return omikron_dir
    raise RuntimeError("스테이징에서 omikron 폴더를 찾지 못했습니다.")

def install_new_version(new_root: Path):
    """
    new_root: STAGING_DIR/omikron
    - main.exe   → ROOT/main.exe
    - _internal/ → ROOT/_internal
    - version.txt → ROOT/version.txt
    """
    ts = time.strftime("%Y%m%d-%H%M%S")
    backup_dir = BACKUP_ROOT / f"backup-{ts}"
    backup_dir.mkdir(parents=True, exist_ok=True)

    try:
        # 기존 payload 백업
        for fname in PAYLOAD_FILES:
            target = ROOT / fname
            if target.exists():
                target.replace(backup_dir / fname)

        for dname in PAYLOAD_DIRS:
            target = ROOT / dname
            if target.exists():
                target.replace(backup_dir / dname)

        # 새 파일/폴더 이동
        for fname in PAYLOAD_FILES:
            src = new_root / fname
            if not src.exists():
                raise RuntimeError(f"새 버전에서 {fname} 를 찾을 수 없습니다.")
            dst = ROOT / fname
            src.replace(dst)

        for dname in PAYLOAD_DIRS:
            src_dir = new_root / dname
            if not src_dir.exists():
                raise RuntimeError(f"새 버전에서 {dname}/ 디렉터리를 찾을 수 없습니다.")
            dst_dir = ROOT / dname
            src_dir.replace(dst_dir)

        # version.txt
        zip_ver_path = new_root / ZIP_VERSION_FILE
        if zip_ver_path.exists():
            new_ver = zip_ver_path.read_text(encoding="utf-8").strip()
            LOCAL_VER_PATH.write_text(new_ver, encoding="utf-8")

        # 성공 → 백업 삭제
        shutil.rmtree(backup_dir, ignore_errors=True)

    except Exception as e:
        log(f"설치 중 오류: {e}")
        # 롤백
        try:
            for fname in PAYLOAD_FILES:
                t = ROOT / fname
                if t.exists() and t.is_file():
                    t.unlink()

            for dname in PAYLOAD_DIRS:
                tdir = ROOT / dname
                if tdir.exists():
                    shutil.rmtree(tdir, ignore_errors=True)

            for fname in PAYLOAD_FILES:
                b = backup_dir / fname
                if b.exists():
                    b.replace(ROOT / fname)
            for dname in PAYLOAD_DIRS:
                bdir = backup_dir / dname
                if bdir.exists():
                    bdir.replace(ROOT / dname)
        except Exception as e2:
            log(f"롤백 실패: {e2}")
        raise

def launch_main():
    try:
        subprocess.Popen(
            [str(MAIN_EXE_PATH), *LAUNCH_ARGS],
            cwd=str(ROOT),
            close_fds=True
        )
    except Exception as e:
        log(f"메인 실행 실패: {e}")
    sys.exit(0)

# ===================== 스플래시 UI =====================
class Updater:
    def __init__(self):
        self.root = tk.Tk()

        # === 윈도우 프레임 제거 & 항상 위 ===
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)

        # === 투명 배경 + 둥근 흰 카드 ===
        self.trans_color = "#00FF00"   # 완전 투명으로 사용할 색
        self.card_color  = "#FFFFFF"   # 카드(사각형) 색

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
            radius=20,
            fill=self.card_color
        )

        # === 로고 ===
        self.logo_img = None
        logo_path = resource_path("src/assets/omikron.png")
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
                text="Omikron",
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
        self.status_var = tk.StringVar(value="시작 중…")
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
            "Omikron.Horizontal.TProgressbar",
            troughcolor="#EEEEEE",
            bordercolor="#EEEEEE",
            background="#FF4D4F",
            lightcolor="#FF4D4F",
            darkcolor="#FF4D4F",
        )
        self.progress.configure(style="Omikron.Horizontal.TProgressbar")

        # 카드 안, 아래쪽에 frame 올리기
        self.canvas.create_window(
            self.width // 2,
            self.height - 60,
            window=self.bottom_frame
        )

        # 100ms 후 업데이트 쓰레드 시작
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
        """메인 스레드에서 안전하게 UI 업데이트."""
        def _update():
            self.status_var.set(text)
            if progress is not None:
                self.progress_var.set(progress)
        self.root.after(0, _update)

    # ------------------ 업데이트 플로우 ------------------
    def start_update_thread(self):
        t = threading.Thread(target=self.run_update_flow, daemon=True)
        t.start()

    def run_update_flow(self):
        self.set_status("프로그램 실행 여부 확인 중…", 5)
        if is_main_running():
            log("메인 프로그램이 이미 실행 중입니다. 업데이트를 건너뜁니다.")
            self.set_status("이미 Omikron이 실행 중입니다.\n업데이트는 건너뜁니다.", 100)
            time.sleep(1.5)

            def _close():
                self.root.destroy()
            self.root.after(0, _close)
            return

        # 1. 현재 버전 읽기
        self.set_status("현재 버전 확인 중…", 10)
        current = read_local_version()
        log(f"현재 버전: {current}")

        try:
            # 2. GitHub latest 확인
            self.set_status("서버에서 최신 버전 확인 중…", 25)
            latest, asset, sha_asset = fetch_latest_zip_asset()
            if not asset:
                log("릴리스 ZIP 자산을 찾지 못했습니다.")
                self.set_status("업데이트 정보를 찾지 못했습니다.\n앱을 실행합니다.", 100)
                time.sleep(0.5)
                self.finish_and_launch()
                return

            if cmp_semver(current, latest) >= 0:
                log(f"최신입니다. (latest={latest})")
                self.set_status("이미 최신 버전입니다.\n앱을 실행합니다.", 100)
                time.sleep(0.5)
                self.finish_and_launch()
                return

            log(f"업데이트 필요: {current} → {latest}")

            # 3. ZIP 다운로드
            self.set_status("업데이트 파일 다운로드 중…", 50)
            zip_path = DOWNLOAD_DIR / asset["name"]
            download_asset(asset, zip_path)

            # 4. 체크섬 검증(선택)
            if sha_asset:
                self.set_status("파일 무결성 확인 중…", 65)
                sha_text = gh_get(sha_asset["browser_download_url"]).decode("utf-8")
                verify_sha256(zip_path, sha_text)

            # 5. 스테이징 초기화 및 해제
            self.set_status("업데이트 파일 압축 해제 중…", 80)
            if STAGING_DIR.exists():
                shutil.rmtree(STAGING_DIR, ignore_errors=True)
            STAGING_DIR.mkdir(parents=True, exist_ok=True)
            safe_extract_zip(zip_path, STAGING_DIR)

            new_root = find_new_root(STAGING_DIR)

            # 6. 설치
            self.set_status("업데이트 적용 중…", 90)
            install_new_version(new_root)

            self.set_status("업데이트 완료!\n앱을 실행합니다.", 100)
            time.sleep(0.5)

        except (HTTPError, URLError) as e:
            log(f"네트워크 오류: {e}")
            self.set_status("네트워크 오류가 발생했지만\n앱을 실행합니다.", 100)
            time.sleep(0.5)
        except Exception as e:
            log(f"업데이트 실패: {e}")
            self.set_status("업데이트 실패.\n이전 버전으로 실행합니다.", 100)
            time.sleep(0.5)

        self.finish_and_launch()

    def finish_and_launch(self):
        def _finish():
            self.root.destroy()
            launch_main()
        self.root.after(0, _finish)

    def run(self):
        self.root.mainloop()

# ===================== 엔트리 =====================
if __name__ == "__main__":
    app = Updater()
    app.run()
