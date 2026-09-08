"""Updater operations without UI; all installation writes are recoverable."""

import ctypes
import hashlib
import json
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path, PureWindowsPath
import re
import shutil
import stat
import subprocess
import tempfile
import time
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
import uuid
import zipfile


GITHUB_OWNER = "LeeMin-hyeong"
GITHUB_REPO = "TestDataManagement"
ASSET_NAME = "tdm-win.zip"
HTTP_TIMEOUT = 30
MAX_DOWNLOAD_BYTES = 2 * 1024**3
MAX_EXTRACT_BYTES = 4 * 1024**3
PAYLOAD = ("main.exe", "_internal", "version.txt")
LOGGER = logging.getLogger("tdm.updater")
LOG_PATH = None


class UpdateError(RuntimeError):
    pass


class InstallError(UpdateError):
    def __init__(self, message, *, rollback_ok, backup_dir):
        super().__init__(message)
        self.rollback_ok = rollback_ok
        self.backup_dir = backup_dir


def configure_logging():
    global LOG_PATH
    LOGGER.setLevel(logging.INFO)
    LOGGER.propagate = False
    if LOGGER.handlers:
        return LOG_PATH
    candidates = [Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir())) / "tdm" / "logs",
                  Path(tempfile.gettempdir()) / "tdm-updater-logs"]
    for directory in candidates:
        try:
            directory.mkdir(parents=True, exist_ok=True)
            path = directory / "updater.log"
            handler = RotatingFileHandler(path, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
            handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
            LOGGER.addHandler(handler)
            LOG_PATH = path
            break
        except OSError:
            continue
    if not LOGGER.handlers:
        LOGGER.addHandler(logging.NullHandler())
    return LOG_PATH


def describe_error(exc):
    if isinstance(exc, HTTPError):
        return f"업데이트 서버 요청이 실패했습니다 (HTTP {exc.code}). 잠시 후 다시 실행해 주세요."
    if isinstance(exc, (URLError, TimeoutError)):
        return "업데이트 서버에 연결하지 못했습니다. 인터넷 연결을 확인한 후 다시 실행해 주세요."
    if isinstance(exc, PermissionError):
        return "파일 접근이 거부되었습니다. 실행 중인 tdm을 종료하고 설치 폴더의 쓰기 권한을 확인해 주세요."
    if isinstance(exc, zipfile.BadZipFile):
        return "업데이트 압축 파일이 손상되었습니다. 다시 실행하여 새로 다운로드해 주세요."
    if isinstance(exc, OSError) and exc.errno == 28:
        return "디스크 공간이 부족합니다. 설치 드라이브의 여유 공간을 확보해 주세요."
    return str(exc) or type(exc).__name__


class UpdaterLock:
    """Windows releases this named mutex even if the updater crashes."""
    def __init__(self, root):
        self.handle = None
        self.root = Path(root).resolve()

    def acquire(self):
        if os.name != "nt":
            raise UpdateError("이 업데이터는 Windows에서 실행해 주세요.")
        from ctypes import wintypes
        self.kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        self.kernel.CreateMutexW.argtypes = [ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR]
        self.kernel.CreateMutexW.restype = wintypes.HANDLE
        self.kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        self.kernel.CloseHandle.restype = wintypes.BOOL
        digest = hashlib.sha256(str(self.root).casefold().encode("utf-8")).hexdigest()
        self.handle = self.kernel.CreateMutexW(None, False, "Global\\tdm-updater-" + digest)
        error = ctypes.get_last_error()
        if not self.handle:
            raise ctypes.WinError(error)
        if error == 183:  # ERROR_ALREADY_EXISTS
            self.close()
            raise UpdateError("이 설치 폴더의 업데이터가 이미 실행 중입니다. 기존 업데이트 창을 확인해 주세요.")

    def close(self):
        if self.handle:
            self.kernel.CloseHandle(self.handle)
            self.handle = None


def is_main_running(root):
    """Compare executable paths, rather than unrelated processes named main.exe."""
    if os.name != "nt":
        return False
    from ctypes import wintypes

    class ProcessEntry(ctypes.Structure):
        _fields_ = [
            ("dwSize", wintypes.DWORD), ("cntUsage", wintypes.DWORD),
            ("th32ProcessID", wintypes.DWORD), ("th32DefaultHeapID", ctypes.c_size_t),
            ("th32ModuleID", wintypes.DWORD), ("cntThreads", wintypes.DWORD),
            ("th32ParentProcessID", wintypes.DWORD), ("pcPriClassBase", wintypes.LONG),
            ("dwFlags", wintypes.DWORD), ("szExeFile", wintypes.WCHAR * 260),
        ]

    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
    kernel.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    for function in (kernel.Process32FirstW, kernel.Process32NextW):
        function.argtypes = [wintypes.HANDLE, ctypes.POINTER(ProcessEntry)]
        function.restype = wintypes.BOOL
    kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel.OpenProcess.restype = wintypes.HANDLE
    kernel.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR,
                                                ctypes.POINTER(wintypes.DWORD)]
    kernel.QueryFullProcessImageNameW.restype = wintypes.BOOL
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel.CloseHandle.restype = wintypes.BOOL
    expected = os.path.normcase(str((Path(root) / "main.exe").resolve()))
    snapshot = kernel.CreateToolhelp32Snapshot(0x2, 0)  # TH32CS_SNAPPROCESS
    if snapshot == ctypes.c_void_p(-1).value:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        entry = ProcessEntry()
        entry.dwSize = ctypes.sizeof(entry)
        available = kernel.Process32FirstW(snapshot, ctypes.byref(entry))
        while available:
            if entry.szExeFile.casefold() == "main.exe":
                handle = kernel.OpenProcess(0x1000, False, entry.th32ProcessID)
                if not handle:
                    if ctypes.get_last_error() != 87:  # A process may have exited.
                        raise UpdateError("실행 중인 main.exe의 위치를 확인하지 못했습니다. 해당 프로그램을 종료한 후 다시 실행해 주세요.")
                else:
                    try:
                        buffer = ctypes.create_unicode_buffer(32768)
                        length = wintypes.DWORD(len(buffer))
                        if not kernel.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(length)):
                            raise UpdateError("프로그램 실행 상태를 확인하지 못했습니다. main.exe를 종료한 후 다시 실행해 주세요.")
                        if os.path.normcase(str(Path(buffer.value).resolve())) == expected:
                            return True
                    finally:
                        kernel.CloseHandle(handle)
            available = kernel.Process32NextW(snapshot, ctypes.byref(entry))
        if ctypes.get_last_error() != 18:  # ERROR_NO_MORE_FILES
            raise ctypes.WinError(ctypes.get_last_error())
        return False
    finally:
        kernel.CloseHandle(snapshot)


def parse_semver(version):
    if not isinstance(version, str) or not re.fullmatch(r"[vV]?\d+(?:\.\d+)*", version.strip()):
        raise UpdateError(f"버전 정보가 올바르지 않습니다: {version!r}")
    return [int(part) for part in version.strip().lstrip("vV").split(".")]


def cmp_semver(a, b):
    left, right = parse_semver(a), parse_semver(b)
    length = max(len(left), len(right))
    left += [0] * (length - len(left))
    right += [0] * (length - len(right))
    return (left > right) - (left < right)


def read_local_version(root):
    path = Path(root) / "version.txt"
    if not path.exists():
        return "0.0.0"
    value = path.read_text(encoding="utf-8-sig").strip().lstrip("vV")
    parse_semver(value)
    return value


def gh_get(url):
    request = Request(url, headers={"User-Agent": "tdm-updater"})
    with urlopen(request, timeout=HTTP_TIMEOUT) as response:
        data = response.read(2_000_001)
    if len(data) > 2_000_000:
        raise UpdateError("업데이트 서버 응답이 너무 큽니다.")
    return data


def fetch_latest_zip_asset():
    data = json.loads(gh_get(f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"))
    if not isinstance(data, dict) or not isinstance(data.get("assets"), list):
        raise UpdateError("업데이트 서버 응답 형식이 올바르지 않습니다.")
    version = data.get("tag_name") or data.get("name")
    parse_semver(version)
    assets = [a for a in data["assets"] if isinstance(a, dict) and a.get("name") == ASSET_NAME]
    checksums = [a for a in data["assets"] if isinstance(a, dict) and a.get("name") == ASSET_NAME + ".sha256"]
    if len(assets) != 1 or len(checksums) > 1:
        raise UpdateError("릴리스에서 올바른 tdm-win.zip 업데이트 파일을 찾지 못했습니다.")
    for asset in assets + checksums:
        if not isinstance(asset.get("browser_download_url"), str) or not asset["browser_download_url"].startswith("https://"):
            raise UpdateError("업데이트 다운로드 주소가 올바르지 않습니다.")
    return version.strip().lstrip("vV"), assets[0], checksums[0] if checksums else None


def download_asset(asset, destination):
    destination = Path(destination)
    partial = destination.with_suffix(".part")
    expected_size = asset.get("size")
    if expected_size is not None and (not isinstance(expected_size, int) or not 0 < expected_size <= MAX_DOWNLOAD_BYTES):
        raise UpdateError("업데이트 파일 크기가 올바르지 않습니다.")
    request = Request(asset["browser_download_url"], headers={"User-Agent": "tdm-updater"})
    try:
        with urlopen(request, timeout=HTTP_TIMEOUT) as response, partial.open("wb") as output:
            length = response.headers.get("Content-Length")
            received = 0
            deadline = time.monotonic() + 600
            while True:
                chunk = response.read(1024 * 1024)
                if not chunk:
                    break
                received += len(chunk)
                if received > MAX_DOWNLOAD_BYTES or time.monotonic() > deadline:
                    raise UpdateError("업데이트 다운로드가 허용 크기 또는 제한 시간을 초과했습니다.")
                output.write(chunk)
            if not received or (length is not None and received != int(length)) or (expected_size is not None and received != expected_size):
                raise UpdateError("업데이트 파일 다운로드가 완료되지 않았습니다. 다시 실행해 주세요.")
        partial.replace(destination)
        return destination
    finally:
        if partial.exists():
            try:
                partial.unlink()
            except OSError:
                LOGGER.warning("Partial download cleanup failed: %s", partial, exc_info=True)


def verify_sha256(path, text):
    parts = text.strip().split()
    if not parts or not re.fullmatch(r"[0-9a-fA-F]{64}", parts[0]):
        raise UpdateError("업데이트 체크섬 형식이 올바르지 않습니다.")
    digest = hashlib.sha256()
    with Path(path).open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    if digest.hexdigest() != parts[0].lower():
        raise UpdateError("업데이트 파일 무결성 검증에 실패했습니다. 다시 다운로드해 주세요.")


def safe_child(root, name):
    root = Path(root).resolve()
    child = root / name
    # Never delete/move outside the explicitly selected installation or staging root.
    if not child.resolve().is_relative_to(root) or child.resolve() == root:
        raise UpdateError(f"안전하지 않은 파일 경로입니다: {child}")
    current = child
    while current != root:
        if current.exists() or current.is_symlink():
            attributes = getattr(current.lstat(), "st_file_attributes", 0)
            if current.is_symlink() or attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT:
                raise UpdateError(f"연결된 파일/폴더는 업데이트할 수 없습니다: {current}")
        current = current.parent
    return child


def remove_child(root, name):
    child = safe_child(root, name)
    if child.is_dir():
        shutil.rmtree(child)
    elif child.exists():
        child.unlink()


def safe_extract_zip(zip_path, destination):
    destination = Path(destination).resolve()
    with zipfile.ZipFile(zip_path) as archive:
        seen = set()
        total = 0
        members = archive.infolist()
        if len(members) > 100_000:
            raise UpdateError("압축 파일에 파일이 너무 많습니다.")
        for member in members:
            name = member.filename.rstrip("/")
            parts = name.split("/")
            mode = member.external_attr >> 16
            if (not name or "\\" in name or PureWindowsPath(name).drive
                    or any(p in ("", ".", "..") or ":" in p or p.endswith((" ", "."))
                           or PureWindowsPath(p).is_reserved() for p in parts)
                    or stat.S_ISLNK(mode) or member.flag_bits & 1):
                raise UpdateError(f"압축 파일에 안전하지 않은 경로가 있습니다: {member.filename}")
            if name.casefold() in seen:
                raise UpdateError(f"압축 파일에 중복 경로가 있습니다: {name}")
            seen.add(name.casefold())
            safe_child(destination, name)
            total += member.file_size
            if total > MAX_EXTRACT_BYTES:
                raise UpdateError("압축 해제할 파일 크기가 너무 큽니다.")
        if shutil.disk_usage(destination.parent).free < total + 64 * 1024**2:
            raise UpdateError("업데이트 압축을 해제할 디스크 공간이 부족합니다.")
        archive.extractall(destination)


def validate_payload(root):
    root = Path(root).resolve()
    required = ("main.exe", "_internal/dist-front/index.html", "_internal/PySide6/QtWebEngineProcess.exe")
    for name in required:
        path = safe_child(root, name)
        if not path.is_file() or not path.stat().st_size:
            raise UpdateError(f"프로그램 필수 파일이 없거나 비어 있습니다: {path}")
    internal = safe_child(root, "_internal")
    dlls = [p for p in internal.glob("python3*.dll") if p.name != "python3.dll"]
    if not dlls or not any(safe_child(internal, p.name).is_file() and p.stat().st_size for p in dlls):
        raise UpdateError(f"Python 실행 라이브러리가 없습니다: {internal}")


class Transaction:
    """Journal original existence before the first rename; keep backups until launch."""
    def __init__(self, root, new_root, version):
        self.root = Path(root).resolve()
        self.new_root = Path(new_root).resolve()
        self.version = version
        self.directory = safe_child(self.root, ".update_tmp/backup/" + uuid.uuid4().hex)
        self.record = None

    def _write_record(self):
        temporary = safe_child(self.directory, "transaction.tmp")
        with temporary.open("w", encoding="utf-8") as output:
            json.dump(self.record, output, ensure_ascii=False)
            output.flush()
            os.fsync(output.fileno())
        temporary.replace(safe_child(self.directory, "transaction.json"))

    def install(self):
        validate_payload(self.new_root)
        parse_semver(self.version)
        staged_version = safe_child(self.new_root, "version.txt")
        if staged_version.exists() and cmp_semver(staged_version.read_text(encoding="utf-8-sig").strip(), self.version):
            raise UpdateError("릴리스 버전과 압축 파일의 version.txt가 일치하지 않습니다.")
        staged_version.write_text(self.version.strip().lstrip("vV"), encoding="utf-8")
        originals = {}
        for name in PAYLOAD:
            path = safe_child(self.root, name)
            if path.exists() and (path.is_dir() != (name == "_internal")):
                raise UpdateError(f"프로그램 경로의 파일 형식이 올바르지 않습니다: {path}")
            originals[name] = path.exists()
        self.directory.mkdir(parents=True, exist_ok=False)
        self.record = {"schema": 1, "phase": "pending", "originals": originals}
        self._write_record()
        try:
            for name in PAYLOAD:
                if originals[name]:
                    safe_child(self.root, name).replace(safe_child(self.directory, name))
            for name in PAYLOAD:
                safe_child(self.new_root, name).replace(safe_child(self.root, name))
            validate_payload(self.root)
            return self
        except Exception as exc:
            LOGGER.exception("Installation failed; rolling back")
            try:
                self.rollback()
            except Exception as recovery:
                raise InstallError(
                    f"업데이트 설치 실패: {describe_error(exc)}\n복구도 완료하지 못했습니다: {describe_error(recovery)}",
                    rollback_ok=False, backup_dir=self.directory,
                ) from exc
            raise InstallError(f"업데이트 설치 실패: {describe_error(exc)}\n업데이트 이전 상태로 복구했습니다.",
                               rollback_ok=True, backup_dir=self.directory) from exc

    def rollback(self):
        if not self.record or self.record["phase"] != "pending":
            return
        failures = []
        for name in reversed(PAYLOAD):
            try:
                saved = safe_child(self.directory, name)
                target = safe_child(self.root, name)
                if saved.exists():
                    remove_child(self.root, name)
                    saved.replace(target)
                elif not self.record["originals"][name]:
                    remove_child(self.root, name)
                elif not target.exists():
                    raise UpdateError(f"원본과 백업을 모두 찾지 못했습니다: {name}")
                # A missing backup with an original target means it was never
                # moved, or was already restored before an interrupted rollback.
            except Exception as exc:
                failures.append(f"{name}: {describe_error(exc)}")
                LOGGER.exception("Rollback failed for %s", name)
        if failures:
            raise InstallError("\n".join(failures), rollback_ok=False, backup_dir=self.directory)
        self.record["phase"] = "rolled_back"
        try:
            self._write_record()
        except Exception:
            self.record["phase"] = "pending"
            raise

    def commit(self):
        if self.record["phase"] != "pending":
            raise UpdateError("완료할 수 없는 업데이트 상태입니다.")
        self.record["phase"] = "committed"
        try:
            self._write_record()
        except Exception:
            self.record["phase"] = "pending"
            raise
        try:
            remove_child(self.directory.parent, self.directory.name)
        except Exception as exc:
            LOGGER.warning("Committed backup cleanup failed", exc_info=True)
            return f"업데이트는 적용되었으나 백업 정리를 완료하지 못했습니다.\n{self.directory}\n{describe_error(exc)}"
        return None


def recover_pending(root):
    root = Path(root).resolve()
    directory = safe_child(root, ".update_tmp/backup")
    if not directory.exists():
        return []
    pending = []
    messages = []
    for path in sorted(directory.iterdir()):
        path = safe_child(directory, path.name)
        journal = safe_child(path, "transaction.json") if path.is_dir() else None
        if journal is None or not journal.exists():
            continue
        try:
            record = json.loads(journal.read_text(encoding="utf-8"))
            if (record["schema"] != 1 or record["phase"] not in ("pending", "committed", "rolled_back")
                    or set(record["originals"]) != set(PAYLOAD)
                    or any(type(value) is not bool for value in record["originals"].values())):
                raise ValueError("Invalid journal")
        except Exception as exc:
            raise InstallError(f"이전 업데이트 복구 기록을 읽을 수 없습니다: {journal}",
                               rollback_ok=False, backup_dir=path) from exc
        if record["phase"] == "pending":
            transaction = Transaction(root, root, "0")
            transaction.directory, transaction.record = path, record
            pending.append(transaction)
        else:
            try:
                remove_child(directory, path.name)
            except OSError:
                LOGGER.warning("Old backup cleanup failed: %s", path, exc_info=True)
                messages.append(f"이전 업데이트 백업을 정리하지 못했습니다. 프로그램 파일은 변경하지 않았습니다.\n{path}")
    if len(pending) > 1:
        raise InstallError("완료되지 않은 업데이트 기록이 여러 개입니다. 백업 폴더를 보존하고 관리자에게 문의해 주세요.",
                           rollback_ok=False, backup_dir=directory)
    for transaction in pending:
        try:
            transaction.rollback()
        except Exception as exc:
            raise InstallError(f"중단된 업데이트 복구에 실패했습니다: {describe_error(exc)}",
                               rollback_ok=False, backup_dir=transaction.directory) from exc
        messages.append("이전에 중단된 업데이트를 발견하여 업데이트 이전 상태로 복구했습니다.")
    return messages


def launch_main(root, monitor_seconds=5.0):
    """Detect failed process creation / early exit. Survival is not a UI-ready signal."""
    root = Path(root).resolve()
    validate_payload(root)
    if not (root / "license.json").is_file():
        raise UpdateError(f"라이선스 파일이 없습니다. main.exe와 같은 폴더에 license.json을 배치해 주세요.\n{root}")
    environment = os.environ.copy()
    # A onefile PyInstaller updater must launch main as an independent app.
    environment["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    try:
        process = subprocess.Popen([str(root / "main.exe")], cwd=str(root), env=environment,
                                   close_fds=True)
    except OSError as exc:
        raise UpdateError(f"메인 프로그램을 실행하지 못했습니다.\n{describe_error(exc)}") from exc
    deadline = time.monotonic() + monitor_seconds
    while True:
        code = process.poll()
        if code is not None:
            raise UpdateError(f"메인 프로그램이 시작 직후 종료되었습니다 (종료 코드: {code}).\n"
                              "라이선스 오류창과 프로그램 파일을 확인해 주세요.")
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return process
        time.sleep(min(0.1, remaining))
