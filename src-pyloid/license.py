import base64
import json
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
import tkinter as tk

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
from cryptography.hazmat.primitives.serialization import load_pem_public_key


# NOTE: Replace this with your real public key (PEM).
PUBLIC_KEY_PEM = """-----BEGIN PUBLIC KEY-----
MCowBQYDK2VwAyEAxU+Q4YXt2w4xyfp7k6VSmI9IsnpallsGw6Hfrr4tpjo=
-----END PUBLIC KEY-----"""

LICENSE_FILENAME = "license.json"
ALLOWED_PRODUCTS = {"tdm", "TestDataManagement"}


@dataclass
class LicenseData:
    name: str
    license_id: str
    issued_at: str
    expires_at: str | None
    product: str


def _app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd().resolve()


def _license_path() -> Path:
    return _app_base_dir() / LICENSE_FILENAME


def _load_license_file() -> dict:
    path = _license_path()
    if not path.exists():
        raise FileNotFoundError(f"Missing {LICENSE_FILENAME} at {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def _canonical_payload(data: dict) -> bytes:
    payload = {
        "name": data.get("name", ""),
        "license_id": data.get("license_id", ""),
        "issued_at": data.get("issued_at", ""),
        "expires_at": data.get("expires_at", ""),
        "product": data.get("product", ""),
    }
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode("utf-8")


def _parse_license(data: dict) -> LicenseData:
    return LicenseData(
        name=str(data.get("name", "")).strip(),
        license_id=str(data.get("license_id", "")).strip(),
        issued_at=str(data.get("issued_at", "")).strip(),
        expires_at=(str(data.get("expires_at")).strip() if data.get("expires_at") else None),
        product=str(data.get("product", "")).strip(),
    )


def _verify_signature(data: dict) -> None:
    public_key_in_file = (data.get("public_key") or "").strip()
    if public_key_in_file != PUBLIC_KEY_PEM.strip():
        raise ValueError("Public key mismatch.")

    signature_b64 = (data.get("signature") or "").strip()
    if not signature_b64:
        raise ValueError("Missing signature.")

    signature = base64.b64decode(signature_b64)
    public_key = load_pem_public_key(PUBLIC_KEY_PEM.encode("utf-8"))
    if not isinstance(public_key, Ed25519PublicKey):
        raise ValueError("Unsupported public key type.")

    payload = _canonical_payload(data)
    public_key.verify(signature, payload)


def _check_expiry(license_data: LicenseData) -> None:
    if not license_data.expires_at:
        return
    try:
        expires = datetime.fromisoformat(license_data.expires_at)
    except ValueError as exc:
        raise ValueError("Invalid expires_at format.") from exc
    now = datetime.now(timezone.utc) if expires.tzinfo else datetime.now()
    if now > expires:
        raise ValueError("License has expired.")


def verify_license() -> tuple[bool, str]:
    try:
        data = _load_license_file()
        _verify_signature(data)
        license_data = _parse_license(data)
        _check_expiry(license_data)
        if license_data.product and license_data.product not in ALLOWED_PRODUCTS:
            raise ValueError("Product mismatch.")
        return True, "OK"
    except Exception as exc:  # noqa: BLE001
        return False, str(exc)


def _show_license_error(message: str) -> None:
    ui = tk.Tk()

    width = 420
    height = 160
    x = int((ui.winfo_screenwidth() / 2) - (width / 2))
    y = int((ui.winfo_screenheight() / 2) - (height / 2))
    ui.geometry(f"{width}x{height}+{x}+{y}")

    ui.title("tdm")
    ui.resizable(False, False)

    tk.Label(ui).pack()
    tk.Label(ui, text="라이선스 검증에 실패했습니다.").pack()
    tk.Label(ui, text="프로그램 실행 위치에 'license.json' 파일이 필요합니다.").pack()
    tk.Label(ui, text=f"상세: {message}").pack()
    tk.Label(ui).pack()

    def _exit_now() -> None:
        try:
            ui.destroy()
        finally:
            sys.exit(1)

    ui.protocol("WM_DELETE_WINDOW", _exit_now)
    button = tk.Button(ui, cursor="hand2", text="확인", width=15, command=_exit_now)
    button.pack()

    ui.mainloop()


def verify_license_or_exit() -> None:
    ok, message = verify_license()
    if not ok:
        _show_license_error(message)
        sys.exit(1)
