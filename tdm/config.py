import json
import os
from pathlib import Path

from tdm.domain.errors import FileOpenException


CONFIG_PATH = Path("./config.json")
REQUIRED_KEYS = (
    "dataFileName",
    "dataDir",
    "url",
    "dailyTest",
    "makeupTest",
    "makeupTestDate",
    "termsAccepted",
    "noticeSeenId",
)


def _default_config() -> dict:
    return {
        "dataFileName": "",
        "dataDir": "",
        "url": "",
        "dailyTest": "",
        "makeupTest": "",
        "makeupTestDate": "",
        "termsAccepted": False,
        "noticeSeenId": "",
    }


def _normalize_config(raw: dict) -> dict:
    normalized = _default_config()
    for key in REQUIRED_KEYS:
        value = raw.get(key, "")
        if key == "termsAccepted":
            normalized[key] = bool(value)
        elif key == "noticeSeenId":
            normalized[key] = value if isinstance(value, str) else str(value)
        else:
            normalized[key] = value if isinstance(value, str) else str(value)
    return normalized


def _load_config() -> tuple[dict, bool]:
    if not CONFIG_PATH.exists():
        return _default_config(), False

    try:
        with CONFIG_PATH.open(encoding="utf-8") as f:
            raw = json.load(f)
        if not isinstance(raw, dict):
            return _default_config(), False
        return _normalize_config(raw), True
    except Exception:
        return _default_config(), False


def _save_config(cfg: dict) -> None:
    with CONFIG_PATH.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=4)


def _sync_runtime_values() -> None:
    global DATA_FILE_NAME, URL, TEST_RESULT_MESSAGE
    global MAKEUP_TEST_NO_SCHEDULE_MESSAGE, MAKEUP_TEST_SCHEDULE_MESSAGE
    global DATA_DIR, DATA_DIR_VALID, TERMS_ACCEPTED, NOTICE_SEEN_ID

    DATA_FILE_NAME = config.get("dataFileName", "").strip()
    URL = config.get("url", "").strip()
    TEST_RESULT_MESSAGE = config.get("dailyTest", "")
    MAKEUP_TEST_NO_SCHEDULE_MESSAGE = config.get("makeupTest", "")
    MAKEUP_TEST_SCHEDULE_MESSAGE = config.get("makeupTestDate", "")
    DATA_DIR = config.get("dataDir", "").strip()
    DATA_DIR_VALID = bool(DATA_DIR) and os.path.isdir(DATA_DIR)
    TERMS_ACCEPTED = bool(config.get("termsAccepted", False))
    NOTICE_SEEN_ID = config.get("noticeSeenId", "").strip()


def _ensure_data_directories() -> None:
    if not DATA_DIR:
        return
    os.makedirs(f"{DATA_DIR}/data", exist_ok=True)
    os.makedirs(f"{DATA_DIR}/data/backup", exist_ok=True)


config, CONFIG_READY = _load_config()
CONFIG_EXISTS = CONFIG_PATH.exists()
_sync_runtime_values()


def is_initialized() -> bool:
    return (
        CONFIG_READY
        and bool(URL)
        and bool(DATA_DIR)
        and bool(DATA_FILE_NAME)
        and bool(TEST_RESULT_MESSAGE)
        and bool(MAKEUP_TEST_NO_SCHEDULE_MESSAGE)
        and bool(MAKEUP_TEST_SCHEDULE_MESSAGE)
        and DATA_DIR_VALID
    )


def is_terms_accepted() -> bool:
    return TERMS_ACCEPTED


def get_notice_seen_id() -> str:
    return NOTICE_SEEN_ID


def accept_terms() -> None:
    global TERMS_ACCEPTED, CONFIG_READY
    TERMS_ACCEPTED = True
    config["termsAccepted"] = True
    CONFIG_READY = True
    _save_config(config)


def set_notice_seen_id(notice_id: str) -> None:
    global NOTICE_SEEN_ID, CONFIG_READY
    NOTICE_SEEN_ID = (notice_id or "").strip()
    config["noticeSeenId"] = NOTICE_SEEN_ID
    CONFIG_READY = True
    _save_config(config)


def initialize_config(
    url: str,
    data_dir: str,
    data_file_name: str,
    daily_test_message: str,
    makeup_test_message: str,
    makeup_test_date_message: str,
) -> None:
    global config, CONFIG_READY, CONFIG_EXISTS

    config = _default_config()
    config["url"] = (url or "").strip()
    config["dataDir"] = (data_dir or "").strip()
    config["dataFileName"] = (data_file_name or "").strip()
    config["dailyTest"] = daily_test_message or ""
    config["makeupTest"] = makeup_test_message or ""
    config["makeupTestDate"] = makeup_test_date_message or ""
    config["termsAccepted"] = False

    _save_config(config)
    CONFIG_EXISTS = True
    CONFIG_READY = True
    _sync_runtime_values()
    _ensure_data_directories()


def update_message_templates(
    url: str,
    daily_test_message: str,
    makeup_test_message: str,
    makeup_test_date_message: str,
) -> None:
    global CONFIG_READY

    config["url"] = (url or "").strip()
    config["dailyTest"] = daily_test_message or ""
    config["makeupTest"] = makeup_test_message or ""
    config["makeupTestDate"] = makeup_test_date_message or ""
    CONFIG_READY = True
    _save_config(config)
    _sync_runtime_values()


def change_data_file_name(new_filename: str) -> None:
    global DATA_FILE_NAME
    try:
        os.rename(f"{DATA_DIR}/data/{DATA_FILE_NAME}.xlsx", f"{DATA_DIR}/data/{new_filename}.xlsx")
        DATA_FILE_NAME = config["dataFileName"] = new_filename
        _save_config(config)
    except FileExistsError:
        raise FileExistsError("A file with the same name already exists.")
    except PermissionError:
        raise FileOpenException(f"Cannot rename file while open: {DATA_FILE_NAME}.xlsx")


def change_data_path(dir_path: str) -> None:
    global DATA_DIR, DATA_DIR_VALID, CONFIG_READY

    DATA_DIR = config["dataDir"] = dir_path
    DATA_DIR_VALID = True
    CONFIG_READY = True
    _save_config(config)
    _ensure_data_directories()


def change_data_file_name_by_select(new_filename: str) -> None:
    global DATA_FILE_NAME, CONFIG_READY

    DATA_FILE_NAME = config["dataFileName"] = new_filename
    CONFIG_READY = True
    _save_config(config)
