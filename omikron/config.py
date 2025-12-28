import os
import json

from omikron.errorui import no_config_file_error, corrupted_config_file_error
from omikron.exception import FileOpenException

def _load_config():
    try:
        with open("./config.json", encoding="UTF8") as f:
            return json.load(f)
    except Exception:
        no_config_file_error()

def _save_config(config: dict):
    with open("./config.json", "w", encoding="UTF8") as f:
        json.dump(config, f, ensure_ascii=False, indent="    ")

config = _load_config()

try:
    DATA_FILE_NAME                  = config["dataFileName"]
    URL                             = config["url"]
    TEST_RESULT_MESSAGE             = config["dailyTest"]
    MAKEUP_TEST_NO_SCHEDULE_MESSAGE = config["makeupTest"]
    MAKEUP_TEST_SCHEDULE_MESSAGE    = config["makeupTestDate"]
    DATA_DIR                        = config["dataDir"]
except Exception:
    corrupted_config_file_error()

if not os.path.isdir(DATA_DIR):
    DATA_DIR = config["dataDir"] = "."
    _save_config(config)

def change_data_file_name(new_filename: str):
    global config, DATA_FILE_NAME, DATA_DIR
    try:
        os.rename(f"{DATA_DIR}/data/{DATA_FILE_NAME}.xlsx", f"{DATA_DIR}/data/{new_filename}.xlsx")
    except Exception:
        raise FileOpenException(f"{DATA_FILE_NAME} 파일을 닫은 뒤 다시 시도해주세요")
    DATA_FILE_NAME = config["dataFileName"] = new_filename
    _save_config(config)

def change_data_path(dir_path:str):
    global config, DATA_DIR
    
    DATA_DIR = config["dataDir"] = dir_path

    _save_config(config)

    if not os.path.exists(f"{DATA_DIR}/data"):
        os.makedirs(f"{DATA_DIR}/data")
    if not os.path.exists(f"{DATA_DIR}/data/backup"):
        os.makedirs(f"{DATA_DIR}/data/backup")

def change_data_file_name_by_select(new_filename:str):
    global config, DATA_FILE_NAME

    DATA_FILE_NAME = config["dataFileName"] = new_filename

    _save_config(config)
