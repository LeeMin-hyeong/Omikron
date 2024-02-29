import os
import json

from omikron.errorui import no_config_file_error, corrupted_config_file_error

try:
    config = json.load(open("./config.json", encoding="UTF8"))
except:
    no_config_file_error()

try:
    DATA_FILE_NAME                  = config["dataFileName"]
    URL                             = config["url"]
    TEST_RESULT_MESSAGE             = config["dailyTest"]
    MAKEUP_TEST_NO_SCHEDULE_MESSAGE = config["makeupTest"]
    MAKEUP_TEST_SCHEDULE_MESSAGE    = config["makeupTestDate"]
except:
    corrupted_config_file_error()

def change_data_file_name(new_filename:str):
    global DATA_FILE_NAME

    os.rename(f"./data/{DATA_FILE_NAME}.xlsx", f"./data/{new_filename}.xlsx")

    DATA_FILE_NAME = config["dataFileName"] = new_filename

    json.dump(config, open("./config.json", 'w', encoding="UTF8"), ensure_ascii=False, indent="    ")
