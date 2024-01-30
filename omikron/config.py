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