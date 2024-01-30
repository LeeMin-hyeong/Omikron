import sys
import queue

LOG_LINE_LENGTH = 35

class OmikronLog:
    log_queue = queue.Queue()

    def log(message:str):
        if len(message) <= LOG_LINE_LENGTH:
            OmikronLog.log_queue.put(message)
        else:
            for index in range(LOG_LINE_LENGTH, 0, -1):
                if message[index] == " ":
                    OmikronLog.log(message[:index])
                    OmikronLog.log(message[index+1:])
                    break
            else:
                OmikronLog.log(message[:LOG_LINE_LENGTH])
                OmikronLog.log(message[LOG_LINE_LENGTH:])

    def error(message:str):
        message = "[오류] " + message
        OmikronLog.log(message)

    def warning(message:str):
        message = "[경고] " + message
        OmikronLog.log(message)
