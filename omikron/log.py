import queue

LOG_LINE_LENGTH = 35

class OmikronLog:
    log_queue = queue.Queue()

    def log(message:str):
        if len(message) <= LOG_LINE_LENGTH:
            OmikronLog.log_queue.put(message)
        else:
            for index in range(LOG_LINE_LENGTH, -1, -1):
                if message[index] == " ":
                    OmikronLog.log(message[:index])
                    OmikronLog.log(message[index+1:])
                    break
            else:
                OmikronLog.log(message[:LOG_LINE_LENGTH])
                OmikronLog.log(message[LOG_LINE_LENGTH:])

    def error(message:str):
        OmikronLog.log("[오류] " + message)

    def warning(message:str):
        OmikronLog.log("[경고] " + message)
