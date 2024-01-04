import queue

from defs import MESSAGE_INTERFACE_WIDTH

class Message:
    message_queue = queue.Queue()

    def error(message:str):
        pass

    def message(message:str):
        while(len(message) > 0):
            Message.message_queue.put(f"{message[:MESSAGE_INTERFACE_WIDTH]}")
            message = message[MESSAGE_INTERFACE_WIDTH:]
