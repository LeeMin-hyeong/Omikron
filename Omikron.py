import threading
import tkinter as tk
import OmikronDB as odb

class GUI():
    def __init__(self, ui):
        self.ui = ui
        self.ui.geometry('300x250')
        self.ui.title('Omikron')
        self.ui.resizable(False, False)

        self.scroll = tk.Scrollbar(self.ui, orient='vertical')
        self.log = tk.Listbox(self.ui, yscrollcommand=self.scroll.set, width=41, height=5)
        self.scroll.config(command=self.log.yview)
        self.log.pack()

        self.appendLog('Omikron 데이터 프로그램')

        classInfoButton = tk.Button(self.ui, text='반 정보 기록 양식 생성', width=40, command=lambda: self.classInfoThread())
        classInfoButton.pack()
        classInfoButton = tk.Button(self.ui, text='재시험 정보 생성', width=40, command=lambda: self.makeupTestInfoThread())
        classInfoButton.pack()
        makeDataFileButton = tk.Button(self.ui, text='데이터 파일 생성', width=40, command=lambda: self.makeDataFileThread())
        makeDataFileButton.pack()
        makeDataFormButton = tk.Button(self.ui, text='데일리 테스트 기록 양식 생성', width=40, command=lambda: self.makeDataFormThread())
        makeDataFormButton.pack()
        saveDataButton = tk.Button(self.ui, text='데이터 저장', width=40, command=lambda: odb.saveData(self))
        saveDataButton.pack()
        sendMessageButton = tk.Button(self.ui, text='시험 결과 전송(휴일 미지원)', width=40, command=lambda: self.sendMessageThread())
        sendMessageButton.pack()

    def appendLog(self, msg):
        self.log.insert(tk.END, msg)
        self.log.update()
        self.log.see(tk.END)

    def classInfoThread(self):
        thread = threading.Thread(target=lambda: odb.classInfo(self))
        thread.daemon = True
        thread.start()

    def makeDataFileThread(self):
        thread = threading.Thread(target=lambda: odb.makeDataFile(self))
        thread.daemon = True
        thread.start()

    def makeDataFormThread(self):
        thread = threading.Thread(target=lambda: odb.makeDataForm(self))
        thread.daemon = True
        thread.start()

    def sendMessageThread(self):
        thread = threading.Thread(target=lambda: odb.sendMessage(self))
        thread.daemon = True
        thread.start()

    def makeupTestInfoThread(self):
        thread = threading.Thread(target=lambda: odb.makeupTestInfo(self))
        thread.daemon = True
        thread.start()

ui = tk.Tk()
gui = GUI(ui)
ui.mainloop()
