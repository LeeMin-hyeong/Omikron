import os
import json
import threading
import tkinter as tk
import OmikronDB as odb

config = json.load(open('config.json', encoding='UTF8'))

class GUI():
    def __init__(self, ui):
        self.ui = ui
        self.ui.geometry('300x410+460+460')
        self.ui.title('Omikron')
        self.ui.resizable(False, False)

        tk.Label(self.ui, text='Omikron 데이터 프로그램').pack()
        self.scroll = tk.Scrollbar(self.ui, orient='vertical')
        self.log = tk.Listbox(self.ui, yscrollcommand=self.scroll.set, width=41, height=5)
        self.scroll.config(command=self.log.yview)
        self.log.pack()
        
        tk.Label(self.ui, text='< 기수 변경 관련 >').pack()

        self.classInfoButton = tk.Button(self.ui, text='반 정보 기록 양식 생성', width=40, command=lambda: self.classInfoThread())
        self.classInfoButton.pack()
        if os.path.isfile('반 정보.xlsx'):
            self.classInfoButton['state'] = tk.DISABLED

        self.makeupTestInfoButton = tk.Button(self.ui, text='재시험 정보 기록 양식 생성', width=40, command=lambda: self.makeupTestInfoThread())
        self.makeupTestInfoButton.pack()
        if os.path.isfile('재시험 정보.xlsx'):
            self.makeupTestInfoButton['state'] = tk.DISABLED

        self.makeDataFileButton = tk.Button(self.ui, text='데이터 파일 생성', width=40, command=lambda: self.makeDataFileThread())
        self.makeDataFileButton.pack()
        if os.path.isfile('./data/' + config['dataFileName'] + '.xlsx'):
            self.makeDataFileButton['state'] = tk.DISABLED

        self.updateClassButton = tk.Button(self.ui, text='반 업데이트', width=40, command=lambda: self.updateClassThread())
        self.updateClassButton.pack()

        tk.Label(self.ui, text='\n< 데이터 저장 및 문자 전송 >').pack()

        self.makeDataFormButton = tk.Button(self.ui, text='데일리 테스트 기록 양식 생성', width=40, command=lambda: self.makeDataFormThread())
        self.makeDataFormButton.pack()

        self.saveDataButton = tk.Button(self.ui, text='데이터 엑셀 파일에 저장', width=40, command=lambda: self.saveDataThread())
        self.saveDataButton.pack()
        
        self.sendMessageButton = tk.Button(self.ui, text='시험 결과 전송', width=40, command=lambda: self.sendMessageThread())
        self.sendMessageButton.pack()

        tk.Label(self.ui, text='\n< 데이터 관리 >').pack()

        self.applyColorButton = tk.Button(self.ui, text='데이터 엑셀 파일 조건부 서식 재지정', width=40, command=lambda: odb.applyColor(self))
        self.applyColorButton.pack()

    def appendLog(self, msg):
        self.log.insert(tk.END, msg)
        self.log.update()
        self.log.see(tk.END)

    def classInfoThread(self):
        self.classInfoButton['state'] = tk.DISABLED
        thread = threading.Thread(target=lambda: odb.classInfo(self))
        thread.daemon = True
        thread.start()

    def makeDataFileThread(self):
        self.makeDataFileButton['state'] = tk.DISABLED
        thread = threading.Thread(target=lambda: odb.makeDataFile(self))
        thread.daemon = True
        thread.start()

    def saveDataThread(self):
        thread = threading.Thread(target=lambda: odb.saveData(self))
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
        self.makeupTestInfoButton['state'] = tk.DISABLED
        thread = threading.Thread(target=lambda: odb.makeupTestInfo(self))
        thread.daemon = True
        thread.start()

    def updateClassThread(self):
        self.updateClassButton['state'] = tk.DISABLED
        thread = threading.Thread(target=lambda: odb.updateClass(self))
        thread.daemon = True
        thread.start()

ui = tk.Tk()
gui = GUI(ui)
ui.mainloop()
