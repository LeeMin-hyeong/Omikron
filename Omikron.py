import tkinter as tk
import OmikronDB as odb

class GUI():
    def __init__(self, ui):
        self.ui = ui
        self.ui.geometry('300x220')
        self.ui.title('Omikron')
        self.ui.resizable(False, False)

        self.scroll = tk.Scrollbar(self.ui, orient='vertical')
        self.log = tk.Listbox(self.ui, yscrollcommand=self.scroll.set, width=41, height=5)
        self.scroll.config(command=self.log.yview)
        self.log.pack()


        self.appendLog('Omikron 데이터 프로그램')

        classInfoButton = tk.Button(self.ui, text='반 정보 기록 양식 생성', width=40, command=lambda: odb.classInfo(self))
        classInfoButton.pack()
        makeDataFileButton = tk.Button(self.ui, text='데이터 파일 생성', width=40, command=lambda: odb.makeDataFile(self))
        makeDataFileButton.pack()
        makeDataFormButton = tk.Button(self.ui, text='데일리 테스트 기록 양식 생성', width=40, command=lambda: odb.makeDataForm(self))
        makeDataFormButton.pack()
        saveDataButton = tk.Button(self.ui, text='데이터 저장', width=40, command=lambda: odb.saveData(self))
        saveDataButton.pack()
        sendMessageButton = tk.Button(self.ui, text='시험 결과 전송(재시 미지원)', width=40, command=lambda: odb.sendMessage(self))
        sendMessageButton.pack()

    def appendLog(self, msg):
        self.log.insert(tk.END, msg)
        self.log.update()
        self.log.see(tk.END)

ui = tk.Tk()
gui = GUI(ui)
ui.mainloop()
