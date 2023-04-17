# Omikron v1.2.0-alpha
import os
import json
import threading
import tkinter as tk
import omikrondb as odb

config = json.load(open('config.json', encoding='UTF8'))

class GUI():
    def __init__(self, ui):
        self.ui = ui
        self.width = 320
        self.height = 435 # button +25
        self.x = int((self.ui.winfo_screenwidth()/4) - (self.width/2))
        self.y = int((self.ui.winfo_screenheight()/2) - (self.height/2))
        self.ui.geometry(f'{self.width}x{self.height}+{self.x}+{self.y}')
        self.ui.title('Omikron')
        self.ui.resizable(False, False)

        tk.Label(self.ui, text='Omikron 데이터 프로그램').pack()
        self.scroll = tk.Scrollbar(self.ui, orient='vertical')
        self.log = tk.Listbox(self.ui, yscrollcommand=self.scroll.set, width=51, height=5)
        self.scroll.config(command=self.log.yview)
        self.log.pack()
        
        tk.Label(self.ui, text='< 기수 변경 관련 >').pack()

        self.class_info_button = tk.Button(self.ui, text='반 정보 기록 양식 생성', width=40, command=lambda: self.class_info_thread())
        self.class_info_button.pack()
        if os.path.isfile('반 정보.xlsx'):
            self.class_info_button['state'] = tk.DISABLED

        self.student_info_button = tk.Button(self.ui, text='학생 정보 기록 양식 생성', width=40, command=lambda: self.student_info_thread())
        self.student_info_button.pack()
        if os.path.isfile('학생 정보.xlsx'):
            self.student_info_button['state'] = tk.DISABLED

        self.make_data_file_button = tk.Button(self.ui, text='데이터 파일 생성', width=40, command=lambda: self.make_data_file_thread())
        self.make_data_file_button.pack()
        if os.path.isfile('./data/' + config['dataFileName'] + '.xlsx'):
            self.make_data_file_button['state'] = tk.DISABLED

        self.update_class_button = tk.Button(self.ui, text='반 업데이트', width=40, command=lambda: self.update_class_thread())
        self.update_class_button.pack()

        tk.Label(self.ui, text='\n< 데이터 저장 및 문자 전송 >').pack()

        self.make_data_form_button = tk.Button(self.ui, text='데일리 테스트 기록 양식 생성', width=40, command=lambda: self.mkae_data_form_thread())
        self.make_data_form_button.pack()

        self.save_data_button = tk.Button(self.ui, text='데이터 엑셀 파일에 저장', width=40, command=lambda: self.save_data_thread())
        self.save_data_button.pack()
        
        self.send_message_button = tk.Button(self.ui, text='시험 결과 전송', width=40, command=lambda: self.send_message_thread())
        self.send_message_button.pack()

        tk.Label(self.ui, text='\n< 데이터 관리 >').pack()

        self.apply_color_button = tk.Button(self.ui, text='데이터 엑셀 파일 조건부 서식 재지정', width=40, command=lambda: odb.apply_color(self))
        self.apply_color_button.pack()

        self.student_menagement_button = tk.Button(self.ui, text='신규 등록 / 퇴원 관리', width=40, command=None)
        self.student_menagement_button.pack()

    def append_log(self, msg:str):
        self.log.insert(tk.END, msg)
        self.log.update()
        self.log.see(tk.END)

    def class_info_thread(self):
        self.class_info_button['state'] = tk.DISABLED
        thread = threading.Thread(target=lambda: odb.class_info(self))
        thread.daemon = True
        thread.start()

    def make_data_file_thread(self):
        self.make_data_file_button['state'] = tk.DISABLED
        thread = threading.Thread(target=lambda: odb.make_data_file(self))
        thread.daemon = True
        thread.start()

    def save_data_thread(self):
        thread = threading.Thread(target=lambda: odb.save_data(self))
        thread.daemon = True
        thread.start()

    def mkae_data_form_thread(self):
        thread = threading.Thread(target=lambda: odb.make_data_form(self))
        thread.daemon = True
        thread.start()

    def send_message_thread(self):
        thread = threading.Thread(target=lambda: odb.send_message(self))
        thread.daemon = True
        thread.start()

    def student_info_thread(self):
        self.student_info_button['state'] = tk.DISABLED
        thread = threading.Thread(target=lambda: odb.student_info(self))
        thread.daemon = True
        thread.start()

    def update_class_thread(self):
        self.update_class_button['state'] = tk.DISABLED
        thread = threading.Thread(target=lambda: odb.update_class(self))
        thread.daemon = True
        thread.start()

ui = tk.Tk()
gui = GUI(ui)
ui.mainloop()
