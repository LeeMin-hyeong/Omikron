# Omikron v1.2.0-beta5
import json
import queue
import os.path
import pythoncom # only works in Windows
import threading
import tkinter as tk
import tkinter.messagebox
import openpyxl as xl
import win32com.client # only works in Windows

from copy import copy
from omikronconst import *
from tkinter import ttk, filedialog
from datetime import date as DATE, datetime, timedelta
from dateutil.relativedelta import relativedelta
from openpyxl.cell import Cell
from openpyxl.utils.cell import get_column_letter as gcl
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Protection, Side
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from win32process import CREATE_NO_WINDOW # only works in Windows
from webdriver_manager.chrome import ChromeDriverManager

config = json.load(open("./config.json", encoding="UTF8"))
os.environ["WDM_PROGRESS_BAR"] = "0"
try:
    service = Service(ChromeDriverManager().install())
except:
    service = Service(ChromeDriverManager(version=config["webDriverManagerVersion"]).install())
service.creation_flags = CREATE_NO_WINDOW

if not os.path.exists("./data"):
    os.makedirs("./data")
if not os.path.exists("./data/backup"):
    os.makedirs("./data/backup")

class GUI():
    def __init__(self, ui:tk.Tk):
        self.q = queue.Queue()
        self.thread_end_flag = False
        self.ui = ui
        self.width = 320
        self.height = 515 # button +25
        self.x = int((self.ui.winfo_screenwidth()/4) - (self.width/2))
        self.y = int((self.ui.winfo_screenheight()/2) - (self.height/2))
        self.ui.geometry(f"{self.width}x{self.height}+{self.x}+{self.y}")
        self.ui.title("Omikron")
        self.ui.resizable(False, False)

        tk.Label(self.ui, text="Omikron 데이터 프로그램").pack()
        self.scroll = tk.Scrollbar(self.ui, orient="vertical")
        self.log = tk.Listbox(self.ui, yscrollcommand=self.scroll.set, width=51, height=5)
        self.scroll.config(command=self.log.yview)
        self.log.pack()
        
        tk.Label(self.ui, text="< 기수 변경 관련 >").pack()

        self.make_class_info_file_button = tk.Button(self.ui, text="반 정보 기록 양식 생성", width=40, command=lambda: self.make_class_info_file_thread())
        self.make_class_info_file_button.pack()

        self.make_student_info_file_button = tk.Button(self.ui, text="학생 정보 기록 양식 생성", width=40, command=lambda: self.make_student_info_file_thread())
        self.make_student_info_file_button.pack()

        self.make_data_file_button = tk.Button(self.ui, text="데이터 파일 생성", width=40, command=lambda: self.make_data_file_thread())
        self.make_data_file_button.pack()

        self.update_class_button = tk.Button(self.ui, text="반 업데이트", width=40, command=lambda: self.update_class_thread())
        self.update_class_button.pack()

        tk.Label(self.ui, text="\n< 데이터 저장 및 문자 전송 >").pack()

        self.make_data_form_button = tk.Button(self.ui, text="데일리 테스트 기록 양식 생성", width=40, command=lambda: self.make_data_form_thread())
        self.make_data_form_button.pack()

        self.save_data_button = tk.Button(self.ui, text="데이터 엑셀 파일에 저장", width=40, command=lambda: self.save_data_thread())
        self.save_data_button.pack()

        self.send_message_button = tk.Button(self.ui, text="시험 결과 전송", width=40, command=lambda: self.send_message_thread())
        self.send_message_button.pack()

        tk.Label(self.ui, text="\n< 데이터 관리 >").pack()

        self.apply_color_button = tk.Button(self.ui, text="데이터 엑셀 파일 조건부 서식 재지정", width=40, command=lambda: apply_color(self))
        self.apply_color_button.pack()

        tk.Label(self.ui, text="< 학생 관리 >").pack()
        self.add_student_button = tk.Button(self.ui, text="신규생 추가", width=40, command=lambda: self.add_student_thread())
        self.add_student_button.pack()

        self.delete_student_button = tk.Button(self.ui, text="퇴원 처리", width=40, command=lambda: self.delete_student_thread())
        self.delete_student_button.pack()

        self.move_student_button = tk.Button(self.ui, text="학생 반 이동", width=40, command=lambda: self.move_student_thread())
        self.move_student_button.pack()

    # ui
    def thread_log(self):
        try:
            msg = self.q.get(block=False)
        except queue.Empty:
            self.ui.after(100, self.thread_log)
            return
        self.log.insert(tk.END, msg)
        self.log.update()
        self.log.see(tk.END)
        self.ui.after(100, self.thread_log)

    def check_files(self):
        check1 = check2 = check3 = False
        if os.path.isfile("반 정보.xlsx"):
            self.make_class_info_file_button["state"] = tk.DISABLED
            check1 = True
        else:
            self.make_class_info_file_button["state"] = tk.NORMAL
            self.make_student_info_file_button["state"] = tk.DISABLED
            self.make_data_file_button["state"] = tk.DISABLED
        if os.path.isfile("학생 정보.xlsx"):
            self.make_student_info_file_button["state"] = tk.DISABLED
            check2 = True
        else: 
            self.make_student_info_file_button["state"] = tk.NORMAL
        if os.path.isfile(f"./data/{config['dataFileName']}.xlsx"):
            self.make_data_file_button["state"] = tk.DISABLED
            check3 = True
        else:
            self.make_data_file_button["state"] = tk.NORMAL
        
        if check1 and check2 and check3:
            self.update_class_button["state"] = tk.NORMAL
            self.make_data_form_button["state"] = tk.NORMAL
            self.save_data_button["state"] = tk.NORMAL
            self.send_message_button["state"] = tk.NORMAL
            self.apply_color_button["state"] = tk.NORMAL
            self.add_student_button["state"] = tk.NORMAL
            self.delete_student_button["state"] = tk.NORMAL
            self.move_student_button["state"] = tk.NORMAL
        else:
            self.update_class_button["state"] = tk.DISABLED
            self.make_data_form_button["state"] = tk.DISABLED
            self.save_data_button["state"] = tk.DISABLED
            self.send_message_button["state"] = tk.DISABLED
            self.apply_color_button["state"] = tk.DISABLED
            self.add_student_button["state"] = tk.DISABLED
            self.delete_student_button["state"] = tk.DISABLED
            self.move_student_button["state"] = tk.DISABLED
        
        self.ui.after(100, self.check_files)

    def check_thread_end(self):
        if self.thread_end_flag:
            self.thread_end_flag = False
            self.ui.wm_attributes("-topmost", 1)
            self.ui.wm_attributes("-topmost", 0)
        self.ui.after(100, self.check_thread_end)

    # dialog
    def holiday_dialog(self) -> dict:
        def quitEvent():
            for i in range(7):
                if var_list[i].get():
                    makeup_test_date[weekday[i]] += timedelta(days=7)
            self.ui.wm_attributes("-disabled", False)
            popup.quit()
            popup.destroy()

        self.ui.wm_attributes("-disabled", True)
        popup = tk.Toplevel()
        width = 200
        height = 300
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("휴일 선택")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quitEvent)

        today = DATE.today()
        weekday = ("월", "화", "수", "목", "금", "토", "일")
        makeup_test_date = {weekday[i] : today + relativedelta(weekday=i) for i in range(7)}
        for key, value in makeup_test_date.items():
            if value == today: makeup_test_date[key] += timedelta(days=7)

        mon = tk.BooleanVar()
        tue = tk.BooleanVar()
        wed = tk.BooleanVar()
        thu = tk.BooleanVar()
        fri = tk.BooleanVar()
        sat = tk.BooleanVar()
        sun = tk.BooleanVar()
        var_list = [mon, tue, wed, thu, fri, sat, sun]
        tk.Label(popup, text="\n다음 중 휴일을 선택해주세요\n").pack()
        sort = today.weekday()+1
        for i in range(7):
            tk.Checkbutton(popup, text=f"{str(makeup_test_date[weekday[(sort+i)%7]])} {weekday[(sort+i)%7]}", variable=var_list[(sort+i)%7]).pack()
        tk.Label(popup, text="\n").pack()
        tk.Button(popup, text="확인", width=10 , command=quitEvent).pack()
        
        popup.mainloop()    
        
        return makeup_test_date

    def select_student_name_dialog(self) -> str:
        def quitEvent():
            self.ui.wm_attributes("-disabled", False)
            popup.quit()
            popup.destroy()

        self.ui.wm_attributes("-disabled", True)
        popup = tk.Toplevel()
        width = 250
        height = 140
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("퇴원 관리")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quitEvent)

        # 반 정보 확인
        class_wb = xl.load_workbook("./반 정보.xlsx")
        try:
            class_ws = class_wb["반 정보"]
        except:
            self.q.put(r"[오류] '반 정보.xlsx'의 시트명을")
            self.q.put(r"'반 정보'로 변경해 주세요.")
            return
        
        data_file_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx")
        data_file_ws = data_file_wb.worksheets[0]
        for i in range(1, data_file_ws.max_column):
            if data_file_ws.cell(1, i).value == "이름":
                STUDENT_NAME_COLUMN = i
                break
        for i in range(1, data_file_ws.max_column):
            if data_file_ws.cell(1, i).value == "반":
                CLASS_NAME_COLUMN = i
                break

        class_dict = {}
        for i in range(2, class_ws.max_row + 1):
            class_name = class_ws.cell(i, ClassInfo.CLASS_NAME_COLUMN).value
            student_list = [data_file_ws.cell(j, STUDENT_NAME_COLUMN).value for j in range(2, data_file_ws.max_row)\
                            if data_file_ws.cell(j, CLASS_NAME_COLUMN).value == class_name and\
                                data_file_ws.cell(j, STUDENT_NAME_COLUMN).value != "날짜" and\
                                    data_file_ws.cell(j, STUDENT_NAME_COLUMN).value != "시험명" and\
                                        data_file_ws.cell(j, STUDENT_NAME_COLUMN).value != "시험 평균" and\
                                            not data_file_ws.cell(j, STUDENT_NAME_COLUMN).font.strike]
            class_dict[class_name] = student_list

            
        tk.Label(popup).pack()
        def class_call_back(event):
            class_name = event.widget.get()
            student_combo.set("학생 선택")
            student_combo["values"] = class_dict[class_name]
        class_combo = ttk.Combobox(popup, values=list(class_dict.keys()), state="readonly")
        class_combo.set("반 선택")
        class_combo.bind("<<ComboboxSelected>>", class_call_back)
        class_combo.pack()

        tk.Label(popup).pack()
        selected_student = tk.StringVar()
        student_combo = ttk.Combobox(popup, values=None, state="readonly", textvariable=selected_student)
        student_combo.set("학생 선택")
        student_combo.pack()

        tk.Label(popup).pack()
        tk.Button(popup, text="삭제", width=10 , command=quitEvent).pack()
        
        popup.mainloop()
        
        student = selected_student.get()
        if student == "학생 선택":
            return None
        else:
            return student

    def move_student_dialog(self):
        def quitEvent():
            self.ui.wm_attributes("-disabled", False)
            popup.quit()
            popup.destroy()

        self.ui.wm_attributes("-disabled", True)
        popup = tk.Toplevel()
        width = 250
        height = 160
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("학생 반 이동")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quitEvent)

        # 반 정보 확인
        class_wb = xl.load_workbook("./반 정보.xlsx")
        try:
            class_ws = class_wb["반 정보"]
        except:
            self.q.put(r"[오류] '반 정보.xlsx'의 시트명을")
            self.q.put(r"'반 정보'로 변경해 주세요.")
            return
        
        data_file_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx")
        data_file_ws = data_file_wb.worksheets[0]
        for i in range(1, data_file_ws.max_column):
            if data_file_ws.cell(1, i).value == "이름":
                STUDENT_NAME_COLUMN = i
                break
        for i in range(1, data_file_ws.max_column):
            if data_file_ws.cell(1, i).value == "반":
                CLASS_NAME_COLUMN = i
                break

        class_dict = {}
        for i in range(2, class_ws.max_row + 1):
            class_name = class_ws.cell(i, ClassInfo.CLASS_NAME_COLUMN).value
            student_list = [data_file_ws.cell(j, STUDENT_NAME_COLUMN).value for j in range(2, data_file_ws.max_row)\
                            if data_file_ws.cell(j, CLASS_NAME_COLUMN).value == class_name and\
                                data_file_ws.cell(j, STUDENT_NAME_COLUMN).value != "날짜" and\
                                    data_file_ws.cell(j, STUDENT_NAME_COLUMN).value != "시험명" and\
                                        data_file_ws.cell(j, STUDENT_NAME_COLUMN).value != "시험 평균" and\
                                            not data_file_ws.cell(j, STUDENT_NAME_COLUMN).font.strike]
            class_dict[class_name] = student_list

        tk.Label(popup).pack()
        def class_call_back(event):
            class_name = event.widget.get()
            student_combo.set("학생 선택")
            student_combo["values"] = class_dict[class_name]
        current_class_var = tk.StringVar()
        current_class_combo = ttk.Combobox(popup, values=list(class_dict.keys()), state="readonly", textvariable=current_class_var)
        current_class_combo.set("반 선택")
        current_class_combo.bind("<<ComboboxSelected>>", class_call_back)
        current_class_combo.pack()

        selected_student = tk.StringVar()
        student_combo = ttk.Combobox(popup, values=None, state="readonly", textvariable=selected_student)
        student_combo.set("학생 선택")
        student_combo.pack()

        tk.Label(popup).pack()
        target_class_var = tk.StringVar()
        current_class_combo = ttk.Combobox(popup, values=list(class_dict.keys()), state="readonly", textvariable=target_class_var)
        current_class_combo.set("이동할 반 선택")
        current_class_combo.pack()

        tk.Label(popup).pack()
        tk.Button(popup, text="반 이동", width=10 , command=quitEvent).pack()
        
        popup.mainloop()
        
        target_student_name = selected_student.get()
        target_class_name = target_class_var.get()
        current_class_name = current_class_var.get()
        if target_student_name == "학생 선택" or target_class_name == "이동할 반 선택":
            return None
        else:
            return target_student_name, target_class_name, current_class_name

    def add_student_dialog(self):
        def quitEvent():
            self.ui.wm_attributes("-disabled", False)
            popup.quit()
            popup.destroy()

        self.ui.wm_attributes("-disabled", True)
        popup = tk.Toplevel()
        width = 250
        height = 140
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("신규생 추가")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quitEvent)

        # 반 정보 확인
        class_wb = xl.load_workbook("./반 정보.xlsx")
        try:
            class_ws = class_wb["반 정보"]
        except:
            self.q.put(r"[오류] '반 정보.xlsx'의 시트명을")
            self.q.put(r"'반 정보'로 변경해 주세요.")
            return
        
        class_names = [class_ws.cell(i, ClassInfo.CLASS_NAME_COLUMN).value for i in range(2, class_ws.max_row + 1)]

        tk.Label(popup).pack()
        target_class_var = tk.StringVar()
        class_combo = ttk.Combobox(popup, values=class_names, state="readonly", textvariable=target_class_var, width=25)
        class_combo.set("학생을 추가할 반 선택")
        class_combo.pack()

        tk.Label(popup).pack()
        target_student_var = tk.StringVar()
        tk.Entry(popup, textvariable=target_student_var, width=28).pack()

        tk.Label(popup).pack()
        tk.Button(popup, text="신규생 추가", width=10 , command=quitEvent).pack()
        
        popup.mainloop()
        
        target_class_name = target_class_var.get()
        target_student_name = target_student_var.get()
        if target_class_name == "학생을 추가할 반 선택":
            return None
        else:
            return target_student_name, target_class_name

    # threads
    def make_class_info_file_thread(self):
        thread = threading.Thread(target=lambda: make_class_info_file(self))
        thread.daemon = True
        thread.start()

    def make_student_info_file_thread(self):
        thread = threading.Thread(target=lambda: make_student_info_file(self))
        thread.daemon = True
        thread.start()

    def make_data_file_thread(self):
        thread = threading.Thread(target=lambda: make_data_file(self))
        thread.daemon = True
        thread.start()

    def update_class_thread(self):
        if self.update_class_button["text"] == "반 업데이트":
            if os.path.isfile("./temp.xlsx"):
                os.remove("./temp.xlsx")
            self.update_class_button["state"] = tk.DISABLED
            global ret
            ret = check_update_class(self)
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True
            wb = excel.Workbooks.Open(f"{os.getcwd()}\\temp.xlsx")
            self.update_class_button["text"] = "반 정보 수정 후 반 업데이트 계속하기"
            if not tkinter.messagebox.askokcancel("반 정보 변경 확인", "반 정보 파일의 빈칸을 채운 뒤 Excel을 종료하고\n버튼을 눌러주세요.\n삭제할 반은 행을 삭제해 주세요.\n취소 선택 시 반 업데이트가 중단됩니다."):
                self.update_class_button["text"] = "반 업데이트"
                wb.Close()
                if os.path.isfile("./temp.xlsx"):
                    os.remove("./temp.xlsx")
                    self.q.put(r"반 업데이트를 중단합니다.")
        else:
            if os.path.isfile("./~$temp.xlsx"):
                self.q.put(r"임시 파일을 닫은 뒤 다시 시도해 주세요.")
                return
            if os.path.isfile(f"./data/~${config['dataFileName']}.xlsx"):
                self.q.put(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
                return
            if not tkinter.messagebox.askokcancel("반 정보 변경 확인", "반 업데이트를 계속하시겠습니까?"):
                if os.path.isfile("./temp.xlsx"):
                    os.remove("./temp.xlsx")
                    self.q.put(r"반 업데이트를 중단합니다.")
            thread = threading.Thread(target=lambda: update_class(self, ret[0], ret[1]))
            thread.daemon = True
            thread.start()
            del ret
            self.update_class_button["text"] = "반 업데이트"
        self.update_class_button["state"] = tk.NORMAL

    def make_data_form_thread(self):
        self.make_data_form_button["state"] = tk.DISABLED
        thread = threading.Thread(target=lambda: make_data_form(self))
        thread.daemon = True
        thread.start()
        self.make_data_form_button['state'] = tk.NORMAL

    def save_data_thread(self):
        if os.path.isfile("./data/~$재시험 명단.xlsx"):
            self.q.put(r"재시험 명단 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        if os.path.isfile(f"./data/~${config['dataFileName']}.xlsx"):
            self.q.put(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        self.save_data_button["state"] = tk.DISABLED
        makeup_test_date = self.holiday_dialog()
        filepath = filedialog.askopenfilename(initialdir="./", title="데일리테스트 기록 양식 선택", filetypes=(("Excel files", "*.xlsx"),("all files", "*.*")))
        if filepath == "": return
        thread = threading.Thread(target=lambda: save_data(self, filepath, makeup_test_date))
        thread.daemon = True
        thread.start()
        self.save_data_button["state"] = tk.NORMAL

    def send_message_thread(self):
        self.send_message_button["state"] = tk.DISABLED
        makeup_test_date = self.holiday_dialog()
        filepath = filedialog.askopenfilename(initialdir="./", title="데일리테스트 기록 양식 선택", filetypes=(("Excel files", "*.xlsx"),("all files", "*.*")))
        if filepath == "": return
        thread = threading.Thread(target=lambda: send_message(self, filepath, makeup_test_date))
        thread.daemon = True
        thread.start()
        self.send_message_button["state"] = tk.NORMAL

    def add_student_thread(self):
        if os.path.isfile(f"./data/~${config['dataFileName']}.xlsx"):
            self.q.put(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        if os.path.isfile("./data/~$학생 정보.xlsx"):
            self.q.put(r"학생 정보 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        self.add_student_button["state"] = tk.DISABLED
        tmp = self.add_student_dialog()
        if tmp is not None:
            # 학생 추가 확인
            if not tkinter.messagebox.askyesno("학생 추가 확인", f"{tmp[0]} 학생을 {tmp[1]} 반에 추가하시겠습니까?"):
                self.add_student_button["state"] = tk.NORMAL
                return
            thread = threading.Thread(target=lambda: add_student(self, tmp[0], tmp[1]))
            thread.daemon = True
            thread.start()
        self.add_student_button["state"] = tk.NORMAL

    def delete_student_thread(self):
        if os.path.isfile(f"./data/~${config['dataFileName']}.xlsx"):
            self.q.put(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        if os.path.isfile("./data/~$학생 정보.xlsx"):
            self.q.put(r"학생 정보 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        self.delete_student_button["state"] = tk.DISABLED
        student = self.select_student_name_dialog()
        if student is not None:
            # 퇴원 처리 확인
            if not tkinter.messagebox.askyesno("퇴원 확인", f"{student} 학생을 퇴원 처리하시겠습니까?"):
                self.delete_student_button["state"] = tk.NORMAL
                return
            thread = threading.Thread(target=lambda: delete_student(self, student))
            thread.daemon = True
            thread.start()
        self.delete_student_button["state"] = tk.NORMAL

    def move_student_thread(self):
        if os.path.isfile(f"./data/~${config['dataFileName']}.xlsx"):
            self.q.put(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        if os.path.isfile("./data/~$학생 정보.xlsx"):
            self.q.put(r"학생 정보 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        self.move_student_button["state"] = tk.DISABLED
        tmp = self.move_student_dialog()
        if tmp is not None:
            # 학생 반 이동 확인
            if tmp[1] == tmp[2]:
                self.q.put(r"학생의 현재 반과 이동할 반이 같아 취소되었습니다.")
                self.move_student_button["state"] = tk.NORMAL
                return
            if not tkinter.messagebox.askyesno("학생 반 이동 확인", f"{tmp[2]} 반의 {tmp[0]} 학생을\n{tmp[1]} 반으로 이동시키겠습니까?"):
                self.move_student_button["state"] = tk.NORMAL
                return
            thread = threading.Thread(target=lambda: move_student(self, tmp[0], tmp[1], tmp[2]))
            thread.daemon = True
            thread.start()
        self.move_student_button["state"] = tk.NORMAL

# tasks
def make_class_info_file(gui:GUI):
    gui.q.put("반 정보 입력 파일 생성 중...")

    ini_wb = xl.Workbook()
    ini_ws = ini_wb.active
    ini_ws.title = "반 정보"
    ini_ws[gcl(ClassInfo.CLASS_NAME_COLUMN)+"1"] = "반명"
    ini_ws[gcl(ClassInfo.TEACHER_COLUMN)+"1"] = "선생님명"
    ini_ws[gcl(ClassInfo.DATE_COLUMN)+"1"] = "요일"
    ini_ws[gcl(ClassInfo.TEST_TIME_COLUMN)+"1"] = "시간"

    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    driver = webdriver.Chrome(service = service, options = options)

    # 아이소식 접속
    driver.get(config["url"])
    table_names = driver.find_elements(By.CLASS_NAME, "style1")

    # 반 루프
    for tableName in table_names:
        WRITE_LOCATION = ini_ws.max_row + 1
        ini_ws.cell(WRITE_LOCATION, 1).value = tableName.text.rstrip()

    # 정렬 및 테두리
    for j in range(1, ini_ws.max_row + 1):
        for k in range(1, ini_ws.max_column + 1):
            ini_ws.cell(j, k).alignment = Alignment(horizontal="center", vertical="center")
            ini_ws.cell(j, k).border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    ini_wb.save("./반 정보.xlsx")
    gui.q.put("반 정보 입력 파일 생성을 완료했습니다.")
    gui.q.put("반 정보를 입력해 주세요.")
    gui.thread_end_flag = True

def make_student_info_file(gui:GUI):
    if not os.path.isfile("./학생 정보.xlsx"):
        gui.q.put("학생 정보 파일 생성 중...")

        ini_wb = xl.Workbook()
        ini_ws = ini_wb.active
        ini_ws.title = "학생 정보"
        ini_ws[gcl(StudentInfo.STUDENT_NAME_COLUMN)+"1"] = "이름"
        ini_ws[gcl(StudentInfo.CLASS_NAME_COLUMN)+"1"] = "반명"
        ini_ws[gcl(StudentInfo.TEACHER_COLUMN)+"1"] = "담당"
        ini_ws[gcl(StudentInfo.MAKEUP_TEST_WEEK_DATE_COLUMN)+"1"] = "재시험 응시 요일"
        ini_ws[gcl(StudentInfo.MAKEUP_TEST_TIME_COLUMN)+"1"] = "재시험 응시 시간"
        ini_ws[gcl(StudentInfo.NEW_STUDENT_CHECK_COLUMN)+"1"] = "기수 신규생"
        ini_ws["Z1"] = "N"
        ini_ws.auto_filter.ref = "A:"+gcl(StudentInfo.MAX)
        ini_ws.column_dimensions.group("Z", hidden=True)
        ini_wb.save("./학생 정보.xlsx")
    else:
        gui.q.put("학생 정보 파일 업데이트 중...")
    
    student_wb = xl.load_workbook("./학생 정보.xlsx")
    try:
        student_ws = student_wb["학생 정보"]
    except:
        gui.q.put(r"[오류] '학생 정보.xlsx'의 시트명을")
        gui.q.put(r"'학생 정보'로 변경해 주세요.")
        return
    
    # 반 정보 확인
    class_wb = xl.load_workbook("./반 정보.xlsx")
    try:
        class_ws = class_wb["반 정보"]
    except:
        gui.q.put(r"[오류] '반 정보.xlsx'의 시트명을")
        gui.q.put(r"'반 정보'로 변경해 주세요.")
        return

    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    driver = webdriver.Chrome(service = service, options = options)

    dv = DataValidation(type="list", formula1="=Z1",  allow_blank=True, errorStyle="stop", showErrorMessage=True)
    student_ws.add_data_validation(dv)

    # 아이소식 접속
    driver.get(config["url"])
    table_names = driver.find_elements(By.CLASS_NAME, "style1")

    # 반 루프
    for i in range(3, len(table_names)):
        trs = driver.find_element(By.ID, "table_" + str(i)).find_elements(By.CLASS_NAME, "style12")
        WRITE_LOCATION = student_ws.max_row + 1
        teacher = ""

        class_name = table_names[i].text.rstrip()
        for j in range(2, class_ws.max_row + 1):
            if class_ws.cell(j, 1).value == class_name:
                teacher = class_ws.cell(j, 2).value
                is_class_exist = True
        if not is_class_exist:
            continue

        # 학생 루프
        for tr in trs:
            WRITE_LOCATION = ini_ws.max_row + 1
            ini_ws.cell(WRITE_LOCATION, StudentInfo.STUDENT_NAME_COLUMN).value = tr.find_element(By.CLASS_NAME, "style9").text
            ini_ws.cell(WRITE_LOCATION, StudentInfo.CLASS_NAME_COLUMN).value = class_name
            ini_ws.cell(WRITE_LOCATION, StudentInfo.TEACHER_COLUMN).value = teacher
            dv.add(ini_ws.cell(WRITE_LOCATION, StudentInfo.NEW_STUDENT_CHECK_COLUMN))

    # 정렬 및 테두리
    for j in range(1, ini_ws.max_row + 1):
        for k in range(1, StudentInfo.MAX + 1):
            ini_ws.cell(j, k).alignment = Alignment(horizontal="center", vertical="center")
            ini_ws.cell(j, k).border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    
    ini_wb.save("./학생 정보.xlsx")
    gui.q.put("학생 정보 파일을 생성했습니다.")
    gui.thread_end_flag = True

def make_data_file(gui:GUI):
    gui.q.put("데이터파일 생성 중...")

    ini_wb = xl.Workbook()
    ini_ws = ini_wb.active
    ini_ws.title = "데일리테스트"
    ini_ws[gcl(DataFile.TEST_TIME_COLUMN)+"1"] = "시간"
    ini_ws[gcl(DataFile.DATE_COLUMN)+"1"] = "요일"
    ini_ws[gcl(DataFile.CLASS_NAME_COLUMN)+"1"] = "반"
    ini_ws[gcl(DataFile.TEACHER_COLUMN)+"1"] = "담당"
    ini_ws[gcl(DataFile.STUDENT_NAME_COLUMN)+"1"] = "이름"
    ini_ws[gcl(DataFile.AVERAGE_SCORE_COLUMN)+"1"] = "학생 평균"
    ini_ws.freeze_panes = gcl(DataFile.DATA_COLUMN) + "2"
    ini_ws.auto_filter.ref = "A:" + gcl(DataFile.MAX)

    # 반 정보 확인
    class_wb = xl.load_workbook("./반 정보.xlsx")
    try:
        class_ws = class_wb["반 정보"]
    except:
        gui.q.put(r"[오류] '반 정보.xlsx'의 시트명을")
        gui.q.put(r"'반 정보'로 변경해 주세요.")
        return

    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    driver = webdriver.Chrome(service = service, options = options)

    # 아이소식 접속
    driver.get(config["url"])
    table_names = driver.find_elements(By.CLASS_NAME, "style1")

    # 반 루프
    for i in range(3, len(table_names)):
        trs = driver.find_element(By.ID, "table_" + str(i)).find_elements(By.CLASS_NAME, "style12")
        WRITE_LOCATION = ini_ws.max_row + 1

        class_name = table_names[i].text.rstrip()
        time = ""
        date = ""
        teacher = ""
        is_class_exist = False
        for j in range(2, class_ws.max_row + 1):
            if class_ws.cell(j, ClassInfo.CLASS_NAME_COLUMN).value == class_name:
                teacher = class_ws.cell(j, ClassInfo.TEACHER_COLUMN).value
                date = class_ws.cell(j, ClassInfo.DATE_COLUMN).value
                time = class_ws.cell(j, ClassInfo.TEST_TIME_COLUMN).value
                is_class_exist = True
        if not is_class_exist or len(trs) == 0:
            continue
        
        # 시험명
        ini_ws.cell(WRITE_LOCATION, DataFile.TEST_TIME_COLUMN).value = time
        ini_ws.cell(WRITE_LOCATION, DataFile.DATE_COLUMN).value = date
        ini_ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value = class_name
        ini_ws.cell(WRITE_LOCATION, DataFile.TEACHER_COLUMN).value = teacher
        ini_ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value = "날짜"
        
        WRITE_LOCATION = ini_ws.max_row + 1
        ini_ws.cell(WRITE_LOCATION, DataFile.TEST_TIME_COLUMN).value = time
        ini_ws.cell(WRITE_LOCATION, DataFile.DATE_COLUMN).value = date
        ini_ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value = class_name
        ini_ws.cell(WRITE_LOCATION, DataFile.TEACHER_COLUMN).value = teacher
        ini_ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value = "시험명"
        class_start = WRITE_LOCATION + 1

        # 학생 루프
        for tr in trs:
            WRITE_LOCATION = ini_ws.max_row + 1
            ini_ws.cell(WRITE_LOCATION, DataFile.TEST_TIME_COLUMN).value = time
            ini_ws.cell(WRITE_LOCATION, DataFile.DATE_COLUMN).value = date
            ini_ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value = class_name
            ini_ws.cell(WRITE_LOCATION, DataFile.TEACHER_COLUMN).value = teacher
            ini_ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value = tr.find_element(By.CLASS_NAME, "style9").text
            ini_ws.cell(WRITE_LOCATION, DataFile.AVERAGE_SCORE_COLUMN).value = f"=ROUND(AVERAGE(G{str(WRITE_LOCATION)}:XFD{str(WRITE_LOCATION)}), 0)"
            ini_ws.cell(WRITE_LOCATION, DataFile.AVERAGE_SCORE_COLUMN).font = Font(bold=True)
        
        # 시험별 평균
        WRITE_LOCATION = ini_ws.max_row + 1
        class_end = WRITE_LOCATION - 1
        ini_ws.cell(WRITE_LOCATION, DataFile.TEST_TIME_COLUMN).value = time
        ini_ws.cell(WRITE_LOCATION, DataFile.DATE_COLUMN).value = date
        ini_ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value = class_name
        ini_ws.cell(WRITE_LOCATION, DataFile.TEACHER_COLUMN).value = teacher
        ini_ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value = "시험 평균"
        ini_ws[f"F{str(WRITE_LOCATION)}"] = ArrayFormula(f"F{str(WRITE_LOCATION)}", f"=ROUND(AVERAGE(IFERROR(F{str(class_start)}:F{str(class_end)}, \"\")), 0)")
        ini_ws.cell(WRITE_LOCATION, DataFile.AVERAGE_SCORE_COLUMN).font = Font(bold=True)

        for j in range(1, DataFile.DATA_COLUMN):
            ini_ws.cell(WRITE_LOCATION, j).border = Border(bottom = Side(border_style="medium", color="000000"))

    # 정렬
    for i in range(1, ini_ws.max_row + 1):
        for j in range(1, ini_ws.max_column + 1):
            ini_ws.cell(i, j).alignment = Alignment(horizontal="center", vertical="center")
    
    # 모의고사 sheet 생성
    copy_ws = ini_wb.copy_worksheet(ini_wb["데일리테스트"])
    copy_ws.title = "모의고사"
    copy_ws.freeze_panes = gcl(DataFile.DATA_COLUMN) + "2"
    copy_ws.auto_filter.ref = "A:" + gcl(DataFile.STUDENT_NAME_COLUMN)

    ini_wb.save(f"./data/{config['dataFileName']}.xlsx")
    gui.q.put("데이터 파일을 생성했습니다.")
    gui.thread_end_flag = True

def check_update_class(gui:GUI):
    # 반 정보 확인
    class_wb = xl.load_workbook("./반 정보.xlsx")
    try:
        class_ws = class_wb["반 정보"]
    except:
        gui.q.put(r"[오류] '반 정보.xlsx'의 시트명을")
        gui.q.put(r"'반 정보'로 변경해 주세요.")
        return

    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    driver = webdriver.Chrome(service = service, options = options)

    gui.q.put("아이소식으로부터 반 정보를 업데이트 하는 중...")
    # 아이소식 접속
    driver.get(config["url"])
    table_names = driver.find_elements(By.CLASS_NAME, "style1")
    current_class = [class_ws.cell(i, ClassInfo.CLASS_NAME_COLUMN).value for i in range(2, class_ws.max_row+1) if class_ws.cell(i, ClassInfo.CLASS_NAME_COLUMN).value is not None]

    unregistered_class = {table_names[i].text.rstrip() : i for i in range(3, len(table_names)) if not table_names[i].text.rstrip() in current_class}

    for i in range(2, class_ws.max_row+2):
        if class_ws.cell(i, ClassInfo.CLASS_NAME_COLUMN).value is None:
            WRITE_LOCATION = i
            break
    
    for new_class_name in list(unregistered_class.keys()):
        class_ws.cell(WRITE_LOCATION, ClassInfo.CLASS_NAME_COLUMN).value = new_class_name
        WRITE_LOCATION += 1
    
    # 정렬 및 테두리
    for j in range(1, class_ws.max_row + 1):
        for k in range(1, class_ws.max_column + 1):
            class_ws.cell(j, k).alignment = Alignment(horizontal="center", vertical="center")
            class_ws.cell(j, k).border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    
    class_wb.save("./temp.xlsx")
    return current_class, unregistered_class

def update_class(gui:GUI, current_class:list, unregistered_class:dict):
    gui.q.put("백업 파일 생성중...")
    data_file_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx")
    data_file_wb.save(f"./data/backup/{config['dataFileName']}({datetime.today().strftime('%Y%m%d')}).xlsx")
    class_wb = xl.load_workbook("./반 정보.xlsx")
    class_wb.save(f"./data/backup/반 정보({datetime.today().strftime('%Y%m%d')}).xlsx")

    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    driver = webdriver.Chrome(service = service, options = options)
    driver.get(config["url"])

    gui.q.put("수정된 반 정보를 바탕으로 업데이트 중...")
    class_wb_temp = xl.load_workbook("./temp.xlsx")

    class_temp_ws = class_wb_temp["반 정보"]
    update_class = [class_temp_ws.cell(i, ClassInfo.CLASS_NAME_COLUMN).value for i in range(2, class_temp_ws.max_row+1) if class_temp_ws.cell(i, ClassInfo.CLASS_NAME_COLUMN).value is not None]
    delete_class = [c for c in current_class if not c in update_class]
    check_list = [c for c in list(unregistered_class.keys()) if c in update_class]

    data_file_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx")
    data_file_ws = data_file_wb["데일리테스트"]

    if len(check_list) == 0 and len(delete_class) == 0:
        data_file_wb.save(f"./data/{config['dataFileName']}.xlsx")
        class_wb_temp.save("./반 정보.xlsx")
        os.remove("./temp.xlsx")
        gui.q.put("업데이트 된 항목이 없습니다.")
        return
    
    # 데이터 파일에서 이전 데이터 이동 및 삭제
    if len(delete_class) != 0:
        gui.q.put("이전 데이터 제거 중...")
        if not os.path.isfile("./data/지난 데이터.xlsx"):
            ini_wb = xl.Workbook()
            ini_ws = ini_wb.active
            ini_ws.title = "데일리테스트"
            ini_ws[gcl(DataFile.TEST_TIME_COLUMN)+"1"] = "시간"
            ini_ws[gcl(DataFile.DATE_COLUMN)+"1"] = "요일"
            ini_ws[gcl(DataFile.CLASS_NAME_COLUMN)+"1"] = "반"
            ini_ws[gcl(DataFile.TEACHER_COLUMN)+"1"] = "담당"
            ini_ws[gcl(DataFile.STUDENT_NAME_COLUMN)+"1"] = "이름"
            ini_ws[gcl(DataFile.AVERAGE_SCORE_COLUMN)+"1"] = "학생 평균"
            ini_ws.freeze_panes = gcl(DataFile.DATA_COLUMN) + "2"
            ini_ws.auto_filter.ref = "A:" + gcl(DataFile.MAX)

            copy_ws = ini_wb.copy_worksheet(ini_wb["데일리테스트"])
            copy_ws.title = "모의고사"
            copy_ws.freeze_panes = gcl(DataFile.DATA_COLUMN) + "2"
            copy_ws.auto_filter.ref = "A:" + gcl(DataFile.STUDENT_NAME_COLUMN)

            ini_wb.save("./data/지난 데이터.xlsx")
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        try:
            wb = excel.Workbooks.Open(f"{os.getcwd()}\\data\\{config['dataFileName']}.xlsx")
            wb.Save()
            wb.Close()
        except:
            gui.q.put("데이터 파일 창을 끄고 다시 실행해 주세요.")
            return
        
        post_data_wb = xl.load_workbook("./data/지난 데이터.xlsx")
        data_file_temp_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx", data_only=True)
        for sheet_name in data_file_temp_wb.sheetnames:
            data_file_temp_ws = data_file_temp_wb[sheet_name]
            post_data_ws = post_data_wb[sheet_name]
            data_file_ws = data_file_wb[sheet_name]
            
            # 동적 열 탐색
            for i in range(1, data_file_temp_ws.max_column+1):
                temp = data_file_temp_ws.cell(1, i).value
                if temp == "시간":
                    TEST_TIME_COLUMN = i
                elif temp == "요일":
                    DATE_COLUMN = i
                elif temp == "반":
                    CLASS_NAME_COLUMN = i
                elif temp == "담당":
                    TEACHER_COLUMN = i
                elif temp == "이름":
                    STUDENT_NAME_COLUMN = i
                elif temp == "학생 평균":
                    AVERAGE_SCORE_COLUMN = i
            
            # 지난 데이터 행 삭제
            for row in range(2, data_file_ws.max_row+1):
                while data_file_ws.cell(row, CLASS_NAME_COLUMN).value in delete_class:
                    data_file_ws.delete_rows(row)

            # 지난 데이터 행 복사
            for row in range(2, data_file_temp_ws.max_row+1):
                if data_file_temp_ws.cell(row, CLASS_NAME_COLUMN).value in delete_class:
                    write_row = post_data_ws.max_row+1
                    copy_cell(post_data_ws.cell(write_row, DataFile.TEST_TIME_COLUMN), data_file_temp_ws.cell(row, TEST_TIME_COLUMN))
                    copy_cell(post_data_ws.cell(write_row, DataFile.DATE_COLUMN), data_file_temp_ws.cell(row, DATE_COLUMN))
                    copy_cell(post_data_ws.cell(write_row, DataFile.CLASS_NAME_COLUMN), data_file_temp_ws.cell(row, CLASS_NAME_COLUMN))
                    copy_cell(post_data_ws.cell(write_row, DataFile.TEACHER_COLUMN), data_file_temp_ws.cell(row, TEACHER_COLUMN))
                    copy_cell(post_data_ws.cell(write_row, DataFile.STUDENT_NAME_COLUMN), data_file_temp_ws.cell(row, STUDENT_NAME_COLUMN))
                    copy_cell(post_data_ws.cell(write_row, DataFile.AVERAGE_SCORE_COLUMN), data_file_temp_ws.cell(row, AVERAGE_SCORE_COLUMN))
                    write_column = DataFile.MAX+1
                    for col in range(AVERAGE_SCORE_COLUMN+1, data_file_temp_ws.max_column):
                        copy_cell(post_data_ws.cell(write_row, write_column), data_file_temp_ws.cell(row, col))
                        write_column += 1

            # 평균 범위 재지정
            rescoping_formula()

            # 필터 범위 재조정
            data_file_ws.auto_filter.ref = "A:" + gcl(AVERAGE_SCORE_COLUMN)
        
        post_data_wb.save("./data/지난 데이터.xlsx")
        data_file_wb.save(f"./data/{config['dataFileName']}.xlsx")

    # 데이터 파일에 새 반 추가
    if len(check_list) != 0:
        gui.q.put("신규 반 추가중...")
        data_file_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx")
        for sheet_name in data_file_wb.sheetnames:
            data_file_ws = data_file_wb[sheet_name]
            for i in range(1, data_file_ws.max_column+1):
                temp = data_file_ws.cell(1, i).value
                if temp == "시간":
                    TEST_TIME_COLUMN = i
                elif temp == "요일":
                    DATE_COLUMN = i
                elif temp == "반":
                    CLASS_NAME_COLUMN = i
                elif temp == "담당":
                    TEACHER_COLUMN = i
                elif temp == "이름":
                    STUDENT_NAME_COLUMN = i
                elif temp == "학생 평균":
                    AVERAGE_SCORE_COLUMN = i
            
            for i in range(2, data_file_ws.max_row+2):
                if data_file_ws.cell(i, CLASS_NAME_COLUMN).value is None:
                    WRITE_LOCATION = i-1
                    break
            
            for new_class, new_class_index in unregistered_class.items():
                if not new_class in update_class: continue
                trs = driver.find_element(By.ID, "table_" + str(new_class_index)).find_elements(By.CLASS_NAME, "style12")

                class_name = new_class
                time = ""
                date = ""
                teacher = ""
                is_class_exist = False
                for j in range(2, class_temp_ws.max_row + 1):
                    if class_temp_ws.cell(j, ClassInfo.CLASS_NAME_COLUMN).value == class_name:
                        teacher = class_temp_ws.cell(j, ClassInfo.TEACHER_COLUMN).value
                        date = class_temp_ws.cell(j, ClassInfo.DATE_COLUMN).value
                        time = class_temp_ws.cell(j, ClassInfo.TEST_TIME_COLUMN).value
                        is_class_exist = True
                if not is_class_exist or len(trs) == 0:
                    continue
                WRITE_LOCATION += 1
                # 시험명
                data_file_ws.cell(WRITE_LOCATION, TEST_TIME_COLUMN).value = time
                data_file_ws.cell(WRITE_LOCATION, DATE_COLUMN).value = date
                data_file_ws.cell(WRITE_LOCATION, CLASS_NAME_COLUMN).value = class_name
                data_file_ws.cell(WRITE_LOCATION, TEACHER_COLUMN).value = teacher
                data_file_ws.cell(WRITE_LOCATION, STUDENT_NAME_COLUMN).value = "날짜"
                
                WRITE_LOCATION += 1
                data_file_ws.cell(WRITE_LOCATION, TEST_TIME_COLUMN).value = time
                data_file_ws.cell(WRITE_LOCATION, DATE_COLUMN).value = date
                data_file_ws.cell(WRITE_LOCATION, CLASS_NAME_COLUMN).value = class_name
                data_file_ws.cell(WRITE_LOCATION, TEACHER_COLUMN).value = teacher
                data_file_ws.cell(WRITE_LOCATION, STUDENT_NAME_COLUMN).value = "시험명"
                start = WRITE_LOCATION + 1

                # 학생 루프
                for tr in trs:
                    WRITE_LOCATION += 1
                    data_file_ws.cell(WRITE_LOCATION, TEST_TIME_COLUMN).value = time
                    data_file_ws.cell(WRITE_LOCATION, DATE_COLUMN).value = date
                    data_file_ws.cell(WRITE_LOCATION, CLASS_NAME_COLUMN).value = class_name
                    data_file_ws.cell(WRITE_LOCATION, TEACHER_COLUMN).value = teacher
                    data_file_ws.cell(WRITE_LOCATION, STUDENT_NAME_COLUMN).value = tr.find_element(By.CLASS_NAME, "style9").text
                    data_file_ws.cell(WRITE_LOCATION, AVERAGE_SCORE_COLUMN).value = f"=ROUND(AVERAGE({gcl(AVERAGE_SCORE_COLUMN+1)}{str(WRITE_LOCATION)}:XFD{str(WRITE_LOCATION)}), 0)"
                    data_file_ws.cell(WRITE_LOCATION, AVERAGE_SCORE_COLUMN).font = Font(bold=True)
                
                # 시험별 평균
                WRITE_LOCATION += 1
                end = WRITE_LOCATION - 1
                data_file_ws.cell(WRITE_LOCATION, TEST_TIME_COLUMN).value = time
                data_file_ws.cell(WRITE_LOCATION, DATE_COLUMN).value = date
                data_file_ws.cell(WRITE_LOCATION, CLASS_NAME_COLUMN).value = class_name
                data_file_ws.cell(WRITE_LOCATION, TEACHER_COLUMN).value = teacher
                data_file_ws.cell(WRITE_LOCATION, STUDENT_NAME_COLUMN).value = "시험 평균"
                data_file_ws.cell(WRITE_LOCATION, AVERAGE_SCORE_COLUMN).value = f"=ROUND(AVERAGE({gcl(AVERAGE_SCORE_COLUMN)}{str(start)}:{gcl(AVERAGE_SCORE_COLUMN)}{str(end)}), 0)"
                data_file_ws.cell(WRITE_LOCATION, AVERAGE_SCORE_COLUMN).font = Font(bold=True)

                for j in range(1, AVERAGE_SCORE_COLUMN+1):
                    data_file_ws.cell(WRITE_LOCATION, j).border = Border(bottom = Side(border_style="medium", color="000000"))

            # 정렬
            for i in range(1, data_file_ws.max_row + 1):
                for j in range(1, data_file_ws.max_column + 1):
                    data_file_ws.cell(i, j).alignment = Alignment(horizontal="center", vertical="center")

            # 필터 범위 재지정
            data_file_ws.auto_filter.ref = "A:" + gcl(AVERAGE_SCORE_COLUMN)
    
    # 변경 사항 저장
    data_file_wb.save(f"./data/{config['dataFileName']}.xlsx")
    class_wb_temp.save("./반 정보.xlsx")
    os.remove("./temp.xlsx")

    gui.q.put("반 업데이트를 완료하였습니다.")
    gui.thread_end_flag = True
    pythoncom.CoUninitialize()

def make_data_form(gui:GUI):
    gui.q.put("데일리테스트 기록 양식 생성 중...")

    ini_wb = xl.Workbook()
    ini_ws = ini_wb.active
    ini_ws.title = "데일리테스트 기록 양식"
    ini_ws[gcl(DataForm.DATE_COLUMN)+"1"] = "요일"
    ini_ws[gcl(DataForm.TEST_TIME_COLUMN)+"1"] = "시간"
    ini_ws[gcl(DataForm.CLASS_NAME_COLUMN)+"1"] = "반"
    ini_ws[gcl(DataForm.STUDENT_NAME_COLUMN)+"1"] = "이름"
    ini_ws[gcl(DataForm.TEACHER_COLUMN)+"1"] = "담당T"
    ini_ws[gcl(DataForm.DAILYTEST_TEST_NAME_COLUMN)+"1"] = "시험명"
    ini_ws[gcl(DataForm.DAILYTEST_SCORE_COLUMN)+"1"] = "점수"
    ini_ws[gcl(DataForm.DAILYTEST_AVERAGE_COLUMN)+"1"] = "평균"
    ini_ws[gcl(DataForm.MOCKTEST_TEST_NAME_COLUMN)+"1"] = "시험대비 모의고사명"
    ini_ws[gcl(DataForm.MOCKTEST_SCORE_COLUMN)+"1"] = "모의고사 점수"
    ini_ws[gcl(DataForm.MOCKTEST_AVERAGE_COLUMN)+"1"] = "모의고사 평균"
    ini_ws[gcl(DataForm.MAKEUP_TEST_CHECK_COLUMN)+"1"] = "재시문자 X"
    ini_ws["Y1"] = "X"
    ini_ws["Z1"] = "x"
    ini_ws.column_dimensions.group("Y", "Z", hidden=True)
    ini_ws.auto_filter.ref = "A:"+gcl(DataForm.TEST_TIME_COLUMN)
    
    class_wb = xl.load_workbook("./반 정보.xlsx")
    try:
        class_ws = class_wb["반 정보"]
    except:
        gui.q.put(r"[오류] '반 정보.xlsx'의 시트명을")
        gui.q.put(r"'반 정보'로 변경해 주세요.")
        return

    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    driver = webdriver.Chrome(service = service, options = options)

    # 아이소식 접속
    gui.q.put("아이소식 접속 중")
    driver.get(config["url"])
    table_names = driver.find_elements(By.CLASS_NAME, "style1")

    dv = DataValidation(type="list", formula1="=Y1:Z1", showDropDown=True, allow_blank=True, showErrorMessage=True)
    ini_ws.add_data_validation(dv)

    #반 루프
    for i in range(3, len(table_names)):
        trs = driver.find_element(By.ID, "table_" + str(i)).find_elements(By.CLASS_NAME, "style12")
        WRITE_LOCATION = start = ini_ws.max_row + 1

        class_name = table_names[i].text.rstrip()
        teacher = ""
        date = ""
        time = ""
        is_class_exist = False

        for row in range(2, class_ws.max_row + 1):
            if class_ws.cell(row, ClassInfo.CLASS_NAME_COLUMN).value == class_name:
                teacher = class_ws.cell(row, ClassInfo.TEACHER_COLUMN).value
                date = class_ws.cell(row, ClassInfo.DATE_COLUMN).value
                time = class_ws.cell(row, ClassInfo.TEST_TIME_COLUMN).value
                is_class_exist = True
        if not is_class_exist or len(trs) == 0:
            continue
        ini_ws.cell(WRITE_LOCATION, DataForm.CLASS_NAME_COLUMN).value = class_name
        ini_ws.cell(WRITE_LOCATION, DataForm.TEACHER_COLUMN).value = teacher

        #학생 루프
        for tr in trs:
            ini_ws.cell(WRITE_LOCATION, DataForm.DATE_COLUMN).value = date
            ini_ws.cell(WRITE_LOCATION, DataForm.TEST_TIME_COLUMN).value = time
            ini_ws.cell(WRITE_LOCATION, DataForm.STUDENT_NAME_COLUMN).value = tr.find_element(By.CLASS_NAME, "style9").text
            dv.add(ini_ws.cell(WRITE_LOCATION,DataForm.MAKEUP_TEST_CHECK_COLUMN))
            WRITE_LOCATION = ini_ws.max_row + 1
        
        end = WRITE_LOCATION - 1

        # 시험 평균
        ini_ws.cell(start, DataForm.DAILYTEST_AVERAGE_COLUMN).value = f"=ROUND(AVERAGE({gcl(DataForm.DAILYTEST_SCORE_COLUMN)+str(start)}:{gcl(DataForm.DAILYTEST_SCORE_COLUMN)+str(end)}), 0)"
        # 모의고사 평균
        ini_ws.cell(start, DataForm.MOCKTEST_AVERAGE_COLUMN).value = f"=ROUND(AVERAGE({gcl(DataForm.MOCKTEST_SCORE_COLUMN)+str(start)}:{gcl(DataForm.MOCKTEST_SCORE_COLUMN)+str(end)}), 0)"
        
        # 정렬 및 테두리
        for j in range(1, ini_ws.max_row + 1):
            for k in range(1, DataForm.MAX+1):
                ini_ws.cell(j, k).alignment = Alignment(horizontal="center", vertical="center")
                ini_ws.cell(j, k).border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        
        # 셀 병합
        if start < end:
            ini_ws.merge_cells(f"{gcl(DataForm.CLASS_NAME_COLUMN)+str(start)}:{gcl(DataForm.CLASS_NAME_COLUMN)+str(end)}")
            ini_ws.merge_cells(f"{gcl(DataForm.TEACHER_COLUMN)+str(start)}:{gcl(DataForm.TEACHER_COLUMN)+str(end)}")
            ini_ws.merge_cells(f"{gcl(DataForm.DAILYTEST_TEST_NAME_COLUMN)+str(start)}:{gcl(DataForm.DAILYTEST_TEST_NAME_COLUMN)+str(end)}")
            ini_ws.merge_cells(f"{gcl(DataForm.DAILYTEST_AVERAGE_COLUMN)+str(start)}:{gcl(DataForm.DAILYTEST_AVERAGE_COLUMN)+str(end)}")
            ini_ws.merge_cells(f"{gcl(DataForm.MOCKTEST_TEST_NAME_COLUMN)+str(start)}:{gcl(DataForm.MOCKTEST_TEST_NAME_COLUMN)+str(end)}")
            ini_ws.merge_cells(f"{gcl(DataForm.MOCKTEST_AVERAGE_COLUMN)+str(start)}:{gcl(DataForm.MOCKTEST_AVERAGE_COLUMN)+str(end)}")
        
    ini_ws.protection.sheet = True
    ini_ws.protection.selectLockedCells = True
    ini_ws.protection.autoFilter = False
    ini_ws.protection.formatColumns = False
    for row in range(2, ini_ws.max_row + 1):
        ini_ws.cell(row, DataForm.DAILYTEST_TEST_NAME_COLUMN).protection = Protection(locked=False)
        ini_ws.cell(row, DataForm.DAILYTEST_SCORE_COLUMN).protection = Protection(locked=False)
        ini_ws.cell(row, DataForm.MOCKTEST_TEST_NAME_COLUMN).protection = Protection(locked=False)
        ini_ws.cell(row, DataForm.MOCKTEST_SCORE_COLUMN).protection = Protection(locked=False)
        ini_ws.cell(row, DataForm.MAKEUP_TEST_CHECK_COLUMN).protection = Protection(locked=False)

    if os.path.isfile(f"./데일리테스트 기록 양식({datetime.today().strftime('%m.%d')}).xlsx"):
        i = 1
        while True:
            if not os.path.isfile(f"./데일리테스트 기록 양식({datetime.today().strftime('%m.%d')})({str(i)}).xlsx"):
                ini_wb.save(f"./데일리테스트 기록 양식({datetime.today().strftime('%m.%d')})({str(i)}).xlsx")
                break
            i += 1
    else:
        ini_wb.save(f"./데일리테스트 기록 양식({datetime.today().strftime('%m.%d')}).xlsx")
    gui.q.put("데일리테스트 기록 양식 생성을 완료했습니다.")
    gui.thread_end_flag = True

def save_data(gui:GUI, filepath:str, makeup_test_date:dict):
    form_wb = xl.load_workbook(filepath, data_only=True)
    form_ws = form_wb["데일리테스트 기록 양식"]

    # 올바른 양식이 아닙니다.
    if not data_validation(gui, form_ws):
        gui.q.put("데이터 저장이 중단되었습니다.")
        return
    
    # 학생 정보 열기
    student_wb = xl.load_workbook("./학생 정보.xlsx")
    try:
        student_ws = student_wb["학생 정보"]
    except:
        gui.q.put(r"[오류] '학생 정보.xlsx'의 시트명을")
        gui.q.put(r"'학생 정보'로 변경해 주세요.")
        return

    # 재시험 명단 열기
    if not os.path.isfile("./data/재시험 명단.xlsx"):
        gui.q.put("재시험 명단 파일 생성 중...")

        ini_wb = xl.Workbook()
        ini_ws = ini_wb.active
        ini_ws.title = "재시험 명단"
        ini_ws[gcl(MakeupTestList.TEST_DATE_COLUMN)+"1"] = "응시일"
        ini_ws[gcl(MakeupTestList.CLASS_NAME_COLUMN)+"1"] = "반"
        ini_ws[gcl(MakeupTestList.TEACHER_COLUMN)+"1"] = "담당T"
        ini_ws[gcl(MakeupTestList.STUDENT_NAME_COLUMN)+"1"] = "이름"
        ini_ws[gcl(MakeupTestList.TEST_NAME_COLUMN)+"1"] = "시험명"
        ini_ws[gcl(MakeupTestList.TEST_SCORE_COLUMN)+"1"] = "시험 점수"
        ini_ws[gcl(MakeupTestList.MAKEUP_TEST_WEEK_DATE_COLUMN)+"1"] = "재시 요일"
        ini_ws[gcl(MakeupTestList.MAKEUP_TEST_TIME_COLUMN)+"1"] = "재시 시간"
        ini_ws[gcl(MakeupTestList.MAKEUP_TEST_DATE_COLUMN)+"1"] = "재시 날짜"
        ini_ws[gcl(MakeupTestList.MAKEUP_TEST_SCORE_COLUMN)+"1"] = "재시 점수"
        ini_ws[gcl(MakeupTestList.ETC_COLUMN)+"1"] = "비고"
        ini_ws.auto_filter.ref = "A:"+gcl(MakeupTestList.MAX)
        ini_wb.save("./data/재시험 명단.xlsx")
    makeup_list_wb = xl.load_workbook("./data/재시험 명단.xlsx")
    try:
        makeup_list_ws = makeup_list_wb["재시험 명단"]
    except:
        gui.q.put(r"[오류] '재시험 명단.xlsx'의 시트명을")
        gui.q.put(r"'재시험 명단'으로 변경해 주세요.")
        return
    
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(f"{os.getcwd()}\\data\\{config['dataFileName']}.xlsx")
    except:
        gui.q.put("데이터 파일 창을 끄고 다시 실행해 주세요.")
        return
    wb.Save()
    wb.Close()

    gui.q.put("백업 파일 생성중...")
    data_file_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx")
    data_file_wb.save(f"./data/backup/{config['dataFileName']}({datetime.today().strftime('%Y%m%d')}).xlsx")
    
    gui.q.put("데이터 저장 중...")

    # 재시험 명단 작성 시작 위치
    for i in range(2, makeup_list_ws.max_row + 2):
        if makeup_list_ws.cell(i, MakeupTestList.TEST_DATE_COLUMN).value is None:
            MAKEUP_TEST_RANGE = MAKEUP_TEST_WRITE_ROW = i
            break
    
    for sheet_name in data_file_wb.sheetnames:
        data_file_ws = data_file_wb[sheet_name]
        if sheet_name == "데일리테스트":
            TEST_NAME_COLUMN = DataForm.DAILYTEST_TEST_NAME_COLUMN
            SCORE_COLUMN = DataForm.DAILYTEST_SCORE_COLUMN
        else:
            TEST_NAME_COLUMN = DataForm.MOCKTEST_TEST_NAME_COLUMN
            SCORE_COLUMN = DataForm.MOCKTEST_SCORE_COLUMN

        # 동적 열 탐색
        for i in range(1, data_file_ws.max_column+1):
            if data_file_ws.cell(1, i).value == "반":
                CLASS_NAME_COLUMN = i
                break
        for i in range(1, data_file_ws.max_column+1):
            if data_file_ws.cell(1, i).value == "이름":
                STUDENT_NAME_COLUMN = i
                break
        for i in range(1, data_file_ws.max_column+1):
            if data_file_ws.cell(1, i).value == "학생 평균":
                AVERAGE_SCORE_COLUMN = i
                break
        
        for i in range(2, form_ws.max_row+1): # 데일리데이터 기록 양식 루프
            # 파일 끝 검사
            if form_ws.cell(i, DataForm.STUDENT_NAME_COLUMN).value is None:
                break
            
            # 반 필터링
            if (form_ws.cell(i, DataForm.CLASS_NAME_COLUMN).value is not None) and (form_ws.cell(i, TEST_NAME_COLUMN).value is not None):
                class_name = form_ws.cell(i, DataForm.CLASS_NAME_COLUMN).value
                test_name = form_ws.cell(i, TEST_NAME_COLUMN).value
                teacher = form_ws.cell(i, DataForm.TEACHER_COLUMN).value

                #반 시작 찾기
                for row in range(2, data_file_ws.max_row+1):
                    if data_file_ws.cell(row, CLASS_NAME_COLUMN).value == class_name:
                        CLASS_START = row
                        break
                # 반 끝 찾기
                for row in range(CLASS_START, data_file_ws.max_row+1):
                    if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "시험 평균":
                        CLASS_END = row
                        break
                
                # 데일리테스트 작성 열 위치 찾기
                for col in range(AVERAGE_SCORE_COLUMN+1, data_file_ws.max_column+2):
                    if data_file_ws.cell(CLASS_START, col).value is None:
                        WRITE_COLUMN = col
                        break
                    if data_file_ws.cell(CLASS_START, col).value.strftime("%y.%m.%d") == DATE.today().strftime("%y.%m.%d"):
                        WRITE_COLUMN = col
                        break
                
                # 입력 틀 작성
                AVERAGE_FORMULA = f"=ROUND(AVERAGE({gcl(WRITE_COLUMN)+str(CLASS_START + 2)}:{gcl(WRITE_COLUMN)+str(CLASS_END - 1)}), 0)"
                data_file_ws.column_dimensions[gcl(WRITE_COLUMN)].width = 14
                data_file_ws.cell(CLASS_START, WRITE_COLUMN).value = DATE.today()
                data_file_ws.cell(CLASS_START, WRITE_COLUMN).number_format = "yyyy.mm.dd(aaa)"
                data_file_ws.cell(CLASS_START, WRITE_COLUMN).alignment = Alignment(horizontal="center", vertical="center")

                data_file_ws.cell(CLASS_START + 1, WRITE_COLUMN).value = test_name
                data_file_ws.cell(CLASS_START + 1, WRITE_COLUMN).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                data_file_ws.cell(CLASS_END, WRITE_COLUMN).value = AVERAGE_FORMULA
                data_file_ws.cell(CLASS_END, WRITE_COLUMN).font = Font(bold=True)
                data_file_ws.cell(CLASS_END, WRITE_COLUMN).alignment = Alignment(horizontal="center", vertical="center")
                data_file_ws.cell(CLASS_END, WRITE_COLUMN).border = Border(bottom=Side(border_style="medium", color="000000"))
                
                # 평균 계산 초기화
                score_sum = score_cnt = 0
            
            score = form_ws.cell(i, SCORE_COLUMN).value
            if score is None:
                continue # 점수 없으면 미응시 처리

            for row in range(CLASS_START + 2, CLASS_END):
                if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == form_ws.cell(i, DataForm.STUDENT_NAME_COLUMN).value: # data name == form name
                    data_file_ws.cell(row, WRITE_COLUMN).value = score
                    if type(score) == int:
                        if score < 60:
                            data_file_ws.cell(row, WRITE_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("EC7E31"))
                        elif score < 70:
                            data_file_ws.cell(row, WRITE_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("F5AF85"))
                        elif score < 80:
                            data_file_ws.cell(row, WRITE_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("FCE4D6"))
                        score_sum += score
                        score_cnt += 1
                    data_file_ws.cell(row, WRITE_COLUMN).alignment = Alignment(horizontal="center", vertical="center")
                    break
            
            score_avg = score_sum/score_cnt
            if score_avg < 60:
                data_file_ws.cell(CLASS_END, WRITE_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("EC7E31"))
            elif score_avg < 70:
                data_file_ws.cell(CLASS_END, WRITE_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("F5AF85"))
            elif score_avg < 80:
                data_file_ws.cell(CLASS_END, WRITE_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("FCE4D6"))
            else:
                data_file_ws.cell(CLASS_END, WRITE_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("DDEBF7"))

            # 재시험 작성
            if (type(score) == int) and (score < 80) and (form_ws.cell(i, DataForm.MAKEUP_TEST_CHECK_COLUMN).value != "x") and (form_ws.cell(i, DataForm.MAKEUP_TEST_CHECK_COLUMN).value != "X"):
                check = makeup_list_ws.max_row
                # 재시험 중복 작성 검사
                duplicated = False
                while check >= 1:
                    try:
                        if makeup_list_ws.cell(check, MakeupTestList.TEST_DATE_COLUMN).value is None:
                            check -= 1
                            continue
                        elif makeup_list_ws.cell(check, MakeupTestList.TEST_DATE_COLUMN).value.strftime("%y.%m.%d") == DATE.today().strftime("%y.%m.%d"):
                            if makeup_list_ws.cell(check, MakeupTestList.STUDENT_NAME_COLUMN).value == form_ws.cell(i, DataForm.STUDENT_NAME_COLUMN).value:
                                if makeup_list_ws.cell(check, MakeupTestList.CLASS_NAME_COLUMN).value == class_name:
                                    duplicated = True
                                    break
                        elif makeup_list_ws.cell(check, MakeupTestList.TEST_DATE_COLUMN).value.strftime("%y.%m.%d") == (DATE.today()+timedelta(days=-1)).strftime("%y.%m.%d"):
                            break
                    except:
                        pass
                    check -= 1
                    
                if duplicated: continue
                
                dates = None
                time = None
                new_student = None

                for row in range(2, student_ws.max_row+1):
                    if student_ws.cell(row, StudentInfo.STUDENT_NAME_COLUMN).value == form_ws.cell(i, DataForm.STUDENT_NAME_COLUMN).value:
                        dates = student_ws.cell(row, StudentInfo.MAKEUP_TEST_WEEK_DATE_COLUMN).value
                        time = student_ws.cell(row, StudentInfo.MAKEUP_TEST_TIME_COLUMN).value
                        new_student = student_ws.cell(row, StudentInfo.NEW_STUDENT_CHECK_COLUMN).value
                        break

                makeup_list_ws.cell(MAKEUP_TEST_WRITE_ROW, MakeupTestList.TEST_DATE_COLUMN).value = DATE.today()
                makeup_list_ws.cell(MAKEUP_TEST_WRITE_ROW, MakeupTestList.CLASS_NAME_COLUMN).value = class_name
                makeup_list_ws.cell(MAKEUP_TEST_WRITE_ROW, MakeupTestList.TEACHER_COLUMN).value = teacher
                makeup_list_ws.cell(MAKEUP_TEST_WRITE_ROW, MakeupTestList.STUDENT_NAME_COLUMN).value = form_ws.cell(i, DataForm.STUDENT_NAME_COLUMN).value
                if (new_student is not None) and (new_student == "N"):
                    makeup_list_ws.cell(MAKEUP_TEST_WRITE_ROW, MakeupTestList.STUDENT_NAME_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("FFFF00"))
                makeup_list_ws.cell(MAKEUP_TEST_WRITE_ROW, MakeupTestList.TEST_NAME_COLUMN).value = test_name
                makeup_list_ws.cell(MAKEUP_TEST_WRITE_ROW, MakeupTestList.TEST_SCORE_COLUMN).value = score
                if dates is not None:
                    makeup_list_ws.cell(MAKEUP_TEST_WRITE_ROW, MakeupTestList.MAKEUP_TEST_WEEK_DATE_COLUMN).value = dates
                    date_list = dates.split("/")
                    result = makeup_test_date[date_list[0].replace(" ", "")]
                    for d in date_list:
                        if result > makeup_test_date[d.replace(" ", "")]:
                            result = makeup_test_date[d.replace(" ", "")]
                    if time is not None:
                        makeup_list_ws.cell(MAKEUP_TEST_WRITE_ROW, MakeupTestList.MAKEUP_TEST_TIME_COLUMN).value = time
                    makeup_list_ws.cell(MAKEUP_TEST_WRITE_ROW, MakeupTestList.MAKEUP_TEST_DATE_COLUMN).value = result
                    makeup_list_ws.cell(MAKEUP_TEST_WRITE_ROW, MakeupTestList.MAKEUP_TEST_DATE_COLUMN).number_format = "mm월 dd일(aaa)"
                MAKEUP_TEST_WRITE_ROW += 1

    gui.q.put("재시험 명단 작성 중...")
    # 정렬 및 테두리
    for row in range(MAKEUP_TEST_RANGE, makeup_list_ws.max_row+1):
        if makeup_list_ws.cell(row, 1).value is None: break
        for col in range(1, makeup_list_ws.max_column + 1):
            makeup_list_ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")
            makeup_list_ws.cell(row, col).border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    try:
        data_file_ws = data_file_wb["데일리테스트"]
        data_file_wb.save(f"./data/{config['dataFileName']}.xlsx")
    except:
        gui.q.put("데이터 파일 창을 끄고 다시 실행해 주세요.")
        return

    try:
        makeup_list_wb.save("./data/재시험 명단.xlsx")
    except:
        gui.q.put("재시험 명단 파일 창을 끄고 다시 실행해 주세요.")
        return
    
    gui.q.put("조건부 서식 적용중...")
    try:
        wb = excel.Workbooks.Open(f"{os.getcwd()}\\data\\{config['dataFileName']}.xlsx")
    except:
        gui.q.put("데이터 파일 창을 끄고 다시 실행해 주세요.")
        return
    wb.Save()
    wb.Close()

    data_file_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx")
    data_file_color_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx", data_only=True)

    for sheet_name in data_file_wb.sheetnames:
        data_file_ws = data_file_wb[sheet_name]
        data_file_color_ws = data_file_color_wb[sheet_name]

        for col in range(1, data_file_ws.max_column):
            if data_file_ws.cell(1, col).value == "이름":
                STUDENT_NAME_COLUMN = col
                break
        for col in range(1, data_file_ws.max_column):
            if data_file_ws.cell(1, col).value == "학생 평균":
                AVERAGE_SCORE_COLUMN = col
                break
        for row in range(2, data_file_color_ws.max_row+1):
            # 학생 별 평균 점수에 대한 조건부 서식
            average_score_data = data_file_color_ws.cell(row, AVERAGE_SCORE_COLUMN).value
            average_score_cell = data_file_ws.cell(row, AVERAGE_SCORE_COLUMN)
            student_name_cell = data_file_ws.cell(row, STUDENT_NAME_COLUMN)
            if type(average_score_data) == int:
                if average_score_data < 60:
                    average_score_cell.fill = PatternFill(fill_type="solid", fgColor=Color("EC7E31"))
                elif average_score_data < 70:
                    average_score_cell.fill = PatternFill(fill_type="solid", fgColor=Color("F5AF85"))
                elif average_score_data < 80:
                    average_score_cell.fill = PatternFill(fill_type="solid", fgColor=Color("FCE4D6"))
                elif student_name_cell.value == "시험 평균":
                    average_score_cell.fill = PatternFill(fill_type="solid", fgColor=Color("DDEBF7"))
                else:
                    average_score_cell.fill = PatternFill(fill_type="solid", fgColor=Color("E2EFDA"))
            # 신규생 하이라이트
            student_name_cell.fill = PatternFill(fill_type=None, fgColor=Color("00FFFFFF"))
            if (student_name_cell.value is not None) or (student_name_cell.value != "날짜") or (student_name_cell.value != "시험명") or (student_name_cell.value != "시험 평균"):
                for j in range(2, student_ws.max_row+1):
                    if (student_name_cell.value == student_ws.cell(j, 1).value) and (student_ws.cell(j, StudentInfo.NEW_STUDENT_CHECK_COLUMN).value == "N"):
                        student_name_cell.fill = PatternFill(fill_type="solid", fgColor=Color("FFFF00"))
                        break
    try:
        data_file_ws = data_file_wb["데일리테스트"]
        data_file_wb.save(f"./data/{config['dataFileName']}.xlsx")
    except:
        gui.q.put("데이터 파일 창을 끄고 다시 실행해 주세요.")
        return
    

    gui.q.put("데이터 저장을 완료했습니다.")
    excel.Visible = True
    wb = excel.Workbooks.Open(f"{os.getcwd()}\\data\\{config['dataFileName']}.xlsx")
    pythoncom.CoUninitialize()
    gui.thread_end_flag = True

def send_message(gui:GUI, filepath:str, makeup_test_date:dict):
    form_wb = xl.load_workbook(filepath, data_only=True)
    form_ws = form_wb["데일리테스트 기록 양식"]

    # 올바른 양식이 아닙니다.
    if not data_validation(gui, form_ws):
        gui.q.put("데이터 저장이 중단되었습니다.")
        return

    gui.q.put("크롬을 실행시키는 중...")
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=service, options=options)

    student_wb = xl.load_workbook("./학생 정보.xlsx")
    try:
        student_ws = student_wb["학생 정보"]
    except:
        gui.q.put("[오류] \"학생 정보.xlsx\"의 시트명을")
        gui.q.put("\"학생 정보\"로 변경해 주세요.")
        return
    
    # 아이소식 접속
    driver.get(config["url"])
    driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(config["dailyTest"])
    
    driver.execute_script("window.open(\"" + config["url"] + "\");")
    driver.switch_to.window(driver.window_handles[1])
    driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(config["makeupTest"])

    driver.execute_script("window.open(\"" + config["url"] + "\");")
    driver.switch_to.window(driver.window_handles[2])
    driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(config["makeupTestDate"])

    gui.q.put("메시지 작성 중...")
    for i in range(2, form_ws.max_row+1):
        driver.switch_to.window(driver.window_handles[0])
        name = form_ws.cell(i, DataForm.STUDENT_NAME_COLUMN).value
        daily_test_score = form_ws.cell(i, DataForm.DAILYTEST_SCORE_COLUMN).value
        mock_test_score = form_ws.cell(i, DataForm.MOCKTEST_SCORE_COLUMN).value
        if form_ws.cell(i, DataForm.CLASS_NAME_COLUMN).value is not None:
            class_name = form_ws.cell(i, DataForm.CLASS_NAME_COLUMN).value
            daily_test_name = form_ws.cell(i, DataForm.DAILYTEST_TEST_NAME_COLUMN).value
            mock_test_name = form_ws.cell(i, DataForm.MOCKTEST_TEST_NAME_COLUMN).value
            daily_test_average = form_ws.cell(i, DataForm.DAILYTEST_AVERAGE_COLUMN).value
            mock_test_average = form_ws.cell(i, DataForm.MOCKTEST_AVERAGE_COLUMN).value

        # 시험 미응시시 건너뛰기
        if daily_test_score is not None:
            test_name = daily_test_name
            score = daily_test_score
            average = daily_test_average
        elif mock_test_score is not None:
            test_name = mock_test_name
            score = mock_test_score
            average = mock_test_average
        else:
            continue

        if type(score) != int:
            continue
        
        table_names = driver.find_elements(By.CLASS_NAME, "style1")
        for j in range(len(table_names)):
            if class_name in table_names[j].text:
                index = j
                break
        else:
            continue

        trs = driver.find_element(By.ID, "table_" + str(index)).find_elements(By.CLASS_NAME, "style12")
        for tr in trs:
            if tr.find_element(By.CLASS_NAME, "style9").text == name:
                tds = tr.find_elements(By.TAG_NAME, "td")
                tds[0].find_element(By.TAG_NAME, "input").send_keys(test_name)
                tds[1].find_element(By.TAG_NAME, "input").send_keys(score)
                tds[2].find_element(By.TAG_NAME, "input").send_keys(average)
                break
        
        if (type(score) == int) and (score < 80) and (form_ws.cell(i, 12).value != "x") and (form_ws.cell(i, 12).value != "X"):
            for j in range(2, student_ws.max_row+1):
                if student_ws.cell(j, 1).value == name:
                    date = student_ws.cell(j, 4).value
                    time = student_ws.cell(j, 5).value
                    break
            if date is None:
                driver.switch_to.window(driver.window_handles[1])
                trs = driver.find_element(By.ID, "table_" + str(index)).find_elements(By.CLASS_NAME, "style12")
                for tr in trs:
                    if tr.find_element(By.CLASS_NAME, "style9").text == name:
                        tds = tr.find_elements(By.TAG_NAME, "td")
                        tds[0].find_element(By.TAG_NAME, "input").send_keys(test_name)
            else:
                date_list = date.split("/")
                result = makeup_test_date[date_list[0].replace(" ", "")]
                timeIndex = 0
                for i in range(len(date_list)):
                    if result > makeup_test_date[date_list[i].replace(" ", "")]:
                        result = makeup_test_date[date_list[i].replace(" ", "")]
                        timeIndex = i
                driver.switch_to.window(driver.window_handles[2])
                trs = driver.find_element(By.ID, "table_" + str(index)).find_elements(By.CLASS_NAME, "style12")
                for tr in trs:
                    if tr.find_element(By.CLASS_NAME, "style9").text == name:
                        tds = tr.find_elements(By.TAG_NAME, "td")
                        tds[0].find_element(By.TAG_NAME, "input").send_keys(test_name)
                        try:
                            if time is not None:
                                if "/" in str(time):
                                    tds[1].find_element(By.TAG_NAME, "input").send_keys(result.strftime("%m월 %d일") + " " + str(time).split("/")[timeIndex] + "시")
                                else:
                                    tds[1].find_element(By.TAG_NAME, "input").send_keys(result.strftime("%m월 %d일") + " " + str(time)+ "시")
                        except:
                            gui.q.put(name + "의 재시험 일정을 요일별 시간으로 설정하거나")
                            gui.q.put("하나의 시간으로 통일해 주세요.")
                            gui.q.put("중단되었습니다.")
                            driver.quit()
                            gui.thread_end_flag = True
                            return

    gui.q.put("메시지 입력을 완료했습니다.")
    gui.q.put("메시지 확인 후 전송해주세요.")
    gui.thread_end_flag = True

def apply_color(gui:GUI):
    student_wb = xl.load_workbook("./학생 정보.xlsx")
    try:
        student_ws = student_wb["학생 정보"]
    except:
        gui.q.put(r"[오류] '학생 정보.xlsx'의 시트명을")
        gui.q.put(r"'학생 정보'로 변경해 주세요.")
        return
    
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(f"{os.getcwd()}\\data\\{config['dataFileName']}.xlsx")
    except:
        gui.q.put("데이터 파일 창을 끄고 다시 실행해 주세요.")
        return
    wb.Save()
    wb.Close()

    gui.q.put("조건부 서식 적용중...")

    data_file_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx")
    data_file_color_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx", data_only=True)
    
    for sheet_name in data_file_wb.sheetnames:
        data_file_ws = data_file_wb[sheet_name]
        data_file_color_ws = data_file_color_wb[sheet_name]

        for i in range(1, data_file_ws.max_column):
            if data_file_ws.cell(1, i).value == "이름":
                STUDENT_NAME_COLUMN = i
                break
        for i in range(1, data_file_ws.max_column):
            if data_file_ws.cell(1, i).value == "학생 평균":
                AVERAGE_SCORE_COLUMN = i
                break
        
        for i in range(2, data_file_color_ws.max_row+1):
            if data_file_color_ws.cell(i, STUDENT_NAME_COLUMN).value is None:
                break
            for j in range(1, data_file_color_ws.max_column+1):
                data_file_ws.column_dimensions[gcl(j)].width = 14
                if data_file_ws.cell(i, STUDENT_NAME_COLUMN).value == "시험 평균" and data_file_ws.cell(i, j).value is not None:
                    data_file_ws.cell(i, j).border = Border(bottom=Side(border_style="medium", color="000000"))
                if j > AVERAGE_SCORE_COLUMN:    
                    if data_file_ws.cell(i, STUDENT_NAME_COLUMN).value == "날짜" and data_file_ws.cell(i, j).value is not None:
                        data_file_ws.cell(i, j).border = Border(top=Side(border_style="medium", color="000000"))
                    if type(data_file_color_ws.cell(i, j).value) == int:
                        if data_file_color_ws.cell(i, j).value < 60:
                            data_file_ws.cell(i, j).fill = PatternFill(fill_type="solid", fgColor=Color("EC7E31"))
                        elif data_file_color_ws.cell(i, j).value < 70:
                            data_file_ws.cell(i, j).fill = PatternFill(fill_type="solid", fgColor=Color("F5AF85"))
                        elif data_file_color_ws.cell(i, j).value < 80:
                            data_file_ws.cell(i, j).fill = PatternFill(fill_type="solid", fgColor=Color("FCE4D6"))
                        elif data_file_color_ws.cell(i, STUDENT_NAME_COLUMN).value == "시험 평균":
                            data_file_ws.cell(i, j).fill = PatternFill(fill_type="solid", fgColor=Color("DDEBF7"))
                        else:
                            data_file_ws.cell(i, j).fill = PatternFill(fill_type=None, fgColor=Color("00FFFFFF"))
                    else:
                        data_file_ws.cell(i, j).fill = PatternFill(fill_type=None, fgColor=Color("00FFFFFF"))
                    if data_file_color_ws.cell(i, STUDENT_NAME_COLUMN).value == "시험 평균":
                        data_file_ws.cell(i, j).font = Font(bold=True)

            # 학생별 평균 조건부 서식
            data_file_ws.cell(i, AVERAGE_SCORE_COLUMN).font = Font(bold=True)
            if type(data_file_color_ws.cell(i, AVERAGE_SCORE_COLUMN).value) == int:
                if data_file_color_ws.cell(i, AVERAGE_SCORE_COLUMN).value < 60:
                    data_file_ws.cell(i, AVERAGE_SCORE_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("EC7E31"))
                elif data_file_color_ws.cell(i, AVERAGE_SCORE_COLUMN).value < 70:
                    data_file_ws.cell(i, AVERAGE_SCORE_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("F5AF85"))
                elif data_file_color_ws.cell(i, AVERAGE_SCORE_COLUMN).value < 80:
                    data_file_ws.cell(i, AVERAGE_SCORE_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("FCE4D6"))
                elif data_file_color_ws.cell(i, STUDENT_NAME_COLUMN).value == "시험 평균":
                    data_file_ws.cell(i, AVERAGE_SCORE_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("DDEBF7"))
                else:
                    data_file_ws.cell(i, AVERAGE_SCORE_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("E2EFDA"))
            name = data_file_ws.cell(i, STUDENT_NAME_COLUMN)
            name.fill = PatternFill(fill_type=None, fgColor=Color("00FFFFFF"))
            if (name.value is not None) or (name.value != "날짜") or (name.value != "시험명") or (name.value != "시험 평균"):
                for j in range(2, student_ws.max_row+1):
                    if (name.value == student_ws.cell(j, 1).value) and (student_ws.cell(j, StudentInfo.NEW_STUDENT_CHECK_COLUMN).value == "N"):
                        name.fill = PatternFill(fill_type="solid", fgColor=Color("FFFF00"))
                        break
    
    try:
        data_file_wb.save(f"./data/{config['dataFileName']}.xlsx")
        gui.q.put("조건부 서식 지정을 완료했습니다.")
        excel.Visible = True
        wb = excel.Workbooks.Open(f"{os.getcwd()}\\data\\{config['dataFileName']}.xlsx")
        pythoncom.CoUninitialize()
    except:
        gui.q.put("데이터 파일 창을 끄고 다시 실행해 주세요.")
        return

def delete_student(gui:GUI, student:str):
    data_file_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx")
    # 데이터 파일 취소선
    for sheet_name in data_file_wb.sheetnames:
        data_file_ws = data_file_wb[sheet_name]

        for col in range(1, data_file_ws.max_column+1):
            if data_file_ws.cell(1, col).value == "이름":
                STUDENT_NAME_COLUMN = col
                break
        for row in range(2, data_file_ws.max_row+1):
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == student:
                for col in range(1, data_file_ws.max_column+1):
                    data_file_ws.cell(row, col).font = Font(strike=True)
    
    student_wb = xl.load_workbook("./학생 정보.xlsx")
    try:
        student_ws = student_wb["학생 정보"]
    except:
        gui.q.put(r"[오류] '학생 정보.xlsx'의 시트명을")
        gui.q.put(r"'학생 정보'로 변경해 주세요.")
        return
    # 학생 정보 삭제
    for row in range(2, student_ws.max_row+1):
        if student_ws.cell(row, StudentInfo.STUDENT_NAME_COLUMN).value == student:
            student_ws.delete_rows(row)
            break

    data_file_wb.save(f"./data/{config['dataFileName']}.xlsx")
    student_wb.save("./학생 정보.xlsx")
    gui.q.put(f"{student} 학생을 퇴원 처리하였습니다.")
    gui.thread_end_flag = True
    return

def add_student(gui:GUI, student:str, target_class:str):
    if not check_student_exists(gui, student, target_class):
        gui.q.put(r"아이소식 해당 반에 학생이 업데이트되지 않아")
        gui.q.put(r"신규생 추가를 중단합니다.")
        gui.thread_end_flag = True
        return
    
    # 학생 정보 파일에 학생 추가
    gui.q.put(r"학생 정보 파일에 학생 추가 중...")
    student_wb = xl.load_workbook("./학생 정보.xlsx")
    try:
        student_ws = student_wb["학생 정보"]
    except:
        gui.q.put(r"[오류] '학생 정보.xlsx'의 시트명을")
        gui.q.put(r"'학생 정보'로 변경해 주세요.")
        return
    for row in range(2, student_ws.max_row+2):
        if student_ws.cell(row, StudentInfo.STUDENT_NAME_COLUMN).value is None:
            student_ws.cell(row, StudentInfo.STUDENT_NAME_COLUMN).value = student
            student_ws.cell(row, StudentInfo.CLASS_NAME_COLUMN).value = target_class
            student_ws.cell(row, StudentInfo.NEW_STUDENT_CHECK_COLUMN).value = "N"
            for col in range(1, StudentInfo.MAX+1):
                student_ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")
                student_ws.cell(row, col).border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
            break
    student_wb.save("./학생 정보.xlsx")

    # 데이터파일에 학생 추가
    gui.q.put(r"데이터 파일에 학생 추가 중...")
    data_file_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx")
    for sheet_name in data_file_wb.sheetnames:
        data_file_ws = data_file_wb[sheet_name]
        for i in range(1, data_file_ws.max_column+1):
            temp = data_file_ws.cell(1, i).value
            if temp == "시간":
                TEST_TIME_COLUMN = i
            elif temp == "요일":
                DATE_COLUMN = i
            elif temp == "반":
                CLASS_NAME_COLUMN = i
            elif temp == "담당":
                TEACHER_COLUMN = i
            elif temp == "이름":
                STUDENT_NAME_COLUMN = i
            elif temp == "학생 평균":
                AVERAGE_SCORE_COLUMN = i
        
        for row in range(2, data_file_ws.max_row+1):
            if data_file_ws.cell(row, CLASS_NAME_COLUMN).value == target_class:
                class_index = row+2
                break
        else: continue # 목표 반이 없으면 건너뛰기

        while data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).value != "시험 평균":
            if data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).value > student:
                break
            elif data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).value == student:
                gui.q.put(f"{student} 학생이 이미 존재합니다.")
                gui.q.put(r"신규생 추가를 중단합니다.")
                gui.thread_end_flag = True
                return
            else: class_index += 1
        data_file_ws.insert_rows(class_index)
        copy_cell(data_file_ws.cell(class_index, TEST_TIME_COLUMN), data_file_ws.cell(class_index-1, TEST_TIME_COLUMN))
        copy_cell(data_file_ws.cell(class_index, DATE_COLUMN), data_file_ws.cell(class_index-1, DATE_COLUMN))
        copy_cell(data_file_ws.cell(class_index, CLASS_NAME_COLUMN), data_file_ws.cell(class_index-1, CLASS_NAME_COLUMN))
        copy_cell(data_file_ws.cell(class_index, TEACHER_COLUMN), data_file_ws.cell(class_index-1, TEACHER_COLUMN))

        data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).value = student
        data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).alignment = Alignment(horizontal="center", vertical="center")
        data_file_ws.cell(class_index, AVERAGE_SCORE_COLUMN).alignment = Alignment(horizontal="center", vertical="center")
        data_file_ws.cell(class_index, AVERAGE_SCORE_COLUMN).font = Font(bold=True)

    data_file_wb.save(f"./data/{config['dataFileName']}.xlsx")

    rescoping_formula()

    gui.q.put(f"{student} 학생을 {target_class} 반에 추가하였습니다.")

    gui.thread_end_flag = True
    return

def move_student(gui:GUI, student:str, target_class:str, current_class:str):
    if not check_student_exists(gui, student, target_class):
        gui.q.put(r"아이소식 해당 반에 학생이 업데이트되지 않아")
        gui.q.put(r"학생 반 이동을 중단합니다.")
        gui.thread_end_flag = True
        return
    
    data_file_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx")
    for sheet_name in data_file_wb.sheetnames:
        data_file_ws = data_file_wb[sheet_name]
        for i in range(1, data_file_ws.max_column+1):
            temp = data_file_ws.cell(1, i).value
            if temp == "시간":
                TEST_TIME_COLUMN = i
            elif temp == "요일":
                DATE_COLUMN = i
            elif temp == "반":
                CLASS_NAME_COLUMN = i
            elif temp == "담당":
                TEACHER_COLUMN = i
            elif temp == "이름":
                STUDENT_NAME_COLUMN = i
            elif temp == "학생 평균":
                AVERAGE_SCORE_COLUMN = i
        
        # 기존 반 데이터 빨간색 처리
        for row in range(2, data_file_ws.max_row+1):
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == student and data_file_ws.cell(row, CLASS_NAME_COLUMN).value == current_class:
                for col in range(1, data_file_ws.max_column+1):
                    data_file_ws.cell(row, col).font = Font(color="FF0000")
        
        # 목표 반에 학생 추가
        for row in range(2, data_file_ws.max_row+1):
            if data_file_ws.cell(row, CLASS_NAME_COLUMN).value == target_class:
                class_index = row+2
                break
        else: continue # 목표 반이 없으면 건너뛰기

        while data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).value != "시험 평균":
            if data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).value > student:
                break
            elif data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).value == student:
                gui.q.put(f"{student} 학생이 이미 존재합니다.")
                gui.q.put(r"학생 반 이동을 중단합니다.")
                gui.thread_end_flag = True
                return
            else: class_index += 1
        data_file_ws.insert_rows(class_index)
        copy_cell(data_file_ws.cell(class_index, TEST_TIME_COLUMN), data_file_ws.cell(class_index-1, TEST_TIME_COLUMN))
        copy_cell(data_file_ws.cell(class_index, DATE_COLUMN), data_file_ws.cell(class_index-1, DATE_COLUMN))
        copy_cell(data_file_ws.cell(class_index, CLASS_NAME_COLUMN), data_file_ws.cell(class_index-1, CLASS_NAME_COLUMN))
        copy_cell(data_file_ws.cell(class_index, TEACHER_COLUMN), data_file_ws.cell(class_index-1, TEACHER_COLUMN))

        data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).value = student
        data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).alignment = Alignment(horizontal="center", vertical="center")
        data_file_ws.cell(class_index, AVERAGE_SCORE_COLUMN).alignment = Alignment(horizontal="center", vertical="center")
        data_file_ws.cell(class_index, AVERAGE_SCORE_COLUMN).font = Font(bold=True)

    data_file_wb.save(f"./data/{config['dataFileName']}.xlsx")

    rescoping_formula()

    gui.q.put(f"{student} 학생을 {current_class} 반에서")
    gui.q.put(f"{target_class} 반으로 이동하였습니다.")
    gui.thread_end_flag = True
    return

def data_validation(gui:GUI, form_ws:Worksheet) -> bool:
    gui.q.put("양식이 올바른지 확인 중...")
    if (form_ws.title != "데일리테스트 기록 양식"):
        gui.q.put("올바른 기록 양식이 아닙니다.")
        return False
    
    form_checked = True
    dailytest_checked = False
    mocktest_checked = False
    for i in range(1, form_ws.max_row+1):
        if form_ws.cell(i, DataForm.CLASS_NAME_COLUMN).value is not None:
            class_name = form_ws.cell(i, DataForm.CLASS_NAME_COLUMN).value
            dailytest_checked = False
            mocktest_checked = False
            dailytest_name = form_ws.cell(i, DataForm.DAILYTEST_TEST_NAME_COLUMN).value
            mocktest_name = form_ws.cell(i, DataForm.MOCKTEST_TEST_NAME_COLUMN).value
        
        if dailytest_checked and mocktest_checked: continue
        
        if not dailytest_checked and form_ws.cell(i, DataForm.DAILYTEST_SCORE_COLUMN).value is not None and dailytest_name is None:
            gui.q.put(f"{class_name}의 시험명이 작성되지 않았습니다.")
            dailytest_checked = True
            form_checked = False
        if not mocktest_checked and form_ws.cell(i, DataForm.MOCKTEST_SCORE_COLUMN).value is not None and mocktest_name is None:
            gui.q.put(f"{class_name}의 모의고사명이 작성되지 않았습니다.")
            mocktest_checked = True
            form_checked = False

    return form_checked

def copy_cell(destination:Cell, source:Cell):
    destination.value = source.value
    destination.font = copy(source.font)
    destination.fill = copy(source.fill)
    destination.border = copy(source.border)
    destination.alignment = copy(source.alignment)
    destination.number_format = copy(source.number_format)

def rescoping_formula():
    data_file_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx")
    for sheet_name in data_file_wb.sheetnames:
        data_file_ws = data_file_wb[sheet_name]
        for i in range(1, data_file_ws.max_column+1):
            temp = data_file_ws.cell(1, i).value
            if temp == "이름":
                STUDENT_NAME_COLUMN = i
            elif temp == "학생 평균":
                AVERAGE_SCORE_COLUMN = i

        # 평균 범위 재지정
        for row in range(2, data_file_ws.max_row+1):
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "날짜":
                class_start = row+2
            elif data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "시험 평균":
                class_end = row-1
                data_file_ws[f"{gcl(AVERAGE_SCORE_COLUMN)}{str(row)}"] = ArrayFormula(f"{gcl(AVERAGE_SCORE_COLUMN)}{str(row)}", f"=ROUND(AVERAGE(IFERROR({gcl(AVERAGE_SCORE_COLUMN)}{str(class_start)}:{gcl(AVERAGE_SCORE_COLUMN)}{str(class_end)}, \"\")), 0)")
                if class_start >= class_end: continue
                for col in range(AVERAGE_SCORE_COLUMN+1, data_file_ws.max_column+1):
                    if data_file_ws.cell(class_start-2, col).value is None: break
                    data_file_ws.cell(row, col).value = f"=ROUND(AVERAGE({gcl(col)}{str(class_start)}:{gcl(col)}{str(class_end)}), 0)"
            elif data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "시험명": continue
            else:
                data_file_ws.cell(row, AVERAGE_SCORE_COLUMN).value = f"=ROUND(AVERAGE({gcl(AVERAGE_SCORE_COLUMN+1)}{str(row)}:XFD{str(row)}), 0)"

    data_file_wb.save(f"./data/{config['dataFileName']}.xlsx")

def check_student_exists(gui:GUI, target_student_name:str, target_class_name:str):
    gui.q.put(r"아이소식으로부터 정보 받아오는 중...")
    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    driver = webdriver.Chrome(service = service, options = options)

    # 아이소식 접속
    driver.get(config["url"])
    table_names = driver.find_elements(By.CLASS_NAME, "style1")

    # 반 루프
    for i in range(3, len(table_names)):
        if target_class_name == table_names[i].text.rstrip():
            trs = driver.find_element(By.ID, "table_" + str(i)).find_elements(By.CLASS_NAME, "style12")
            for tr in trs:
                if target_student_name == tr.find_element(By.CLASS_NAME, "style9").text:
                    return True
    return False

ui = tk.Tk()
gui = GUI(ui)
ui.after(100, gui.thread_log)
ui.after(100, gui.check_files)
ui.after(100, gui.check_thread_end)
ui.mainloop()
