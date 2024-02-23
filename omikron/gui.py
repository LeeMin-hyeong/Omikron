import os
import queue
import threading
import tkinter as tk
import win32com.client

from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from tkinter import ttk, filedialog
from tkinter.messagebox import askokcancel, askyesno
from webbrowser import open_new

import omikron.chrome
import omikron.classinfo
import omikron.config
import omikron.datafile
import omikron.dataform
import omikron.makeuptest
import omikron.studentinfo
import omikron.thread

from omikron.defs import VERSION, ClassInfo, StudentInfo, MakeupTestList
from omikron.log import OmikronLog

class GUI():
    def __init__(self, ui:tk.Tk):
        self.ui = ui
        
        # log queue
        self.log_q = OmikronLog.log_queue
        
        # 작업 종료 플래그
        self.thread_end_flag = omikron.thread.thread_end_flag

        # 창 크기
        self.width = 320
        self.height = 585 # button +25

        # 창 위치
        self.x = int((self.ui.winfo_screenwidth()/4) - (self.width/2))
        self.y = int((self.ui.winfo_screenheight()/2) - (self.height/2))

        self.ui.geometry(f"{self.width}x{self.height}+{self.x}+{self.y}")
        self.ui.title(VERSION)
        self.ui.resizable(False, False)

        self.makeup_test_date = None
        """요일별 재시험 날짜"""

        tk.Label(self.ui, text="Omikron 데이터 프로그램").pack()
        
        # Notion 사용 설명서 하이퍼링크
        def callback(url:str):
            open_new(url)
        link = tk.Label(self.ui, text="[ 사용법 및 도움말 ]", cursor="hand2")
        link.pack()
        link.bind("<Button-1>", lambda _: callback("https://omikron-db.notion.site/ad673cca64c146d28adb3deaf8c83a0d?pvs=4"))

        # 메세지 창
        self.scroll = tk.Scrollbar(self.ui, orient="vertical")
        self.log = tk.Listbox(self.ui, yscrollcommand=self.scroll.set, width=51, height=5)
        self.scroll.config(command=self.log.yview)
        self.log.pack()
        
        # 버튼
        tk.Label(self.ui, text="< 기수 변경 관련 >").pack()

        self.make_class_info_file_button = tk.Button(self.ui, cursor="hand2", text="반 정보 기록 양식 생성", width=40, command=self.make_class_info_file_task)
        self.make_class_info_file_button.pack()

        self.make_data_file_button = tk.Button(self.ui, cursor="hand2", text="데이터 파일 생성", width=40, command=self.data_file_task)
        self.make_data_file_button.pack()

        self.make_student_info_file_button = tk.Button(self.ui, cursor="hand2", text="학생 정보 기록 양식 생성", width=40, command=self.student_info_file_task)
        self.make_student_info_file_button.pack()

        self.update_class_button = tk.Button(self.ui, cursor="hand2", text="반 업데이트", width=40, command=self.update_class_task)
        self.update_class_button.pack()

        tk.Label(self.ui, text="\n< 데이터 저장 및 문자 전송 >").pack()

        self.make_data_form_button = tk.Button(self.ui, cursor="hand2", text="데일리 테스트 기록 양식 생성", width=40, command=self.make_data_form_task)
        self.make_data_form_button.pack()

        self.save_test_data_button = tk.Button(self.ui, cursor="hand2", text="데이터 저장", width=40, command=self.save_test_data_task)
        self.save_test_data_button.pack()

        self.send_message_button = tk.Button(self.ui, cursor="hand2", text="시험 결과 전송", width=40, command=self.send_message_task)
        self.send_message_button.pack()

        self.save_individual_test_button = tk.Button(self.ui, cursor="hand2", text="개별 시험 기록", width=40, command=self.save_individual_test_task)
        self.save_individual_test_button.pack()

        self.save_makeup_test_result_button = tk.Button(self.ui, cursor="hand2", text="재시험 기록", width=40, command=self.save_makeup_test_result_task)
        self.save_makeup_test_result_button.pack()

        tk.Label(self.ui, text="\n< 데이터 관리 >").pack()

        self.apply_color_button = tk.Button(self.ui, cursor="hand2", text="데이터 파일 조건부 서식 재지정", width=40, command=self.conditional_formatting_task)
        self.apply_color_button.pack()

        tk.Label(self.ui, text="< 학생 관리 >").pack()
        self.add_student_button = tk.Button(self.ui, cursor="hand2", text="신규생 추가", width=40, command=self.add_student_task)
        self.add_student_button.pack()

        self.delete_student_button = tk.Button(self.ui, cursor="hand2", text="퇴원 처리", width=40, command=self.delete_student_task)
        self.delete_student_button.pack()

        self.move_student_button = tk.Button(self.ui, cursor="hand2", text="학생 반 이동", width=40, command=self.move_student_task)
        self.move_student_button.pack()

    # ui functions
    def print_log(self):
        """
        OmikronLog.log_queue 의 내용을 확인하여 큐가 비어있지 않다면 화면에 로그 출력

        호출 주기 : 10ms
        """
        try:
            msg = self.log_q.get(block=False)
        except queue.Empty:
            self.ui.after(10, self.print_log)
            return
        self.log.insert(tk.END, msg)
        self.log.update()
        self.log.see(tk.END)
        self.ui.after(10, self.print_log)

    def check_files(self):
        """
        파일 존재 여부를 확인하여 버튼의 활성화 여부 결정

        호출 주기 : 100ms
        """
        classinfo_check = studentinfo_check = datafile_check = False

        if os.path.isfile(f"{ClassInfo.DEFAULT_NAME}.xlsx"):
            self.make_class_info_file_button["state"]   = tk.DISABLED
            self.make_student_info_file_button["state"] = tk.NORMAL
            self.make_data_file_button["state"]         = tk.NORMAL
            classinfo_check = True
        else:
            self.make_class_info_file_button["state"]   = tk.NORMAL
            self.make_student_info_file_button["state"] = tk.DISABLED
            self.make_data_file_button["state"]         = tk.DISABLED
        if os.path.isfile(f"{StudentInfo.DEFAULT_NAME}.xlsx"):
            self.make_student_info_file_button["text"] = "학생 정보 업데이트"
            studentinfo_check = True
        else: 
            self.make_student_info_file_button["text"] = "학생 정보 기록 양식 생성"
        if os.path.isfile(f"./data/{omikron.config.DATA_FILE_NAME}.xlsx"):
            self.make_data_file_button["text"] = "데이터 파일 이름 변경"
            datafile_check = True
        else:
            self.make_data_file_button["text"] = "데이터 파일 생성"

        if classinfo_check and studentinfo_check and datafile_check:
            self.update_class_button["state"]            = tk.NORMAL
            self.make_data_form_button["state"]          = tk.NORMAL
            self.save_test_data_button["state"]          = tk.NORMAL
            self.send_message_button["state"]            = tk.NORMAL
            self.save_individual_test_button["state"]    = tk.NORMAL
            self.apply_color_button["state"]             = tk.NORMAL
            self.add_student_button["state"]             = tk.NORMAL
            self.delete_student_button["state"]          = tk.NORMAL
            self.move_student_button["state"]            = tk.NORMAL
            if os.path.isfile(f"./data/{MakeupTestList.DEFAULT_NAME}.xlsx"):
                self.save_makeup_test_result_button["state"] = tk.NORMAL
            else:
                self.save_makeup_test_result_button["state"] = tk.DISABLED
        else:
            self.update_class_button["state"]            = tk.DISABLED
            self.make_data_form_button["state"]          = tk.DISABLED
            self.save_test_data_button["state"]          = tk.DISABLED
            self.send_message_button["state"]            = tk.DISABLED
            self.save_individual_test_button["state"]    = tk.DISABLED
            self.save_makeup_test_result_button["state"] = tk.DISABLED
            self.apply_color_button["state"]             = tk.DISABLED
            self.add_student_button["state"]             = tk.DISABLED
            self.delete_student_button["state"]          = tk.DISABLED
            self.move_student_button["state"]            = tk.DISABLED

        if os.path.isfile(f"./{ClassInfo.TEMP_FILE_NAME}.xlsx"):
            self.update_class_button["text"] = "반 정보 수정 후 반 업데이트 계속하기"
        else:
            self.update_class_button["text"] = "반 업데이트"
        
        self.ui.after(100, self.check_files)

    def check_thread_end(self):
        """
        thread가 종료되면 창을 활성화

        호출 주기 : 100ms
        """
        if omikron.thread.thread_end_flag:
            omikron.thread.thread_end_flag = False
            self.ui.wm_attributes("-topmost", 1)
            self.ui.wm_attributes("-topmost", 0)
        self.ui.after(100, self.check_thread_end)

    # dialogs
    def change_data_file_name_dialog(self):
        def quit_event():
            popup.quit()
            popup.destroy()

        popup = tk.Toplevel(self.ui)
        width = 250
        height = 120
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("데이터 파일 이름 변경")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quit_event)

        tk.Label(popup, text=f"기존 파일명: {omikron.config.DATA_FILE_NAME}", pady=10).pack()

        new_filename_var = tk.StringVar()
        tk.Entry(popup, textvariable=new_filename_var, width=28).pack()

        tk.Label(popup).pack()
        tk.Button(popup, text="데이터 파일 이름 변경", width=20 , command=quit_event).pack()
        
        popup.mainloop()

        return new_filename_var.get()

    def holiday_dialog(self):
        """
        휴일 지정 팝업창

        휴일 정보를 바탕으로 요일별 재시험 날짜 생성
        """
        def quit_event():
            for i in range(7):
                if var_list[i].get():
                    makeup_test_date[weekday[i]] += timedelta(days=7)
            popup.quit()
            popup.destroy()

        popup = tk.Toplevel(self.ui)
        width = 200
        height = 300
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("휴일 선택")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quit_event)

        today = datetime.today().date()
        weekday = ("월", "화", "수", "목", "금", "토", "일")
        makeup_test_date = {weekday[i] : today + relativedelta(weekday=i) for i in range(7)}
        for key, value in makeup_test_date.items():
            if value == today:
                makeup_test_date[key] += timedelta(days=7)

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
            tk.Checkbutton(popup, text=f"{makeup_test_date[weekday[(sort+i)%7]]} {weekday[(sort+i)%7]}", variable=var_list[(sort+i)%7]).pack()
        tk.Label(popup, text="\n").pack()
        tk.Button(popup, text="확인", width=10 , command=quit_event).pack()
        
        popup.mainloop()    
        
        self.makeup_test_date = makeup_test_date

    def delete_student_dialog(self):
        """
        퇴원 처리 학생 선택 팝업

        return 성공 여부, 학생 이름
        """
        def quit_event():
            popup.quit()
            popup.destroy()

        popup = tk.Toplevel(self.ui)
        width = 250
        height = 140
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("퇴원 관리")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quit_event)

        complete, class_student_dict, _ = omikron.datafile.get_data_sorted_dict()
        if not complete: return False, None

        tk.Label(popup).pack()

        def class_selection_call_back(event):
            selected_class_name = event.widget.get()
            student_combo.set("학생 선택")
            student_combo["values"] = list(class_student_dict[selected_class_name].keys())
        class_combo = ttk.Combobox(popup, values=list(class_student_dict.keys()), state="readonly", width=25)
        class_combo.set("반 선택")
        class_combo.bind("<<ComboboxSelected>>", class_selection_call_back)
        class_combo.pack()

        tk.Label(popup).pack()
        selected_student = tk.StringVar()
        student_combo = ttk.Combobox(popup, values=None, state="readonly", textvariable=selected_student, width=25)
        student_combo.set("학생 선택")
        student_combo.pack()

        tk.Label(popup).pack()

        tk.Button(popup, text="퇴원", width=10, command=quit_event).pack()
        
        popup.mainloop()
        
        student_name = selected_student.get()

        if student_name == "학생 선택":
            return False, None

        return True, student_name

    def move_student_dialog(self):
        """
        학생 반 이동 처리 팝업

        return 성공 여부, 학생 이름, 목표 반, 현재 반
        """
        def quit_event():
            popup.quit()
            popup.destroy()

        popup = tk.Toplevel(self.ui)
        width = 250
        height = 160
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("학생 반 이동")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quit_event)

        complete, class_student_dict, _ = omikron.datafile.get_data_sorted_dict()
        if not complete: return False, None, None, None

        tk.Label(popup).pack()

        def class_selection_call_back(event):
            class_name = event.widget.get()
            student_combo.set("학생 선택")
            student_combo["values"] = list(class_student_dict[class_name].keys())

        current_class_var = tk.StringVar()
        current_class_combo = ttk.Combobox(popup, values=list(class_student_dict.keys()), state="readonly", textvariable=current_class_var, width=25)
        current_class_combo.set("반 선택")
        current_class_combo.bind("<<ComboboxSelected>>", class_selection_call_back)
        current_class_combo.pack()

        selected_student = tk.StringVar()
        student_combo = ttk.Combobox(popup, values=None, state="readonly", textvariable=selected_student, width=25)
        student_combo.set("학생 선택")
        student_combo.pack()

        tk.Label(popup).pack()
        target_class_var = tk.StringVar()
        current_class_combo = ttk.Combobox(popup, values=list(class_student_dict.keys()), state="readonly", textvariable=target_class_var, width=25)
        current_class_combo.set("이동할 반 선택")
        current_class_combo.pack()

        tk.Label(popup).pack()

        tk.Button(popup, text="반 이동", width=10 , command=quit_event).pack()
        
        popup.mainloop()

        target_student_name = selected_student.get()
        target_class_name   = target_class_var.get()
        current_class_name  = current_class_var.get()

        if target_student_name == "학생 선택" or target_class_name == "이동할 반 선택":
            return False, None, None, None

        return True, target_student_name, target_class_name, current_class_name

    def add_student_dialog(self):
        """
        신규생 추가 팝업

        return 작업 성공 여부, 학생 이름, 목표 반
        """
        def quit_event():
            popup.quit()
            popup.destroy()

        popup = tk.Toplevel(self.ui)
        width = 250
        height = 140
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("신규생 추가")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quit_event)

        class_wb = omikron.classinfo.open()
        complete, class_ws = omikron.classinfo.open_worksheet(class_wb)
        if not complete: return False, None, None

        class_names = omikron.classinfo.get_class_names(class_ws)
        omikron.classinfo.close(class_wb)

        tk.Label(popup).pack()
        target_class_var = tk.StringVar()
        class_combo = ttk.Combobox(popup, values=class_names, state="readonly", textvariable=target_class_var, width=25)
        class_combo.set("학생을 추가할 반 선택")
        class_combo.pack()

        tk.Label(popup).pack()
        target_student_var = tk.StringVar()
        tk.Entry(popup, textvariable=target_student_var, width=28).pack()

        tk.Label(popup).pack()
        tk.Button(popup, text="신규생 추가", width=10 , command=quit_event).pack()

        popup.mainloop()

        target_class_name   = target_class_var.get()
        target_student_name = target_student_var.get()
        if target_class_name == "학생을 추가할 반 선택" or target_student_name == "":
            return False, None, None
        else:
            return True, target_student_name, target_class_name

    def save_individual_test_dialog(self):
        """ 
        결석 등의 사유로 응시하지 않은 시험을 기록하는 경우

        데이터 저장 및 문자 작성 팝업

        return 성공 여부, 학생 이름, 목표 반, 시험 이름, 작성 행, 작성 열, 시험 점수, 재시험 미응시 여부
        """
        def quit_event():
            popup.quit()
            popup.destroy()

        popup = tk.Toplevel(self.ui)
        width = 250
        height = 170
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("개별 점수 기록")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quit_event)
        
        complete, class_student_dict, class_test_dict = omikron.datafile.get_data_sorted_dict()
        if not complete: return False, None, None, None, None, None, None, None

        tk.Label(popup).pack()

        def class_selection_call_back(event):
            class_name = event.widget.get()
            student_combo.set("학생 선택")
            student_combo["values"] = list(class_student_dict[class_name].keys())
            test_list_combo.set("시험 선택")
            test_list_combo["values"] = list(class_test_dict[class_name].keys())

        target_class_var = tk.StringVar()
        target_class_combo = ttk.Combobox(popup, values=list(class_student_dict.keys()), state="readonly", textvariable=target_class_var, width=100)
        target_class_combo.set("반 선택")
        target_class_combo.bind("<<ComboboxSelected>>", class_selection_call_back)
        target_class_combo.pack()

        target_student_var = tk.StringVar()
        student_combo = ttk.Combobox(popup, values=None, state="readonly", textvariable=target_student_var, width=100)
        student_combo.set("학생 선택")
        student_combo.pack()

        target_test_name_var = tk.StringVar()
        test_list_combo = ttk.Combobox(popup, values=None, state="readonly", textvariable=target_test_name_var, width=100)
        test_list_combo.set("시험 선택")
        test_list_combo.pack()

        score_var = tk.StringVar()
        tk.Entry(popup, textvariable=score_var, width=28).pack()

        makeup_test_check_var = tk.BooleanVar()
        tk.Checkbutton(popup, text="재시험 미응시", variable=makeup_test_check_var).pack()

        tk.Button(popup, text="메세지 전송 및 저장", width=20 , command=quit_event).pack()
        
        popup.mainloop()
        
        target_class_name   = target_class_var.get()
        target_student_name = target_student_var.get()
        target_test_name    = target_test_name_var.get()
        test_score          = score_var.get()
        makeup_test_check   = makeup_test_check_var.get()

        if target_class_name == "반 선택" or target_student_name == "학생 선택" or target_test_name == "시험 선택" or test_score == "":
            return False, None, None, None, None, None, None, None


        try:
            if '.' in test_score:
                test_score = float(test_score)
            else:
                test_score = int(test_score)
        except:
            OmikronLog.error("올바른 점수를 입력해 주세요.")
            return False, None, None, None, None, None, None, None

        target_row = class_student_dict[target_class_name][target_student_name]
        target_col = class_test_dict[target_class_name][target_test_name]

        target_test_name = target_test_name[9:]

        return True, target_student_name, target_class_name, target_test_name, target_row, target_col, test_score, makeup_test_check

    def save_makeup_test_result_dialog(self):
        """
        재시험 결과 작성 팝업
        
        return : 성공 여부, 작성 행, 재시험 점수
        """
        def quit_event():
            popup.quit()
            popup.destroy()

        popup = tk.Toplevel(self.ui)
        width = 250
        height = 150
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("재시험 기록")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quit_event)

        complete, class_student_dict, _ = omikron.datafile.get_data_sorted_dict()
        if not complete: return False, None, None

        complete, student_test_dict = omikron.makeuptest.get_studnet_test_index_dict()
        if not complete: return False, None, None

        def class_selection_call_back(event):
            class_name = event.widget.get()
            student_combo.set("학생 선택")
            student_combo["values"] = list(class_student_dict[class_name])
            makeup_test_list_combo.set("재시험 선택")
            makeup_test_list_combo["values"] = []

        target_class_var = tk.StringVar()
        target_class_combo = ttk.Combobox(popup, values=list(class_student_dict.keys()), state="readonly", textvariable=target_class_var, width=100)
        target_class_combo.set("반 선택")
        target_class_combo.bind("<<ComboboxSelected>>", class_selection_call_back)
        target_class_combo.pack()

        def student_selection_call_back(event):
            student_name = event.widget.get()
            makeup_test_list_combo.set("재시험 선택")
            try:
                makeup_test_list_combo["values"] = list(student_test_dict[student_name].keys())
            except:
                makeup_test_list_combo.set("재시험이 없습니다")
                makeup_test_list_combo["values"] = []

        target_student_var = tk.StringVar()
        student_combo = ttk.Combobox(popup, values=None, state="readonly", textvariable=target_student_var, width=100)
        student_combo.set("학생 선택")
        student_combo.bind("<<ComboboxSelected>>", student_selection_call_back)
        student_combo.pack()

        makeup_test_name_var = tk.StringVar()
        makeup_test_list_combo = ttk.Combobox(popup, values=None, state="readonly", textvariable=makeup_test_name_var, width=100)
        makeup_test_list_combo.set("재시험 선택")
        makeup_test_list_combo.pack()

        makeup_test_score_var = tk.StringVar()
        tk.Entry(popup, textvariable=makeup_test_score_var, width=28).pack()

        tk.Label(popup).pack()
        tk.Button(popup, text="재시험 저장", width=10 , command=quit_event).pack()
        
        popup.mainloop()

        target_class_name   = target_class_var.get()
        target_student_name = target_student_var.get()
        makeup_test_name    = makeup_test_name_var.get()
        makeup_test_score   = makeup_test_score_var.get()
        
        if target_class_name == "반 선택" or target_student_name == "학생 선택" or makeup_test_name == "재시험 선택" or makeup_test_name == "재시험이 없습니다" or makeup_test_score == "":
            return False, None, None
        
        target_row = student_test_dict[target_student_name][makeup_test_name]

        return True, target_row, makeup_test_score

    # tasks
    def make_class_info_file_task(self):
        thread = threading.Thread(target=omikron.thread.make_class_info_file_thread, daemon=True)
        thread.start()

    def data_file_task(self):
        if self.make_data_file_button["text"] == "데이터 파일 생성":
            thread = threading.Thread(target=omikron.thread.make_data_file_thread, daemon=True)
            thread.start()
        else:
            new_filename = self.change_data_file_name_dialog()
            if new_filename is None: return

            if askokcancel("데이터 파일 이름 변경", f"데이터 파일 이름을 {new_filename}으로 변경하시겠습니까?"):
                omikron.config.change_data_file_name(new_filename)

    def student_info_file_task(self):
        if self.make_student_info_file_button["text"] == "학생 정보 기록 양식 생성":
            thread = threading.Thread(target=omikron.thread.make_student_info_file_thread, daemon=True)
            thread.start()
        else:
            thread = threading.Thread(target=omikron.thread.update_student_info_file_thread, daemon=True)
            thread.start()

    def update_class_task(self):
        if omikron.datafile.isopen():
            OmikronLog.log(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        
        if not omikron.datafile.file_validation():
            return

        if self.update_class_button["text"] == "반 업데이트":
            if os.path.isfile(f"./{ClassInfo.TEMP_FILE_NAME}.xlsx"):
                omikron.classinfo.delete_temp()

            try:
                excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
                excel.Visible = True
            except:
                OmikronLog.error("모든 엑셀 프로그램을 종료한 뒤 다시 시도해 주세요.")

            if not omikron.classinfo.make_temp_file_for_update():
                return

            wb = excel.Workbooks.Open(f"{os.getcwd()}\\{ClassInfo.TEMP_FILE_NAME}.xlsx")

            if not askokcancel("반 정보 변경 확인", "반 정보 파일의 빈칸을 채운 뒤 Excel을 종료하고\n버튼을 눌러주세요.\n삭제할 반은 행을 삭제해 주세요.\n취소 선택 시 반 업데이트가 중단됩니다."):
                wb.Close()
                excel.Quit()
                if os.path.isfile(f"./{ClassInfo.TEMP_FILE_NAME}.xlsx"):
                    omikron.classinfo.delete_temp()
                    OmikronLog.log(r"반 업데이트를 중단합니다.")
        else:
            if os.path.isfile(f"./~${ClassInfo.TEMP_FILE_NAME}.xlsx"):
                OmikronLog.log(r"임시 파일을 닫은 뒤 다시 시도해 주세요.")
                return

            if not askokcancel("반 정보 변경 확인", "반 업데이트를 계속하시겠습니까?"):
                if os.path.isfile(f"./{ClassInfo.TEMP_FILE_NAME}.xlsx"):
                    omikron.classinfo.delete_temp()
                    OmikronLog.log(r"반 업데이트를 중단합니다.")

            thread = threading.Thread(target=omikron.thread.update_class_thread, daemon=True)
            thread.start()

    def make_data_form_task(self):
        thread = threading.Thread(target=omikron.thread.make_data_form_thread, daemon=True)
        thread.start()

    def save_test_data_task(self):
        if omikron.datafile.isopen():
            OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        if omikron.makeuptest.isopen():
            OmikronLog.error(r"재시험 명단 파일을 닫은 뒤 다시 시도해 주세요.")
            return

        if self.makeup_test_date is None:
            self.holiday_dialog()

        if not omikron.datafile.file_validation():
            return

        filepath = filedialog.askopenfilename(initialdir="./", title="데일리테스트 기록 양식 선택", filetypes=(("Excel files", "*.xlsx"),("all files", "*.*")))
        if filepath == "": return

        if not omikron.dataform.data_validation(filepath):
            return

        thread = threading.Thread(target=lambda: omikron.thread.save_test_data_thread(filepath, self.makeup_test_date), daemon=True)
        thread.start()

    def send_message_task(self):
        if self.makeup_test_date is None:
            self.holiday_dialog()

        filepath = filedialog.askopenfilename(initialdir="./", title="데일리테스트 기록 양식 선택", filetypes=(("Excel files", "*.xlsx"),("all files", "*.*")))
        if filepath == "": return

        if not omikron.dataform.data_validation(filepath):
            return

        thread = threading.Thread(target=lambda: omikron.thread.send_message_thread(filepath, self.makeup_test_date), daemon=True)
        thread.start()

    def save_individual_test_task(self):
        if omikron.datafile.isopen():
            OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        if omikron.makeuptest.isopen():
            OmikronLog.error(r"재시험 명단 파일을 닫은 뒤 다시 시도해 주세요.")
            return

        if self.makeup_test_date is None:
            self.holiday_dialog()

        if not omikron.datafile.file_validation():
            return

        complete, target_student_name, target_class_name, target_test_name, target_row, target_col, test_score, makeup_test_check = self.save_individual_test_dialog()
        if not complete: return

        is_empty, value = omikron.datafile.is_cell_empty(target_row, target_col)
        if not is_empty:
            if not askyesno("데이터 중복 확인", f"{target_student_name} 학생의 {target_test_name} 시험에 대한 점수({value}점)가 이미 존재합니다.\n덮어쓰시겠습니까?"):
                OmikronLog.log(r"개별 데이터 저장을 취소하였습니다.")
                return

        thread = threading.Thread(target=lambda: omikron.thread.save_individual_test_thread(target_student_name, target_class_name, target_test_name, target_row, target_col, test_score, makeup_test_check, self.makeup_test_date), daemon=True)
        thread.start()

    def save_makeup_test_result_task(self):
        if omikron.makeuptest.isopen():
            OmikronLog.error(r"재시험 명단 파일을 닫은 뒤 다시 시도해 주세요.")
            return

        if not omikron.datafile.file_validation():
            return

        complete, target_row, makeup_test_score = self.save_makeup_test_result_dialog()
        if not complete: return

        thread = threading.Thread(target=lambda: omikron.thread.save_makeup_test_result_thread(target_row, makeup_test_score), daemon=True)
        thread.start()

    def conditional_formatting_task(self):
        if omikron.datafile.isopen():
            OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
            return

        if not omikron.datafile.file_validation():
            return

        thread = threading.Thread(target=omikron.thread.conditional_formatting_thread, daemon=True)
        thread.start()

    def add_student_task(self):
        if omikron.datafile.isopen():
            OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        if omikron.studentinfo.isopen():
            OmikronLog.error(r"학생 정보 파일을 닫은 뒤 다시 시도해 주세요.")
            return

        if not omikron.datafile.file_validation():
            return

        complete, target_student_name, target_class_name = self.add_student_dialog()
        if not complete: return

        # 학생 추가 확인
        if not askyesno("학생 추가 확인", f"{target_student_name} 학생을 {target_class_name} 반에 추가하시겠습니까?"):
            return

        thread = threading.Thread(target=lambda: omikron.thread.add_student_thread(target_student_name, target_class_name), daemon=True)
        thread.start()

    def delete_student_task(self):
        if omikron.datafile.isopen():
            OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        if omikron.studentinfo.isopen():
            OmikronLog.error(r"학생 정보 파일을 닫은 뒤 다시 시도해 주세요.")
            return

        if not omikron.datafile.file_validation():
            return

        complete, target_student_name = self.delete_student_dialog()
        if not complete: return

        # 퇴원 처리 확인
        if not askyesno("퇴원 확인", f"{target_student_name} 학생을 퇴원 처리하시겠습니까?"):
            return

        thread = threading.Thread(target=lambda: omikron.thread.delete_student_thread(target_student_name), daemon=True)
        thread.start()

    def move_student_task(self):
        if omikron.datafile.isopen():
            OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        if omikron.studentinfo.isopen():
            OmikronLog.error(r"학생 정보 파일을 닫은 뒤 다시 시도해 주세요.")
            return

        if not omikron.datafile.file_validation():
            return

        complete, target_student_name, target_class_name, current_class_name = self.move_student_dialog()
        if not complete: return

        # 학생 반 이동 확인
        if target_class_name == current_class_name:
            OmikronLog.error(r"학생의 현재 반과 이동할 반이 같아 취소되었습니다.")
            return

        if not askyesno("학생 반 이동 확인", f"{current_class_name} 반의 {target_student_name} 학생을\n{target_class_name} 반으로 이동시키겠습니까?"):
            return

        thread = threading.Thread(target=lambda: omikron.thread.move_student_thread(target_student_name, target_class_name, current_class_name), daemon=True)
        thread.start()
