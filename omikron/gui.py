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
        self.width  = 320
        self.height = 630 # button +25

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
        tk.Label(self.ui, text="< 데이터 저장 위치 >").pack()

        self.class_info_file_button = tk.Button(self.ui, cursor="hand2", text="데이터 저장 폴더 열기", width=40, command=self.open_data_dir_task)
        self.class_info_file_button.pack()

        self.data_file_button = tk.Button(self.ui, cursor="hand2", text="데이터 저장 위치 변경", width=40, command=self.change_data_dir_task)
        self.data_file_button.pack()

        tk.Label(self.ui, text="< 기수 변경 관련 >").pack()

        self.class_info_file_button = tk.Button(self.ui, cursor="hand2", text="반 정보 기록 양식 생성", width=40, command=self.class_info_file_task)
        self.class_info_file_button.pack()

        self.data_file_button = tk.Button(self.ui, cursor="hand2", text="데이터 파일 생성", width=40, command=self.data_file_task)
        self.data_file_button.pack()

        self.student_info_file_button = tk.Button(self.ui, cursor="hand2", text="학생 정보 기록 양식 생성", width=40, command=self.student_info_file_task)
        self.student_info_file_button.pack()

        self.change_class_info_button = tk.Button(self.ui, cursor="hand2", text="선생님 변경", width=40, command=self.change_class_info_task)
        self.change_class_info_button.pack()

        tk.Label(self.ui, text="< 데이터 저장 및 메시지 전송 >").pack()

        self.make_data_form_button = tk.Button(self.ui, cursor="hand2", text="데일리 테스트 기록 양식 생성", width=40, command=self.make_data_form_task)
        self.make_data_form_button.pack()

        self.save_test_result_button = tk.Button(self.ui, cursor="hand2", text="시험 결과 저장", width=40, command=self.save_test_result_task)
        self.save_test_result_button.pack()

        self.send_message_button = tk.Button(self.ui, cursor="hand2", text="시험 결과 메시지 전송", width=40, command=self.send_message_task)
        self.send_message_button.pack()

        self.save_individual_test_button = tk.Button(self.ui, cursor="hand2", text="개별 시험 결과 저장", width=40, command=self.save_individual_test_task)
        self.save_individual_test_button.pack()

        self.save_makeup_test_result_button = tk.Button(self.ui, cursor="hand2", text="재시험 결과 저장", width=40, command=self.save_makeup_test_result_task)
        self.save_makeup_test_result_button.pack()

        tk.Label(self.ui, text="< 데이터 관리 >").pack()

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

        if os.path.isfile(f"{omikron.config.DATA_DIR}/{ClassInfo.DEFAULT_NAME}.xlsx"):
            if os.path.isfile(f"{omikron.config.DATA_DIR}/{ClassInfo.TEMP_FILE_NAME}.xlsx"):
                self.class_info_file_button["text"] = "반 정보 수정 후 반 업데이트 계속하기"
            else:
                self.class_info_file_button["text"] = "반 업데이트"
            self.class_info_file_button["state"] = tk.DISABLED
            self.data_file_button["state"]       = tk.NORMAL
            classinfo_check = True
        else:
            self.class_info_file_button["text"]  = "반 정보 기록 양식 생성"
            self.class_info_file_button["state"] = tk.NORMAL
            self.data_file_button["state"]       = tk.DISABLED

        if os.path.isfile(f"{omikron.config.DATA_DIR}/{StudentInfo.DEFAULT_NAME}.xlsx"):
            self.student_info_file_button["text"] = "학생 정보 업데이트"
            studentinfo_check = True
        else: 
            self.student_info_file_button["text"] = "학생 정보 기록 양식 생성"

        if os.path.isfile(f"{omikron.config.DATA_DIR}/data/{omikron.config.DATA_FILE_NAME}.xlsx"):
            self.data_file_button["text"]        = "데이터 파일 이름 변경"
            self.class_info_file_button["state"] = tk.NORMAL
            datafile_check = True
        else:
            self.data_file_button["text"] = "데이터 파일 생성"

        if classinfo_check and studentinfo_check and datafile_check:
            self.make_data_form_button["state"]          = tk.NORMAL
            self.save_test_result_button["state"]        = tk.NORMAL
            self.send_message_button["state"]            = tk.NORMAL
            self.save_individual_test_button["state"]    = tk.NORMAL
            self.apply_color_button["state"]             = tk.NORMAL
            self.add_student_button["state"]             = tk.NORMAL
            self.delete_student_button["state"]          = tk.NORMAL
            self.move_student_button["state"]            = tk.NORMAL
            if os.path.isfile(f"{omikron.config.DATA_DIR}/data/{MakeupTestList.DEFAULT_NAME}.xlsx"):
                self.save_makeup_test_result_button["state"] = tk.NORMAL
            else:
                self.save_makeup_test_result_button["state"] = tk.DISABLED
        else:
            self.make_data_form_button["state"]          = tk.DISABLED
            self.save_test_result_button["state"]        = tk.DISABLED
            self.send_message_button["state"]            = tk.DISABLED
            self.save_individual_test_button["state"]    = tk.DISABLED
            self.save_makeup_test_result_button["state"] = tk.DISABLED
            self.apply_color_button["state"]             = tk.DISABLED
            self.add_student_button["state"]             = tk.DISABLED
            self.delete_student_button["state"]          = tk.DISABLED
            self.move_student_button["state"]            = tk.DISABLED

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

        width  = 250
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

        width  = 200
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

        return `성공 여부`, `학생 이름`
        """
        class_student_dict, _ = omikron.datafile.get_data_sorted_dict()

        def quit_event():
            popup.quit()
            popup.destroy()

        popup = tk.Toplevel(self.ui)

        width  = 250
        height = 140
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("퇴원 관리")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quit_event)

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

        return `성공 여부`, `학생 이름`, `목표 반`, `현재 반`
        """
        class_student_dict, _ = omikron.datafile.get_data_sorted_dict()

        def quit_event():
            popup.quit()
            popup.destroy()

        popup = tk.Toplevel(self.ui)

        width  = 250
        height = 160
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("학생 반 이동")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quit_event)

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

        return `성공 여부`, `학생 이름`, `목표 반`
        """
        class_wb = omikron.classinfo.open()
        class_ws = omikron.classinfo.open_worksheet(class_wb)

        class_names = omikron.classinfo.get_class_names(class_ws)

        def quit_event():
            popup.quit()
            popup.destroy()

        popup = tk.Toplevel(self.ui)

        width  = 250
        height = 140
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("신규생 추가")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quit_event)

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

        return `성공 여부`, `학생 이름`, `목표 반`, `시험 이름`, `작성 행`, `작성 열`, `시험 점수`, `재시험 미응시 여부`
        """
        class_student_dict, class_test_dict = omikron.datafile.get_data_sorted_dict()

        def quit_event():
            popup.quit()
            popup.destroy()

        popup = tk.Toplevel(self.ui)

        width  = 250
        height = 170
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("개별 점수 기록")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quit_event)

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

        target_test_name = target_test_name[11:]

        return True, target_student_name, target_class_name, target_test_name, target_row, target_col, test_score, makeup_test_check

    def save_makeup_test_result_dialog(self):
        """
        재시험 결과 작성 팝업
        
        return : `성공 여부`, `작성 행`, `재시험 점수`
        """
        class_student_dict, _ = omikron.datafile.get_data_sorted_dict()

        student_test_dict = omikron.makeuptest.get_studnet_test_index_dict()

        def quit_event():
            popup.quit()
            popup.destroy()

        popup = tk.Toplevel(self.ui)

        width  = 250
        height = 150
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("재시험 결과 저장")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quit_event)

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

    def update_class_dialog(self):
        COLOR_MAP = {
            "green":  "#108a00",   # chrome 전용 (왼↔중만)
            "orange": "#cc7a00",   # classinfo 전용 (중↔오른만)
            "black":  "#000000",   # 공통 (중↔오른만)
        }

        # 데이터 분류
        chrome_set    = set(omikron.chrome.get_class_names())
        classinfo_set = set(omikron.classinfo.get_class_names())

        only_chrome   = sorted(chrome_set - classinfo_set)   # 초록
        only_class    = sorted(classinfo_set - chrome_set)   # 주황
        both          = sorted(classinfo_set & chrome_set)   # 검정

        # 전역 컬러 테이블 (항목 → 색)
        ITEM_COLOR = {}
        for v in only_chrome:
            ITEM_COLOR[v] = "green"
        for v in only_class:
            ITEM_COLOR[v] = "orange"
        for v in both:
            ITEM_COLOR[v] = "black"

        def insert_with_color(lb: tk.Listbox, value: str):
            """목록에 값 추가 후 지정 색상 적용(중복 방지)."""
            existing = set(lb.get(0, tk.END))
            if value in existing:
                return
            lb.insert(tk.END, value)
            idx = lb.size() - 1
            lb.itemconfig(idx, foreground=COLOR_MAP[ITEM_COLOR[value]])

        def get_color(value: str) -> str:
            return ITEM_COLOR.get(value, "black")

        def can_move(value: str, src_name: str, dst_name: str) -> bool:
            """
            이동 제약:
            - green(only chrome): left ↔ mid 만 허용
            - orange/black(only class or both): mid ↔ right 만 허용
            """
            c = get_color(value)
            if c == "green":
                return {src_name, dst_name} <= {"left", "mid"}
            else:  # orange, black
                return {src_name, dst_name} <= {"mid", "right"}

        def move_selected(src: tk.Listbox, dst: tk.Listbox, src_name: str, dst_name: str):
            sel = list(src.curselection())
            if not sel:
                return
            # 선택 항목 → 값 목록
            vals = [src.get(i) for i in sel]

            # 대상에 없는 것만, 허용 이동만 추가
            dst_existing = set(dst.get(0, tk.END))
            moved_idxs = []
            for i, v in zip(sel, vals):
                if not can_move(v, src_name, dst_name):
                    continue
                if v not in dst_existing:
                    insert_with_color(dst, v)
                    moved_idxs.append(i)

            # 삭제는 인덱스 역순
            for i in reversed(moved_idxs):
                src.delete(i)

        def select_all(lb: tk.Listbox, *_):
            lb.select_set(0, tk.END)

        popup = tk.Toplevel(self.ui)
        popup.title("반 리스트 수정")
        width, height = 900, 520
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")

        container = ttk.Frame(popup, padding=10)
        container.grid(row=0, column=0, sticky="nsew")
        popup.rowconfigure(0, weight=1)
        popup.columnconfigure(0, weight=1)

        for c in (0, 2, 4):
            container.columnconfigure(c, weight=1, uniform="cols")
        for c in (1, 3):
            container.columnconfigure(c, minsize=64)
        container.rowconfigure(0, weight=1)
        container.rowconfigure(1, weight=0)
        container.rowconfigure(2, weight=0)

        def build_labeled_list(parent, title: str) -> tk.Listbox:
            lf = ttk.LabelFrame(parent, text=title, padding=(10, 8))
            lf.grid(sticky="nsew")
            lf.rowconfigure(0, weight=1)
            lf.columnconfigure(0, weight=1)

            wrap = ttk.Frame(lf)
            wrap.grid(row=0, column=0, sticky="nsew")
            wrap.rowconfigure(0, weight=1)
            wrap.columnconfigure(0, weight=1)

            sb = ttk.Scrollbar(wrap, orient="vertical")
            lb = tk.Listbox(
                wrap,
                selectmode=tk.EXTENDED,
                activestyle="dotbox",
                yscrollcommand=sb.set,
                exportselection=False,
            )
            sb.config(command=lb.yview)
            lb.grid(row=0, column=0, sticky="nsew")
            sb.grid(row=0, column=1, sticky="ns")
            return lb

        left_cell  = ttk.Frame(container);  left_cell.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        mid_cell   = ttk.Frame(container);  mid_cell.grid(row=0, column=2, sticky="nsew", padx=10)
        right_cell = ttk.Frame(container); right_cell.grid(row=0, column=4, sticky="nsew", padx=(10, 0))

        for cell in (left_cell, mid_cell, right_cell):
            cell.rowconfigure(0, weight=1)
            cell.columnconfigure(0, weight=1)

        lb_left  = build_labeled_list(left_cell,  "추가되지 않은 반")
        lb_mid   = build_labeled_list(mid_cell,   "업데이트(유지) 할 반")
        lb_right = build_labeled_list(right_cell, "삭제할 반")

        def build_arrow_column(parent, to_left, to_right):
            parent.rowconfigure(0, weight=1)
            parent.columnconfigure(0, weight=1)

            col = ttk.Frame(parent); col.grid(row=0, column=0, sticky="nsew")
            inner = ttk.Frame(col); inner.place(relx=0.5, rely=0.5, anchor="center")

            btn_left  = ttk.Button(inner, text="←", width=3, command=to_left)
            btn_right = ttk.Button(inner, text="→", width=3, command=to_right)
            btn_left.pack(pady=6)
            btn_right.pack(pady=6)
            return col

        # 왼쪽 ↔ 가운데
        lm_col = ttk.Frame(container); lm_col.grid(row=0, column=1, sticky="nsew")
        build_arrow_column(
            lm_col,
            to_left = lambda: move_selected(lb_mid,  lb_left,  "mid",  "left"),
            to_right= lambda: move_selected(lb_left, lb_mid,   "left", "mid"),
        )

        # 가운데 ↔ 오른쪽
        mr_col = ttk.Frame(container); mr_col.grid(row=0, column=3, sticky="nsew")
        build_arrow_column(
            mr_col,
            to_left = lambda: move_selected(lb_right, lb_mid,  "right", "mid"),
            to_right= lambda: move_selected(lb_mid,   lb_right,"mid",   "right"),
        )

        ttk.Separator(container, orient="horizontal").grid(
            row=1, column=0, columnspan=5, sticky="ew", pady=(12, 8)
        )

        btnbar = ttk.Frame(container); btnbar.grid(row=2, column=0, columnspan=5, sticky="e")

        result = {"ok": False, "mid": None}

        def on_ok(*_):
            result["ok"] = True
            result["mid"] = list(lb_mid.get(0, tk.END))
            popup.destroy()

        def on_cancel(*_):
            popup.destroy()

        ttk.Button(btnbar, text="취소", command=on_cancel).pack(side="right", padx=(6, 0))
        ttk.Button(btnbar, text="확인", command=on_ok).pack(side="right")

        # 단축키: Ctrl+A (각 리스트별 전체 선택)
        for lb in (lb_left, lb_mid, lb_right):
            lb.bind("<Control-a>", lambda e, _lb=lb: select_all(_lb))

        # ── 초기 채우기
        # 왼쪽: chrome에만 있는 반(초록)
        for item in only_chrome:
            insert_with_color(lb_left, item)
        # 가운데: classinfo에 있는 반 (겹치는 애는 검정, classinfo만 있는 애는 주황)
        for item in both:
            insert_with_color(lb_mid, item)
        for item in only_class:
            insert_with_color(lb_mid, item)
        # 오른쪽: 초기 비움

        self.ui.wait_window(popup)

        if result["ok"]:
            # 중앙 리스트만 리턴
            return True, result["mid"]
        else:
            return False, None

    def change_class_info_dialog(self):
        class_wb = omikron.classinfo.open()
        class_ws = omikron.classinfo.open_worksheet(class_wb)

        class_names = omikron.classinfo.get_class_names(class_ws)

        def quit_event():
            popup.quit()
            popup.destroy()

        popup = tk.Toplevel(self.ui)

        width  = 250
        height = 200
        x = int((popup.winfo_screenwidth()/4) - (width/2))
        y = int((popup.winfo_screenheight()/2) - (height/2))
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.title("선생님 변경")
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", quit_event)

        tk.Label(popup).pack()

        def class_selection_call_back(event):
            selected_class_name = event.widget.get()
            _, teacher_name, _, _ = omikron.classinfo.get_class_info(selected_class_name, ws=class_ws)
            teacher_name_label.config(text=f"현재 선생님: {teacher_name}")
        target_class_var = tk.StringVar()
        class_combo = ttk.Combobox(popup, values=class_names, state="readonly", textvariable=target_class_var, width=25)
        class_combo.set("수정할 반 선택")
        class_combo.bind("<<ComboboxSelected>>", class_selection_call_back)
        class_combo.pack()

        tk.Label(popup).pack()
        teacher_name_label = tk.Label(popup, text="현재 선생님: ")
        teacher_name_label.pack()

        tk.Label(popup).pack()
        target_teacher_var = tk.StringVar()
        tk.Entry(popup, textvariable=target_teacher_var, width=28).pack()

        tk.Label(popup).pack()
        tk.Button(popup, text="선생님 변경", width=10 , command=quit_event).pack()

        popup.mainloop()

        target_class_name   = target_class_var.get()
        target_teacher_name = target_teacher_var.get()
        if target_class_name == "수정할 반 선택" or target_teacher_name == "":
            return False, None, None
        else:
            return True, target_class_name, target_teacher_name

    # tasks
    def open_data_dir_task(self):
        os.startfile(omikron.config.DATA_DIR)
        return

    def change_data_dir_task(self):
        dir_path = filedialog.askdirectory(initialdir=f"{omikron.config.DATA_DIR}/", title="변경할 데이터 저장 위치 선택")
        if dir_path == "": return

        omikron.config.change_data_path(dir_path)

        OmikronLog.log("데이터 저장 위치를 변경하였습니다.")
        return

    def class_info_file_task(self):
        if self.class_info_file_button["text"] == "반 정보 기록 양식 생성":
            thread = threading.Thread(target=omikron.thread.make_class_info_file_thread, daemon=True)
            thread.start()
        elif self.class_info_file_button["text"] == "반 업데이트":
            if omikron.datafile.isopen():
                OmikronLog.log(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
                return

            OmikronLog.log(r"반 업데이트를 시작합니다.")
            OmikronLog.log(r"반 정보를 불러오는 중...")

            completed, new_class_list = self.update_class_dialog()
            if not completed:
                OmikronLog.log(r"반 업데이트를 중단하였습니다.")
                return

            OmikronLog.log(f"{ClassInfo.TEMP_FILE_NAME}.xlsx 생성 중...")
            temp_path =  omikron.classinfo.make_temp_file_for_update(new_class_list)

            try:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = True
                wb = excel.Workbooks.Open(temp_path)
            except:
                OmikronLog.error(r"모든 엑셀 프로그램을 종료한 뒤 다시 시도해 주세요.")

            OmikronLog.log(f"'{ClassInfo.TEMP_FILE_NAME}.xlsx' 수정된 반의 정보를 입력해 주세요")
            if not askokcancel("반 정보 변경 확인", f"{ClassInfo.TEMP_FILE_NAME}.xlsx 파일의\n각 반의 상세 정보를 수정한 뒤 저장하고\n반 업데이트 계속하기 버튼을 눌러주세요.\n\n취소 선택 시 반 업데이트가 중단됩니다."):
                wb.Close()
                excel.Quit()
                if os.path.isfile(f"{omikron.config.DATA_DIR}/{ClassInfo.TEMP_FILE_NAME}.xlsx"):
                    omikron.classinfo.delete_temp()
                    OmikronLog.log(r"반 업데이트를 중단하였습니다.")
        else:
            if omikron.datafile.isopen():
                OmikronLog.log(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
                return

            if os.path.isfile(f"{omikron.config.DATA_DIR}/~${ClassInfo.TEMP_FILE_NAME}.xlsx"):
                OmikronLog.log(f"'{ClassInfo.TEMP_FILE_NAME}' 파일을 닫은 뒤 다시 시도해 주세요.")
                return

            if not askokcancel("반 정보 변경 확인", "반 업데이트를 계속하시겠습니까?"):
                if os.path.isfile(f"{omikron.config.DATA_DIR}/{ClassInfo.TEMP_FILE_NAME}.xlsx"):
                    omikron.classinfo.delete_temp()
                    OmikronLog.log(r"반 업데이트를 중단하였습니다.")

            thread = threading.Thread(target=omikron.thread.update_class_thread, daemon=True)
            thread.start()

    def data_file_task(self):
        if self.data_file_button["text"] == "데이터 파일 생성":
            thread = threading.Thread(target=omikron.thread.make_data_file_thread, daemon=True)
            thread.start()
        else:
            if omikron.datafile.isopen():
                OmikronLog.log(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
                return

            new_filename = self.change_data_file_name_dialog()
            if new_filename == "": return

            if askokcancel("데이터 파일 이름 변경", f"데이터 파일 이름을 '{new_filename}'(으)로 변경하시겠습니까?"):
                omikron.config.change_data_file_name(new_filename)
                OmikronLog.log(f"데이터 파일 이름을 '{new_filename}'(으)로 변경하였습니다.")

    def student_info_file_task(self):
        if self.student_info_file_button["text"] == "학생 정보 기록 양식 생성":
            thread = threading.Thread(target=omikron.thread.make_student_info_file_thread, daemon=True)
            thread.start()
        else:
            if omikron.studentinfo.isopen():
                OmikronLog.log(f"'{StudentInfo.DEFAULT_NAME}' 파일을 닫은 뒤 다시 시도해 주세요.")
                return

            thread = threading.Thread(target=omikron.thread.update_student_info_file_thread, daemon=True)
            thread.start()

    def change_class_info_task(self):
        if omikron.classinfo.isopen():
            OmikronLog.log(f"'{ClassInfo.DEFAULT_NAME}' 파일을 닫은 뒤 다시 시도해 주세요.")
            return

        complete, target_class_name, target_teacher_name = self.change_class_info_dialog()
        if not complete: return

        if not askyesno("선생님 변경 확인", f"{target_class_name} 반의 선생님을 {target_teacher_name}으로 변경하시겠습니까?"):
            return

        thread = threading.Thread(target=lambda: omikron.thread.change_class_info_thread(target_class_name, target_teacher_name), daemon=True)
        thread.start()

    def make_data_form_task(self):
        thread = threading.Thread(target=omikron.thread.make_data_form_thread, daemon=True)
        thread.start()

    def save_test_result_task(self):
        if omikron.datafile.isopen():
            OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        if omikron.makeuptest.isopen():
            OmikronLog.error(f"'{MakeupTestList.DEFAULT_NAME}' 파일을 닫은 뒤 다시 시도해 주세요.")
            return

        if self.makeup_test_date is None:
            self.holiday_dialog()

        filepath = filedialog.askopenfilename(initialdir=f"{omikron.config.DATA_DIR}/", title="데일리테스트 기록 양식 선택", filetypes=(("Excel files", "*.xlsx"),("all files", "*.*")))
        if filepath == "": return

        if not omikron.dataform.data_validation(filepath):
            return

        print("end task")
        thread = threading.Thread(target=lambda: omikron.thread.save_test_result_thread(filepath, self.makeup_test_date), daemon=True)
        thread.start()

    def send_message_task(self):
        if self.makeup_test_date is None:
            self.holiday_dialog()

        filepath = filedialog.askopenfilename(initialdir=f"{omikron.config.DATA_DIR}/", title="데일리테스트 기록 양식 선택", filetypes=(("Excel files", "*.xlsx"),("all files", "*.*")))
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
            OmikronLog.error(f"'{MakeupTestList.DEFAULT_NAME}' 파일을 닫은 뒤 다시 시도해 주세요.")
            return

        if self.makeup_test_date is None:
            self.holiday_dialog()

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
            OmikronLog.error(f"'{MakeupTestList.DEFAULT_NAME}' 파일을 닫은 뒤 다시 시도해 주세요.")
            return

        complete, target_row, makeup_test_score = self.save_makeup_test_result_dialog()
        if not complete: return

        thread = threading.Thread(target=lambda: omikron.thread.save_makeup_test_result_thread(target_row, makeup_test_score), daemon=True)
        thread.start()

    def conditional_formatting_task(self):
        if omikron.datafile.isopen():
            OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
            return

        thread = threading.Thread(target=omikron.thread.conditional_formatting_thread, daemon=True)
        thread.start()

    def add_student_task(self):
        if omikron.datafile.isopen():
            OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
            return
        if omikron.studentinfo.isopen():
            OmikronLog.error(f"'{StudentInfo.DEFAULT_NAME}' 파일을 닫은 뒤 다시 시도해 주세요.")
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
            OmikronLog.error(f"'{StudentInfo.DEFAULT_NAME}' 파일을 닫은 뒤 다시 시도해 주세요.")
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
            OmikronLog.error(f"'{StudentInfo.DEFAULT_NAME}' 파일을 닫은 뒤 다시 시도해 주세요.")
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
