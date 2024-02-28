import os
import openpyxl as xl
import pythoncom       # only works in Windows
import win32com.client # only works in Windows

from datetime import datetime
from openpyxl.utils.cell import get_column_letter as gcl
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side

import omikron.chrome
import omikron.classinfo
import omikron.config
import omikron.datafile
import omikron.dataform
import omikron.makeuptest
import omikron.studentinfo

from omikron.defs import DataFile, DataForm
from omikron.log import OmikronLog
from omikron.util import copy_cell, class_average_color, student_average_color, test_score_color

# 파일 기본 작업
def make_file() -> bool:
    wb = xl.Workbook()
    ws = wb.worksheets[0]
    ws.title = DataFile.FIRST_SHEET_NAME
    ws[gcl(DataFile.TEST_TIME_COLUMN)+"1"]     = "시간"
    ws[gcl(DataFile.CLASS_WEEKDAY_COLUMN)+"1"] = "요일"
    ws[gcl(DataFile.CLASS_NAME_COLUMN)+"1"]    = "반"
    ws[gcl(DataFile.TEACHER_NAME_COLUMN)+"1"]  = "담당"
    ws[gcl(DataFile.STUDENT_NAME_COLUMN)+"1"]  = "이름"
    ws[gcl(DataFile.AVERAGE_SCORE_COLUMN)+"1"] = "학생 평균"
    ws.freeze_panes    = f"{gcl(DataFile.DATA_COLUMN)}2"
    ws.auto_filter.ref = f"A:{gcl(DataFile.MAX)}"

    for col in range(1, DataFile.DATA_COLUMN):
        ws.cell(1, col).border = Border(bottom = Side(border_style="medium", color="000000"))

    class_wb = omikron.classinfo.open(True)
    complete, class_ws = omikron.classinfo.open_worksheet(class_wb)
    if not complete: return False

    # 반 루프
    for class_name, student_list in omikron.chrome.get_class_student_dict().items():
        if len(student_list) == 0:
            continue

        exist, teacher_name, class_weekday, test_time = omikron.classinfo.get_class_info(class_ws, class_name)
        if not exist: continue

        WRITE_LOCATION = ws.max_row + 1

        # 시험명
        ws.cell(WRITE_LOCATION, DataFile.TEST_TIME_COLUMN).value     = test_time
        ws.cell(WRITE_LOCATION, DataFile.CLASS_WEEKDAY_COLUMN).value = class_weekday
        ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value    = class_name
        ws.cell(WRITE_LOCATION, DataFile.TEACHER_NAME_COLUMN).value  = teacher_name
        ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value  = "날짜"
        
        WRITE_LOCATION = ws.max_row + 1
        ws.cell(WRITE_LOCATION, DataFile.TEST_TIME_COLUMN).value     = test_time
        ws.cell(WRITE_LOCATION, DataFile.CLASS_WEEKDAY_COLUMN).value = class_weekday
        ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value    = class_name
        ws.cell(WRITE_LOCATION, DataFile.TEACHER_NAME_COLUMN).value  = teacher_name
        ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value  = "시험명"

        for col in range(1, DataFile.DATA_COLUMN):
            ws.cell(WRITE_LOCATION, col).border = Border(bottom = Side(border_style="thin", color="909090"))

        class_start = WRITE_LOCATION + 1

        # 학생 루프
        for student_name in student_list:
            WRITE_LOCATION = ws.max_row + 1
            ws.cell(WRITE_LOCATION, DataFile.TEST_TIME_COLUMN).value     = test_time
            ws.cell(WRITE_LOCATION, DataFile.CLASS_WEEKDAY_COLUMN).value = class_weekday
            ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value    = class_name
            ws.cell(WRITE_LOCATION, DataFile.TEACHER_NAME_COLUMN).value  = teacher_name
            ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value  = student_name
            ws.cell(WRITE_LOCATION, DataFile.AVERAGE_SCORE_COLUMN).value = f"=ROUND(AVERAGE(G{WRITE_LOCATION}:XFD{WRITE_LOCATION}), 0)"
            ws.cell(WRITE_LOCATION, DataFile.AVERAGE_SCORE_COLUMN).font  = Font(bold=True)
        
        # 시험별 평균
        class_end = WRITE_LOCATION
        WRITE_LOCATION = ws.max_row + 1
        ws.cell(WRITE_LOCATION, DataFile.TEST_TIME_COLUMN).value     = test_time
        ws.cell(WRITE_LOCATION, DataFile.CLASS_WEEKDAY_COLUMN).value = class_weekday
        ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value    = class_name
        ws.cell(WRITE_LOCATION, DataFile.TEACHER_NAME_COLUMN).value  = teacher_name
        ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value  = "시험 평균"
        ws[f"F{WRITE_LOCATION}"] = ArrayFormula(f"F{WRITE_LOCATION}", f"=ROUND(AVERAGE(IFERROR(F{class_start}:F{class_end}, \"\")), 0)")
        ws.cell(WRITE_LOCATION, DataFile.AVERAGE_SCORE_COLUMN).font = Font(bold=True)

        for col in range(1, DataFile.DATA_COLUMN):
            ws.cell(WRITE_LOCATION, col).border = Border(top = Side(border_style="thin", color="909090"), bottom = Side(border_style="medium", color="000000"))

    # 정렬
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")
    
    # 모의고사 sheet 생성
    copy_ws                 = wb.copy_worksheet(wb[DataFile.FIRST_SHEET_NAME])
    copy_ws.title           = DataFile.SECOND_SHEET_NAME
    copy_ws.freeze_panes    = f"{gcl(DataFile.DATA_COLUMN)}2"
    copy_ws.auto_filter.ref = f"A:{gcl(DataFile.MAX)}"

    save(wb)

    return True

def open(data_only:bool=False) -> xl.Workbook:
    return xl.load_workbook(f"./data/{omikron.config.DATA_FILE_NAME}.xlsx", data_only=data_only)

def open_temp(data_only:bool=False) -> xl.Workbook:
    return xl.load_workbook(f"./data/{DataFile.TEMP_FILE_NAME}.xlsx", data_only=data_only)

def save(wb:xl.Workbook):
    wb.save(f"./data/{omikron.config.DATA_FILE_NAME}.xlsx")

def save_to_temp(wb:xl.Workbook):
    wb.save(f"./data/{DataFile.TEMP_FILE_NAME}.xlsx")
    os.system(f"attrib +h ./data/{DataFile.TEMP_FILE_NAME}.xlsx")

def delete_temp():
    os.remove(f"./data/{DataFile.TEMP_FILE_NAME}.xlsx")

def isopen() -> bool:
    return os.path.isfile(f"./data/~${omikron.config.DATA_FILE_NAME}.xlsx")

# 파일 유틸리티
def make_backup_file():
    wb = open()
    wb.save(f"./data/backup/{omikron.config.DATA_FILE_NAME}({datetime.today().strftime('%Y%m%d')}).xlsx")

def get_data_sorted_dict():
    """
    데이터 파일의 대략적 정보를 `dict` 형태로 추출

    return `성공 여부`, `dict[반:학생]`, `dict[반:시험명]`
    """
    wb = open()

    ws = wb[DataFile.FIRST_SHEET_NAME]

    complete, _, _, CLASS_NAME_COLUMN, _, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(ws)
    if not complete: return False, None, None

    class_wb = omikron.classinfo.open()
    complete, class_ws = omikron.classinfo.open_worksheet(class_wb)
    if not complete: return False, None, None

    class_student_dict = {}
    class_test_dict    = {}

    for class_name in omikron.classinfo.get_class_names(class_ws):
        student_index_dict = {}
        test_index_dict    = {}
        for row in range(2, ws.max_row+1):
            if ws.cell(row, CLASS_NAME_COLUMN).value != class_name:
                continue
            if ws.cell(row, STUDENT_NAME_COLUMN).value == "날짜":
                for col in range(AVERAGE_SCORE_COLUMN+1, ws.max_row+1):
                    test_date = ws.cell(row, col).value
                    test_name = ws.cell(row+1, col).value
                    if test_date is None and test_name is None:
                        break
                    if type(test_date) == datetime:
                        test_date = test_date.strftime("%y.%m.%d")
                    else:
                        test_date = str(test_date).split()[0][2:10].replace("-", ".").replace(",", ".").replace("/", ".")
                    test_index_dict[f"{test_date} {test_name}"] = col
                continue
            if ws.cell(row, STUDENT_NAME_COLUMN).value in ("시험명", "시험 평균"):
                continue
            if ws.cell(row, STUDENT_NAME_COLUMN).font.strike:
                continue
            if ws.cell(row, STUDENT_NAME_COLUMN).font.color is not None and ws.cell(row, STUDENT_NAME_COLUMN).font.color.rgb == "FFFF0000":
                continue
            student_index_dict[ws.cell(row, STUDENT_NAME_COLUMN).value] = row

        test_index_dict = dict(sorted(test_index_dict.items(), reverse=True))

        class_student_dict[class_name] = student_index_dict
        class_test_dict[class_name]    = test_index_dict

    class_student_dict = dict(sorted(class_student_dict.items()))

    return True, class_student_dict, class_test_dict

def find_dynamic_columns(ws:Worksheet):
    """
    파일 열(column) 정보 동적 탐색

    시간, 요일, 반, 담당, 이름, 학생 평균

    return `성공 여부`, `TEST_TIME_COLUMN`, `CLASS_WEEKDAY_COLUMN`, `CLASS_NAME_COLUMN`, `TEACHER_NAME_COLUMN`, `STUDENT_NAME_COLUMN`, `AVERAGE_SCORE_COLUMN`
    """

    for col in range(1, ws.max_column+1):
        if ws.cell(1, col).value == "시간":
            TEST_TIME_COLUMN = col
            break
    else:
        OmikronLog.error(f"{ws.title} 시트에 '시간' 열이 없습니다.")
        return False, None, None, None, None, None, None

    for col in range(1, ws.max_column+1):
        if ws.cell(1, col).value == "요일":
            CLASS_WEEKDAY_COLUMN = col
            break
    else:
        OmikronLog.error(f"{ws.title} 시트에 '요일' 열이 없습니다.")
        return False, None, None, None, None, None, None

    for col in range(1, ws.max_column+1):
        if ws.cell(1, col).value == "반":
            CLASS_NAME_COLUMN = col
            break
    else:
        OmikronLog.error(f"{ws.title} 시트에 '반' 열이 없습니다.")
        return False, None, None, None, None, None, None

    for col in range(1, ws.max_column+1):
        if ws.cell(1, col).value == "담당":
            TEACHER_NAME_COLUMN = col
            break
    else:
        OmikronLog.error(f"{ws.title} 시트에 '담당' 열이 없습니다.")
        return False, None, None, None, None, None, None

    for col in range(1, ws.max_column+1):
        if ws.cell(1, col).value == "이름":
            STUDENT_NAME_COLUMN = col
            break
    else:
        OmikronLog.error(f"{ws.title} 시트에 '이름' 열이 없습니다.")
        return False, None, None, None, None, None, None

    for col in range(1, ws.max_column+1):
        if ws.cell(1, col).value == "학생 평균":
            AVERAGE_SCORE_COLUMN = col
            break
    else:
        OmikronLog.error(f"{ws.title} 시트에 '학생 평균' 열이 없습니다.")
        return False, None, None, None, None, None, None
    
    return True, TEST_TIME_COLUMN, CLASS_WEEKDAY_COLUMN, CLASS_NAME_COLUMN, TEACHER_NAME_COLUMN, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN

def is_cell_empty(row:int, col:int) -> bool:
    """
    데이터 파일이 열려있지 않을 때 특정 셀의 값이 비어있는 지 확인

    데일리테스트 시트 한정 기능
    """
    wb = open(data_only=True)
    ws = wb[DataFile.FIRST_SHEET_NAME]

    if ws.cell(row, col).value is None:
        return True, None

    value = ws.cell(row, col).value

    return False, value

# 파일 작업
def save_test_data(filepath:str):
    """
    데이터 양식에 작성된 데이터를 데이터 파일에 저장
    """
    form_wb = omikron.dataform.open(filepath)
    complete, form_ws = omikron.dataform.open_worksheet(form_wb)
    if not complete: return False, None

    # 학생 정보 열기
    student_wb = omikron.studentinfo.open(True)
    complete, student_ws = omikron.studentinfo.open_worksheet(student_wb)
    if not complete: return False, None

    # 백업 생성
    make_backup_file()

    wb = open()

    for sheet_name in wb.sheetnames:
        if sheet_name == DataFile.FIRST_SHEET_NAME:
            TEST_NAME_COLUMN    = DataForm.DAILYTEST_NAME_COLUMN
            TEST_SCORE_COLUMN   = DataForm.DAILYTEST_SCORE_COLUMN
            TEST_AVERAGE_COLUMN = DataForm.DAILYTEST_AVERAGE_COLUMN
        elif sheet_name == DataFile.SECOND_SHEET_NAME:
            TEST_NAME_COLUMN    = DataForm.MOCKTEST_NAME_COLUMN
            TEST_SCORE_COLUMN   = DataForm.MOCKTEST_SCORE_COLUMN
            TEST_AVERAGE_COLUMN = DataForm.MOCKTEST_AVERAGE_COLUMN
        else:
            continue
        ws = wb[sheet_name]

        complete, _, _, CLASS_NAME_COLUMN, _, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(ws)
        if not complete: return False, None

        for i in range(2, form_ws.max_row+1): # 데일리데이터 기록 양식 루프
            # 반 필터링
            if (form_ws.cell(i, DataForm.CLASS_NAME_COLUMN).value is not None) and (form_ws.cell(i, TEST_NAME_COLUMN).value is not None):
                class_name   = form_ws.cell(i, DataForm.CLASS_NAME_COLUMN).value
                test_name    = form_ws.cell(i, TEST_NAME_COLUMN).value
                test_average = form_ws.cell(i, TEST_AVERAGE_COLUMN).value

                #반 시작 찾기
                for row in range(2, ws.max_row+1):
                    if ws.cell(row, CLASS_NAME_COLUMN).value == class_name:
                        CLASS_START = row
                        break
                else:
                    OmikronLog.warning(f"{sheet_name} 시트: {class_name} 반이 존재하지 않습니다.")
                    continue

                # 반 끝 찾기
                for row in range(CLASS_START, ws.max_row+1):
                    if ws.cell(row, STUDENT_NAME_COLUMN).value == "시험 평균":
                        CLASS_END = row
                        break

                # 데이터 작성 열 찾기
                for col in range(AVERAGE_SCORE_COLUMN+1, ws.max_column+2):
                    test_date = ws.cell(CLASS_START, col).value
                    if type(test_date) == datetime and test_date.strftime("%y%m%d") == datetime.today().strftime("%y%m%d"):
                        WRITE_COLUMN = col
                        break
                    elif test_date is None:
                        WRITE_COLUMN = col
                        break

                # 입력 틀 작성
                AVERAGE_FORMULA = f"=ROUND(AVERAGE({gcl(WRITE_COLUMN)+str(CLASS_START + 2)}:{gcl(WRITE_COLUMN)+str(CLASS_END - 1)}), 0)"
                ws.column_dimensions[gcl(WRITE_COLUMN)].width    = 14
                ws.cell(CLASS_START, WRITE_COLUMN).value         = datetime.today().date()
                ws.cell(CLASS_START, WRITE_COLUMN).number_format = "yyyy.mm.dd(aaa)"
                ws.cell(CLASS_START, WRITE_COLUMN).alignment     = Alignment(horizontal="center", vertical="center")
                ws.cell(CLASS_START, WRITE_COLUMN).border        = Border(top = Side(border_style="medium", color="000000"))

                ws.cell(CLASS_START + 1, WRITE_COLUMN).value     = test_name
                ws.cell(CLASS_START + 1, WRITE_COLUMN).alignment = Alignment(horizontal="center", vertical="center", wrapText=True)
                ws.cell(CLASS_START + 1, WRITE_COLUMN).border    = Border(bottom = Side(border_style="thin", color="909090"))

                ws.cell(CLASS_END, WRITE_COLUMN).value           = AVERAGE_FORMULA
                ws.cell(CLASS_END, WRITE_COLUMN).font            = Font(bold=True)
                ws.cell(CLASS_END, WRITE_COLUMN).alignment       = Alignment(horizontal="center", vertical="center")
                ws.cell(CLASS_END, WRITE_COLUMN).border          = Border(top = Side(border_style="thin", color="909090"), bottom = Side(border_style="medium", color="000000"))
                
                if type(test_average) in (int, float):
                    ws.cell(CLASS_END, WRITE_COLUMN).fill = class_average_color(test_average)

            test_score   = form_ws.cell(i, TEST_SCORE_COLUMN).value
            student_name = form_ws.cell(i, DataForm.STUDENT_NAME_COLUMN).value

            if test_score is None:
                continue # 점수 없으면 미응시 처리

            # 학생 찾기
            for row in range(CLASS_START + 2, CLASS_END):
                if ws.cell(row, STUDENT_NAME_COLUMN).value == student_name:
                    ws.cell(row, WRITE_COLUMN).value = test_score
                    if type(test_score) in (int, float):
                        ws.cell(row, WRITE_COLUMN).fill = test_score_color(test_score)

                    ws.cell(row, WRITE_COLUMN).alignment = Alignment(horizontal="center", vertical="center")
                    break
            else:
                OmikronLog.warning(f"{sheet_name} 시트: {class_name} 반에 {student_name} 학생이 존재하지 않습니다.")

    ws = wb[DataFile.FIRST_SHEET_NAME]
    save_to_temp(wb)

    # 조건부 서식 수식 로딩
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(f"{os.getcwd()}\\data\\{DataFile.TEMP_FILE_NAME}.xlsx")
    wb.Save()
    wb.Close()
    excel.Quit()
    pythoncom.CoUninitialize()

    wb           = open_temp()
    data_only_wb = open_temp(data_only=True)

    for sheet_name in wb.sheetnames:
        if sheet_name not in (DataFile.FIRST_SHEET_NAME, DataFile.SECOND_SHEET_NAME):
            continue

        ws           = wb[sheet_name]
        data_only_ws = data_only_wb[sheet_name]

        complete, _, _, _, _, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(ws)
        if not complete: return False, None

        for row in range(2, data_only_ws.max_row+1):
            if data_only_ws.cell(row, STUDENT_NAME_COLUMN).value is None:
                break

            # 학생 별 평균 점수에 대한 조건부 서식
            student_average = data_only_ws.cell(row, AVERAGE_SCORE_COLUMN).value
            if type(student_average) in (int, float):
                if ws.cell(row, STUDENT_NAME_COLUMN).value == "시험 평균":
                    ws.cell(row, AVERAGE_SCORE_COLUMN).fill = class_average_color(student_average)
                else:
                    ws.cell(row, AVERAGE_SCORE_COLUMN).fill = student_average_color(student_average)

            # 신규생 하이라이트
            if ws.cell(row, STUDENT_NAME_COLUMN).value in ("날짜", "시험명", "시험 평균"):
                continue
            if ws.cell(row, STUDENT_NAME_COLUMN).font.strike:
                continue
            if ws.cell(row, STUDENT_NAME_COLUMN).font.color is not None and ws.cell(row, STUDENT_NAME_COLUMN).font.color.rgb == "FFFF0000":
                continue

            complete, _, _, new_student = omikron.studentinfo.get_student_info(student_ws, ws.cell(row, STUDENT_NAME_COLUMN).value)
            if complete:
                if new_student:
                    ws.cell(row, STUDENT_NAME_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("FFFF00"))
                else:
                    ws.cell(row, STUDENT_NAME_COLUMN).fill = PatternFill(fill_type=None)
            else:
                ws.cell(row, STUDENT_NAME_COLUMN).fill = PatternFill(fill_type=None)
                OmikronLog.warning(f"{ws.cell(row, STUDENT_NAME_COLUMN).value} 학생 정보가 존재하지 않습니다.")

    ws = wb[DataFile.FIRST_SHEET_NAME]

    return True, wb

def save_individual_test_data(target_row:int, target_col:int, test_score:int|float):
    # 백업 생성
    make_backup_file()

    wb = open()
    ws = wb[DataFile.FIRST_SHEET_NAME]

    # 시험 점수 기록
    ws.cell(target_row, target_col).value     = test_score
    ws.cell(target_row, target_col).fill      = test_score_color(test_score)
    ws.cell(target_row, target_col).alignment = Alignment(horizontal="center", vertical="center")

    save_to_temp(wb)

    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(f"{os.getcwd()}\\data\\{DataFile.TEMP_FILE_NAME}.xlsx")
    wb.Save()
    wb.Close()
    excel.Quit()
    pythoncom.CoUninitialize()

    wb           = open_temp()
    data_only_wb = open_temp(True)

    ws           = wb[DataFile.FIRST_SHEET_NAME]
    data_only_ws = data_only_wb[DataFile.FIRST_SHEET_NAME]

    complete, _, _, _, _, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(ws)
    if not complete: return False, None, None

    # 학생 평균 조건부 서식 반영
    student_average = data_only_ws.cell(target_row, AVERAGE_SCORE_COLUMN).value
    if type(student_average)  in (int, float):
        ws.cell(target_row, AVERAGE_SCORE_COLUMN).fill = student_average_color(student_average)

    # 시험 평균 조건부 서식 반영
    test_average_row = target_row
    while data_only_ws.cell(test_average_row, STUDENT_NAME_COLUMN).value != "시험 평균":
        test_average_row += 1

    test_average = data_only_ws.cell(test_average_row, target_col).value
    if type(test_average) in (int, float):
        ws.cell(test_average_row, target_col).fill = class_average_color(test_average)

    # 반 평균 조건부 서식 반영
    class_average = data_only_ws.cell(test_average_row, AVERAGE_SCORE_COLUMN).value
    if type(class_average) in (int, float):
        ws.cell(test_average_row, AVERAGE_SCORE_COLUMN).fill = class_average_color(test_average)

    return True, test_average, wb

def conditional_formatting():
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(f"{os.getcwd()}\\data\\{omikron.config.DATA_FILE_NAME}.xlsx")
    wb.Save()
    wb.Close()
    excel.Quit()
    pythoncom.CoUninitialize()

    wb                   = open()
    data_only_wb         = open(data_only=True)
    student_wb           = omikron.studentinfo.open()
    complete, student_ws = omikron.studentinfo.open_worksheet(student_wb)
    if not complete: return False

    for sheet_name in wb.sheetnames:
        if sheet_name not in (DataFile.FIRST_SHEET_NAME, DataFile.SECOND_SHEET_NAME):
            continue

        ws           = wb[sheet_name]
        data_only_ws = data_only_wb[sheet_name]

        complete, _, _, _, _, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(ws)
        if not complete: return False

        for row in range(2, ws.max_row+1):
            if ws.cell(row, STUDENT_NAME_COLUMN).value is None:
                break
            if ws.cell(row, STUDENT_NAME_COLUMN).value == "날짜":
                DATE_ROW = row
            if ws.cell(row, STUDENT_NAME_COLUMN).value != "시험명":
                ws.row_dimensions[row].height = 18

            # 데이터 조건부 서식
            for col in range(1, data_only_ws.max_column+1):
                if col > AVERAGE_SCORE_COLUMN and ws.cell(DATE_ROW, col).value is None:
                    break

                ws.column_dimensions[gcl(col)].width = 14
                if ws.cell(row, STUDENT_NAME_COLUMN).value == "날짜":
                    ws.cell(row, col).border = Border(top = Side(border_style="medium", color="000000"))
                elif ws.cell(row, STUDENT_NAME_COLUMN).value == "시험명":
                    ws.cell(row, col).border = Border(bottom = Side(border_style="thin", color="909090"))
                elif ws.cell(row, STUDENT_NAME_COLUMN).value == "시험 평균":
                    ws.cell(row, col).border = Border(top = Side(border_style="thin", color="909090"), bottom = Side(border_style="medium", color="000000"))
                else:
                    ws.cell(row, col).border = None

                # 학생 평균 점수 열 기준 분기   
                if col <= AVERAGE_SCORE_COLUMN:
                    continue

                if ws.cell(row, STUDENT_NAME_COLUMN).value == "시험명":
                    ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center", wrapText=True)
                elif data_only_ws.cell(row, STUDENT_NAME_COLUMN).value == "시험 평균":
                    ws.cell(row, col).font = Font(bold=True)
                    if type(data_only_ws.cell(row, col).value) in (int, float):
                        ws.cell(row, col).fill = class_average_color(data_only_ws.cell(row, col).value)
                elif type(data_only_ws.cell(row, col).value) in (int, float):
                    ws.cell(row, col).fill = test_score_color(data_only_ws.cell(row, col).value)
                else:
                    ws.cell(row, col).fill = PatternFill(fill_type=None)

            # 학생별 평균 조건부 서식
            if type(data_only_ws.cell(row, AVERAGE_SCORE_COLUMN).value) in (int, float):
                if ws.cell(row, STUDENT_NAME_COLUMN).value == "시험 평균":
                    ws.cell(row, AVERAGE_SCORE_COLUMN).fill = class_average_color(data_only_ws.cell(row, AVERAGE_SCORE_COLUMN).value)
                else:
                    ws.cell(row, AVERAGE_SCORE_COLUMN).fill = student_average_color(data_only_ws.cell(row, AVERAGE_SCORE_COLUMN).value)
            else:
                ws.cell(row, col).fill = PatternFill(fill_type=None)

            # 학생별 평균 폰트 설정
            if ws.cell(row, STUDENT_NAME_COLUMN).value in ("날짜", "시험명", "시험 평균"):
                ws.cell(row, AVERAGE_SCORE_COLUMN).font = Font(bold=True)
                continue
            if ws.cell(row, STUDENT_NAME_COLUMN).font.strike:
                ws.cell(row, AVERAGE_SCORE_COLUMN).font = Font(bold=True, strike=True)
                continue
            if ws.cell(row, STUDENT_NAME_COLUMN).font.color is not None and ws.cell(row, STUDENT_NAME_COLUMN).font.color.rgb == "FFFF0000":
                ws.cell(row, AVERAGE_SCORE_COLUMN).font = Font(bold=True, color="FFFF0000")
                continue

            # 신규생 하이라이트
            complete, _, _, new_student = omikron.studentinfo.get_student_info(student_ws, ws.cell(row, STUDENT_NAME_COLUMN).value)
            if complete:
                if new_student:
                    ws.cell(row, STUDENT_NAME_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("FFFF00"))
                else:
                    ws.cell(row, STUDENT_NAME_COLUMN).fill = PatternFill(fill_type=None)
            else:
                ws.cell(row, STUDENT_NAME_COLUMN).fill = PatternFill(fill_type=None)
                OmikronLog.warning(f"{ws.cell(row, STUDENT_NAME_COLUMN).value} 학생 정보가 존재하지 않습니다.")

    save(wb)
    return True

def update_class():
    """
    수정된 반 정보 파일을 바탕으로 데이터 파일 업데이트
    """
    make_backup_file()

    complete, deleted_class_names, unregistered_class_names = omikron.classinfo.check_difference_between()
    if not complete: return False, None

    # 조건부 서식 수식 로딩
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(f"{os.getcwd()}\\data\\{omikron.config.DATA_FILE_NAME}.xlsx")
    wb.Save()
    wb.Close()
    excel.Quit()
    pythoncom.CoUninitialize()

    wb = open()

    if len(deleted_class_names) > 0:
        if not os.path.isfile(f"./data/{DataFile.POST_DATA_FILE_NAME}.xlsx"):
            wb = xl.Workbook()
            ws = wb.worksheets[0]
            ws.title = DataFile.FIRST_SHEET_NAME
            ws[gcl(DataFile.TEST_TIME_COLUMN)+"1"]     = "시간"
            ws[gcl(DataFile.CLASS_WEEKDAY_COLUMN)+"1"] = "요일"
            ws[gcl(DataFile.CLASS_NAME_COLUMN)+"1"]    = "반"
            ws[gcl(DataFile.TEACHER_NAME_COLUMN)+"1"]  = "담당"
            ws[gcl(DataFile.STUDENT_NAME_COLUMN)+"1"]  = "이름"
            ws[gcl(DataFile.AVERAGE_SCORE_COLUMN)+"1"] = "학생 평균"
            ws.freeze_panes    = f"{gcl(DataFile.DATA_COLUMN)}2"
            ws.auto_filter.ref = f"A:{gcl(DataFile.MAX)}"

            for col in range(1, DataFile.DATA_COLUMN):
                ws.cell(1, col).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(1, col).border    = Border(bottom = Side(border_style="medium", color="000000"))
            
            # 모의고사 sheet 생성
            copy_ws                 = wb.copy_worksheet(wb[DataFile.FIRST_SHEET_NAME])
            copy_ws.title           = DataFile.SECOND_SHEET_NAME
            copy_ws.freeze_panes    = f"{gcl(DataFile.DATA_COLUMN)}2"
            copy_ws.auto_filter.ref = f"A:{gcl(DataFile.MAX)}"

            wb.save(f"./data/{DataFile.POST_DATA_FILE_NAME}.xlsx")

        data_only_wb = open(data_only=True)
        post_data_wb = xl.load_workbook(f"./data/{DataFile.POST_DATA_FILE_NAME}.xlsx")
        for sheet_name in wb.sheetnames:
            if sheet_name not in (DataFile.FIRST_SHEET_NAME, DataFile.SECOND_SHEET_NAME):
                continue

            data_only_ws = data_only_wb[sheet_name]
            post_data_ws = post_data_wb[sheet_name]
            ws           = wb[sheet_name]

            complete, TEST_TIME_COLUMN, CLASS_WEEKDAY_COLUMN, CLASS_NAME_COLUMN, TEACHER_NAME_COLUMN, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(ws)
            if not complete: return False, None

            for row in range(2, ws.max_row+1):
                while ws.cell(row, CLASS_NAME_COLUMN).value in deleted_class_names:
                    ws.delete_rows(row)
            
            for row in range(2, data_only_ws.max_row+1):
                if data_only_ws.cell(row, CLASS_NAME_COLUMN).value in deleted_class_names:
                    POST_DATA_WRITE_ROW = post_data_ws.max_row+1
                    copy_cell(post_data_ws.cell(POST_DATA_WRITE_ROW, DataFile.TEST_TIME_COLUMN),     data_only_ws.cell(row, TEST_TIME_COLUMN))
                    copy_cell(post_data_ws.cell(POST_DATA_WRITE_ROW, DataFile.CLASS_WEEKDAY_COLUMN), data_only_ws.cell(row, CLASS_WEEKDAY_COLUMN))
                    copy_cell(post_data_ws.cell(POST_DATA_WRITE_ROW, DataFile.CLASS_NAME_COLUMN),    data_only_ws.cell(row, CLASS_NAME_COLUMN))
                    copy_cell(post_data_ws.cell(POST_DATA_WRITE_ROW, DataFile.TEACHER_NAME_COLUMN),  data_only_ws.cell(row, TEACHER_NAME_COLUMN))
                    copy_cell(post_data_ws.cell(POST_DATA_WRITE_ROW, DataFile.STUDENT_NAME_COLUMN),  data_only_ws.cell(row, STUDENT_NAME_COLUMN))
                    copy_cell(post_data_ws.cell(POST_DATA_WRITE_ROW, DataFile.AVERAGE_SCORE_COLUMN), data_only_ws.cell(row, AVERAGE_SCORE_COLUMN))
                    POST_DATA_WRITE_COLUMN = DataFile.MAX+1
                    for col in range(AVERAGE_SCORE_COLUMN+1, data_only_ws.max_column+1):
                        copy_cell(post_data_ws.cell(POST_DATA_WRITE_ROW, POST_DATA_WRITE_COLUMN), data_only_ws.cell(row, col))
                        ws.column_dimensions[gcl(POST_DATA_WRITE_COLUMN)].width = 14
                        POST_DATA_WRITE_COLUMN += 1
            
            ws.auto_filter.ref = f"A:{gcl(AVERAGE_SCORE_COLUMN)}"
        
        post_data_wb.save(f"./data/{DataFile.POST_DATA_FILE_NAME}.xlsx")

    if len(unregistered_class_names) > 0:
        class_wb = omikron.classinfo.open_temp()
        complete, class_ws = omikron.classinfo.open_worksheet(class_wb)
        if not complete: return False, None

        class_student_dict = omikron.chrome.get_class_student_dict()

        for sheet_name in wb.sheetnames:
            if sheet_name not in (DataFile.FIRST_SHEET_NAME, DataFile.SECOND_SHEET_NAME):
                continue

            ws = wb[sheet_name]

            complete, TEST_TIME_COLUMN, CLASS_WEEKDAY_COLUMN, CLASS_NAME_COLUMN, TEACHER_NAME_COLUMN, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(ws)
            if not complete: return False, None

            for row in range(ws.max_row+1, 1, -1):
                if ws.cell(row-1, DataFile.STUDENT_NAME_COLUMN).value is not None:
                    WRITE_RANGE = WRITE_LOCATION = row
                    break

            for class_name in unregistered_class_names:
                if len(class_student_dict[class_name]) == 0:
                    continue
                complete, teacher_name, class_weekday, test_time = omikron.classinfo.get_class_info(class_ws, class_name)
                if not complete: continue

                # 시험명
                ws.cell(WRITE_LOCATION, TEST_TIME_COLUMN).value     = test_time
                ws.cell(WRITE_LOCATION, CLASS_WEEKDAY_COLUMN).value = class_weekday
                ws.cell(WRITE_LOCATION, CLASS_NAME_COLUMN).value    = class_name
                ws.cell(WRITE_LOCATION, TEACHER_NAME_COLUMN).value  = teacher_name
                ws.cell(WRITE_LOCATION, STUDENT_NAME_COLUMN).value  = "날짜"
                WRITE_LOCATION += 1
                
                ws.cell(WRITE_LOCATION, TEST_TIME_COLUMN).value     = test_time
                ws.cell(WRITE_LOCATION, CLASS_WEEKDAY_COLUMN).value = class_weekday
                ws.cell(WRITE_LOCATION, CLASS_NAME_COLUMN).value    = class_name
                ws.cell(WRITE_LOCATION, TEACHER_NAME_COLUMN).value  = teacher_name
                ws.cell(WRITE_LOCATION, STUDENT_NAME_COLUMN).value  = "시험명"

                for col in range(1, AVERAGE_SCORE_COLUMN + 1):
                    ws.cell(WRITE_LOCATION, col).border = Border(bottom = Side(border_style="thin", color="909090"))

                WRITE_LOCATION += 1

                # 학생 루프
                for studnet_name in class_student_dict[class_name]:
                    ws.cell(WRITE_LOCATION, TEST_TIME_COLUMN).value     = test_time
                    ws.cell(WRITE_LOCATION, CLASS_WEEKDAY_COLUMN).value = class_weekday
                    ws.cell(WRITE_LOCATION, CLASS_NAME_COLUMN).value    = class_name
                    ws.cell(WRITE_LOCATION, TEACHER_NAME_COLUMN).value  = teacher_name
                    ws.cell(WRITE_LOCATION, STUDENT_NAME_COLUMN).value  = studnet_name
                    WRITE_LOCATION += 1
                
                # 시험별 평균
                ws.cell(WRITE_LOCATION, TEST_TIME_COLUMN).value     = test_time
                ws.cell(WRITE_LOCATION, CLASS_WEEKDAY_COLUMN).value = class_weekday
                ws.cell(WRITE_LOCATION, CLASS_NAME_COLUMN).value    = class_name
                ws.cell(WRITE_LOCATION, TEACHER_NAME_COLUMN).value  = teacher_name
                ws.cell(WRITE_LOCATION, STUDENT_NAME_COLUMN).value  = "시험 평균"

                for col in range(1, AVERAGE_SCORE_COLUMN+1):
                    ws.cell(WRITE_LOCATION, col).border = Border(top = Side(border_style="thin", color="909090"), bottom = Side(border_style="medium", color="000000"))

                WRITE_LOCATION += 1

            # 정렬
            for row in range(WRITE_RANGE, ws.max_row + 1):
                for col in range(1, AVERAGE_SCORE_COLUMN + 1):
                    ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")

            # 필터 범위 재지정
            ws.auto_filter.ref = f"A:{gcl(AVERAGE_SCORE_COLUMN)}"

    return rescoping_formula(wb)

def add_student(student_name:str, target_class_name:str, wb:xl.Workbook=None):
    """
    학생 추가
    
    `move_student` 작업 시 `wb`로 작업중인 파일 정보 전달
    """
    if wb is None:
        wb = open()

    for ws in wb.worksheets:
        if ws.title not in (DataFile.FIRST_SHEET_NAME, DataFile.SECOND_SHEET_NAME):
            continue

        complete, TEST_TIME_COLUMN, CLASS_WEEKDAY_COLUMN, CLASS_NAME_COLUMN, TEACHER_NAME_COLUMN, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(ws)
        if not complete: return False, None

        # 목표 반에 학생 추가
        for row in range(2, ws.max_row+1):
            if ws.cell(row, CLASS_NAME_COLUMN).value == target_class_name:
                class_index = row+2
                break
        else:
            OmikronLog.warning(f"{ws.title} 시트에 {target_class_name} 반이 존재하지 않습니다.")
            continue

        while ws.cell(class_index, STUDENT_NAME_COLUMN).value != "시험 평균":
            if ws.cell(class_index, STUDENT_NAME_COLUMN).value > student_name:
                break
            elif ws.cell(class_index, STUDENT_NAME_COLUMN).font.strike:
                class_index += 1
            elif ws.cell(class_index, STUDENT_NAME_COLUMN).font.color is not None and ws.cell(class_index, STUDENT_NAME_COLUMN).font.color.rgb == "FFFF0000":
                class_index += 1
            elif ws.cell(class_index, STUDENT_NAME_COLUMN).value == student_name:
                OmikronLog.error(f"{student_name} 학생이 이미 존재합니다.")
                return False, None
            else:
                class_index += 1

        ws.insert_rows(class_index)
        ws.cell(class_index, TEST_TIME_COLUMN).value         = ws.cell(class_index-1, TEST_TIME_COLUMN).value
        ws.cell(class_index, CLASS_WEEKDAY_COLUMN).value     = ws.cell(class_index-1, CLASS_WEEKDAY_COLUMN).value
        ws.cell(class_index, CLASS_NAME_COLUMN).value        = ws.cell(class_index-1, CLASS_NAME_COLUMN).value
        ws.cell(class_index, TEACHER_NAME_COLUMN).value      = ws.cell(class_index-1, TEACHER_NAME_COLUMN).value
        ws.cell(class_index, STUDENT_NAME_COLUMN).value      = student_name

        ws.cell(class_index, TEST_TIME_COLUMN).alignment     = Alignment(horizontal="center", vertical="center")
        ws.cell(class_index, CLASS_WEEKDAY_COLUMN).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(class_index, CLASS_NAME_COLUMN).alignment    = Alignment(horizontal="center", vertical="center")
        ws.cell(class_index, TEACHER_NAME_COLUMN).alignment  = Alignment(horizontal="center", vertical="center")
        ws.cell(class_index, STUDENT_NAME_COLUMN).alignment  = Alignment(horizontal="center", vertical="center")

        ws.cell(class_index, AVERAGE_SCORE_COLUMN).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(class_index, AVERAGE_SCORE_COLUMN).font      = Font(bold=True)

    return rescoping_formula(wb)

def delete_student(student_name:str):
    """
    학생 퇴원 처리
    
    퇴원 처리된 학생은 모든 데이터에 취소선 적용
    """
    wb = open()
    for ws in wb.worksheets:
        if ws.title not in (DataFile.FIRST_SHEET_NAME, DataFile.SECOND_SHEET_NAME):
            continue

        complete, _, _, _, _, STUDENT_NAME_COLUMN, _ = find_dynamic_columns(ws)
        if not complete: return False, None

        for row in range(2, ws.max_row+1):
            if ws.cell(row, STUDENT_NAME_COLUMN).value == student_name:
                for col in range(1, ws.max_column+1):
                    ws.cell(row, col).font = Font(strike=True)

    return True, wb

def move_student(student_name:str, target_class_name:str, current_class_name:str):
    """
    학생 반 이동

    학생의 기존 반 데이터 글꼴 색을 빨간색으로 변경 후 목표 반에 학생 추가
    """
    wb = open()

    for ws in wb.worksheets:
        if ws.title not in (DataFile.FIRST_SHEET_NAME, DataFile.SECOND_SHEET_NAME):
            continue

        complete, _, _, CLASS_NAME_COLUMN, _, STUDENT_NAME_COLUMN, _ = find_dynamic_columns(ws)
        if not complete: return False, None

        # 기존 반 데이터 빨간색 처리
        for row in range(2, ws.max_row+1):
            if ws.cell(row, STUDENT_NAME_COLUMN).value == student_name and ws.cell(row, CLASS_NAME_COLUMN).value == current_class_name:
                for col in range(1, ws.max_column+1):
                    if ws.cell(row, col).font.bold:
                        ws.cell(row, col).font = Font(bold=True, color="FFFF0000")
                    else:
                        ws.cell(row, col).font = Font(color="FFFF0000")
                break

    return add_student(student_name, target_class_name, wb)

def rescoping_formula(wb:xl.Workbook):
    """
    데이터 파일 내 평균 산출 수식의 범위 재조정
    """
    for ws in wb.worksheets:
        if ws.title not in (DataFile.FIRST_SHEET_NAME, DataFile.SECOND_SHEET_NAME):
            continue

        complete, _, _, _, _, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(ws)
        if not complete: return False, None

        # 평균 범위 재지정
        for row in range(2, ws.max_row+1):
            if ws.cell(row, STUDENT_NAME_COLUMN).value is None:
                break
            striked = False
            colored = False
            if ws.cell(row, STUDENT_NAME_COLUMN).font.strike:
                striked = True
            if ws.cell(row, STUDENT_NAME_COLUMN).font.color is not None:
                if ws.cell(row, STUDENT_NAME_COLUMN).font.color.rgb == "FFFF0000":
                    colored = True

            if ws.cell(row, STUDENT_NAME_COLUMN).value == "날짜":
                DATE_ROW = row
                CLASS_START = row+2
            elif ws.cell(row, STUDENT_NAME_COLUMN).value == "시험 평균":
                CLASS_END = row-1
                ws[f"{gcl(AVERAGE_SCORE_COLUMN)}{row}"] = ArrayFormula(f"{gcl(AVERAGE_SCORE_COLUMN)}{row}", f"=ROUND(AVERAGE(IFERROR({gcl(AVERAGE_SCORE_COLUMN)}{CLASS_START}:{gcl(AVERAGE_SCORE_COLUMN)}{CLASS_END}, \"\")), 0)")
                if CLASS_START >= CLASS_END:
                    continue
                for col in range(AVERAGE_SCORE_COLUMN+1, ws.max_column+1):
                    if ws.cell(DATE_ROW, col).value is None:
                        break
                    ws.cell(row, col).value = f"=ROUND(AVERAGE({gcl(col)}{CLASS_START}:{gcl(col)}{CLASS_END}), 0)"
                    ws.cell(row, col).font  = Font(bold=True)
            elif ws.cell(row, STUDENT_NAME_COLUMN).value == "시험명":
                continue
            else:
                ws.cell(row, AVERAGE_SCORE_COLUMN).value = f"=ROUND(AVERAGE({gcl(AVERAGE_SCORE_COLUMN+1)}{row}:XFD{row}), 0)"

            if striked:
                ws.cell(row, AVERAGE_SCORE_COLUMN).font = Font(bold=True, strike=True)
            elif colored:
                ws.cell(row, AVERAGE_SCORE_COLUMN).font = Font(bold=True, color="FFFF0000")
            else:
                ws.cell(row, AVERAGE_SCORE_COLUMN).font = Font(bold=True)

    return True, wb
