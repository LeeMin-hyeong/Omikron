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
import omikron.datafile
import omikron.dataform
import omikron.makeuptest
import omikron.studentinfo

from omikron.config import DATA_FILE_NAME
from omikron.defs import DataFile, DataForm
from omikron.log import OmikronLog
from omikron.util import copy_cell, class_average_color, student_average_color, test_score_color

# 파일 기본 작업
def make_file() -> bool:
    ini_wb = xl.Workbook()
    ini_ws = ini_wb.worksheets[0]
    ini_ws.title = "데일리테스트"
    ini_ws[gcl(DataFile.TEST_TIME_COLUMN)+"1"]     = "시간"
    ini_ws[gcl(DataFile.CLASS_WEEKDAY_COLUMN)+"1"] = "요일"
    ini_ws[gcl(DataFile.CLASS_NAME_COLUMN)+"1"]    = "반"
    ini_ws[gcl(DataFile.TEACHER_NAME_COLUMN)+"1"]  = "담당"
    ini_ws[gcl(DataFile.STUDENT_NAME_COLUMN)+"1"]  = "이름"
    ini_ws[gcl(DataFile.AVERAGE_SCORE_COLUMN)+"1"] = "학생 평균"
    ini_ws.freeze_panes    = f"{gcl(DataFile.DATA_COLUMN)}2"
    ini_ws.auto_filter.ref = f"A:{gcl(DataFile.MAX)}"

    for col in range(1, DataFile.DATA_COLUMN):
        ini_ws.cell(1, col).border = Border(bottom = Side(border_style="medium", color="000000"))

    class_wb = omikron.classinfo.open(True)
    complete, class_ws = omikron.classinfo.open_worksheet(class_wb)
    if not complete: return False

    # 반 루프
    for class_name, student_list in omikron.chrome.get_class_student_dict().items():
        if len(student_list) == 0:
            continue

        complete, teacher_name, class_weekday, test_time = omikron.classinfo.get_class_info(class_ws, class_name)
        if not complete: continue

        WRITE_LOCATION = ini_ws.max_row + 1

        # 시험명
        ini_ws.cell(WRITE_LOCATION, DataFile.TEST_TIME_COLUMN).value     = test_time
        ini_ws.cell(WRITE_LOCATION, DataFile.CLASS_WEEKDAY_COLUMN).value = class_weekday
        ini_ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value    = class_name
        ini_ws.cell(WRITE_LOCATION, DataFile.TEACHER_NAME_COLUMN).value  = teacher_name
        ini_ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value  = "날짜"
        
        WRITE_LOCATION = ini_ws.max_row + 1
        ini_ws.cell(WRITE_LOCATION, DataFile.TEST_TIME_COLUMN).value     = test_time
        ini_ws.cell(WRITE_LOCATION, DataFile.CLASS_WEEKDAY_COLUMN).value = class_weekday
        ini_ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value    = class_name
        ini_ws.cell(WRITE_LOCATION, DataFile.TEACHER_NAME_COLUMN).value  = teacher_name
        ini_ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value  = "시험명"

        for col in range(1, DataFile.DATA_COLUMN):
            ini_ws.cell(WRITE_LOCATION, col).border = Border(bottom = Side(border_style="thin", color="909090"))

        class_start = WRITE_LOCATION + 1

        # 학생 루프
        for student_name in student_list:
            WRITE_LOCATION = ini_ws.max_row + 1
            ini_ws.cell(WRITE_LOCATION, DataFile.TEST_TIME_COLUMN).value     = test_time
            ini_ws.cell(WRITE_LOCATION, DataFile.CLASS_WEEKDAY_COLUMN).value = class_weekday
            ini_ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value    = class_name
            ini_ws.cell(WRITE_LOCATION, DataFile.TEACHER_NAME_COLUMN).value  = teacher_name
            ini_ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value  = student_name
            ini_ws.cell(WRITE_LOCATION, DataFile.AVERAGE_SCORE_COLUMN).value = f"=ROUND(AVERAGE(G{str(WRITE_LOCATION)}:XFD{str(WRITE_LOCATION)}), 0)"
            ini_ws.cell(WRITE_LOCATION, DataFile.AVERAGE_SCORE_COLUMN).font  = Font(bold=True)
        
        # 시험별 평균
        class_end = WRITE_LOCATION
        WRITE_LOCATION = ini_ws.max_row + 1
        ini_ws.cell(WRITE_LOCATION, DataFile.TEST_TIME_COLUMN).value     = test_time
        ini_ws.cell(WRITE_LOCATION, DataFile.CLASS_WEEKDAY_COLUMN).value = class_weekday
        ini_ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value    = class_name
        ini_ws.cell(WRITE_LOCATION, DataFile.TEACHER_NAME_COLUMN).value  = teacher_name
        ini_ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value  = "시험 평균"
        ini_ws[f"F{str(WRITE_LOCATION)}"] = ArrayFormula(f"F{str(WRITE_LOCATION)}", f"=ROUND(AVERAGE(IFERROR(F{str(class_start)}:F{str(class_end)}, \"\")), 0)")
        ini_ws.cell(WRITE_LOCATION, DataFile.AVERAGE_SCORE_COLUMN).font = Font(bold=True)

        for col in range(1, DataFile.DATA_COLUMN):
            ini_ws.cell(WRITE_LOCATION, col).border = Border(top = Side(border_style="thin", color="909090"), bottom = Side(border_style="medium", color="000000"))

    # 정렬
    for row in range(1, ini_ws.max_row + 1):
        for col in range(1, ini_ws.max_column + 1):
            ini_ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")
    
    # 모의고사 sheet 생성
    copy_ws                 = ini_wb.copy_worksheet(ini_wb["데일리테스트"])
    copy_ws.title           = "모의고사"
    copy_ws.freeze_panes    = f"{gcl(DataFile.DATA_COLUMN)}2"
    copy_ws.auto_filter.ref = f"A:{gcl(DataFile.MAX)}"

    omikron.classinfo.close(class_wb)
    save(ini_wb)

    return True

def open(data_only:bool=False) -> xl.Workbook:
    return xl.load_workbook(f"./data/{DATA_FILE_NAME}.xlsx", data_only=data_only)

def save(data_file_wb:xl.Workbook):
    data_file_wb.save(f"./data/{DATA_FILE_NAME}.xlsx")
    data_file_wb.close()

def save_to_temp(data_file_wb:xl.Workbook):
    data_file_wb.save(f"./data/{DataFile.TEMP_FILE_NAME}.xlsx")
    os.system(f"attrib +h ./data/{DataFile.TEMP_FILE_NAME}.xlsx")

def open_temp(data_only:bool=False) -> xl.Workbook:
    return xl.load_workbook(f"./data/{DataFile.TEMP_FILE_NAME}.xlsx", data_only=data_only)

def delete_temp():
    os.remove(f"./data/{DataFile.TEMP_FILE_NAME}.xlsx")

def close(data_file_wb:xl.Workbook):
    data_file_wb.close()

def isopen() -> bool:
    return os.path.isfile(f"./data/~${DATA_FILE_NAME}.xlsx")

# 파일 유틸리티
def make_backup_file():
    data_file_wb = open()
    data_file_wb.save(f"./data/backup/{DATA_FILE_NAME}({datetime.today().strftime('%Y%m%d')}).xlsx")
    data_file_wb.close()

def get_data_sorted_dict():
    """
    워크시트를 입력받아 해당 워크시트로부터 반:테스트 리스트 정보 추출

    워크시트를 지정하지 않으면 파일을 열어 정보 추출

    abc가나다 순으로 정렬

    return : {반 : {학생 : 인덱스}}, {반 : {시험 : 인덱스}}
    """
    data_file_wb = open()
    data_file_ws = data_file_wb["데일리테스트"]

    complete, _, _, CLASS_NAME_COLUMN, _, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(data_file_ws)
    if not complete: return False, None, None

    class_wb = omikron.classinfo.open()
    complete, class_ws = omikron.classinfo.open_worksheet(class_wb)
    if not complete: return False, None, None

    class_student_dict = {}
    class_test_dict    = {}

    for class_name in omikron.classinfo.get_class_names(class_ws):
        student_index_dict = {}
        test_index_dict    = {}
        for row in range(2, data_file_ws.max_row+1):
            if data_file_ws.cell(row, CLASS_NAME_COLUMN).value != class_name:
                continue
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "날짜":
                for col in range(AVERAGE_SCORE_COLUMN+1, data_file_ws.max_row+1):
                    test_date = data_file_ws.cell(row, col).value
                    test_name = data_file_ws.cell(row+1, col).value
                    if test_date is None and test_name is None:
                        break
                    if type(test_date) == datetime:
                        test_date = test_date.strftime("%y.%m.%d")
                    else:
                        test_date = str(test_date).split()[0][2:10].replace("-", ".").replace(",", ".").replace("/", ".")
                    test_index_dict[f"{test_date} {test_name}"] = col
                continue
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value in ("시험명", "시험 평균"):
                continue
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).font.strike:
                continue
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).font.color is not None and data_file_ws.cell(row, STUDENT_NAME_COLUMN).font.color.rgb == "FFFF0000":
                continue
            student_index_dict[data_file_ws.cell(row, STUDENT_NAME_COLUMN).value] = row

        test_index_dict = dict(sorted(test_index_dict.items(), reverse=True))

        class_student_dict[class_name] = student_index_dict
        class_test_dict[class_name]    = test_index_dict

    class_student_dict = dict(sorted(class_student_dict.items()))

    omikron.classinfo.close(class_wb)
    close(data_file_wb)

    return True, class_student_dict, class_test_dict

def find_dynamic_columns(data_file_ws:Worksheet):
    """
    파일 열(column) 정보 동적 탐색

    오류 발생 시 에러 메세지 출력

    complete,

    TEST_TIME_COLUMN,

    CLASS_WEEKDAY_COLUMN,

    CLASS_NAME_COLUMN,

    TEACHER_NAME_COLUMN,

    STUDENT_NAME_COLUMN,

    AVERAGE_SCORE_COLUMN

    일부 열 정보가 누락되면 False를 리턴
    """

    for col in range(1, data_file_ws.max_column+1):
        if data_file_ws.cell(1, col).value == "시간":
            TEST_TIME_COLUMN = col
            break
    else:
        OmikronLog.error(f"{data_file_ws.title} 시트에 '시간' 열이 없습니다.")
        return False, None, None, None, None, None, None

    for col in range(1, data_file_ws.max_column+1):
        if data_file_ws.cell(1, col).value == "요일":
            CLASS_WEEKDAY_COLUMN = col
            break
    else:
        OmikronLog.error(f"{data_file_ws.title} 시트에 '요일' 열이 없습니다.")
        return False, None, None, None, None, None, None

    for col in range(1, data_file_ws.max_column+1):
        if data_file_ws.cell(1, col).value == "반":
            CLASS_NAME_COLUMN = col
            break
    else:
        OmikronLog.error(f"{data_file_ws.title} 시트에 '반' 열이 없습니다.")
        return False, None, None, None, None, None, None

    for col in range(1, data_file_ws.max_column+1):
        if data_file_ws.cell(1, col).value == "담당":
            TEACHER_NAME_COLUMN = col
            break
    else:
        OmikronLog.error(f"{data_file_ws.title} 시트에 '담당' 열이 없습니다.")
        return False, None, None, None, None, None, None

    for col in range(1, data_file_ws.max_column+1):
        if data_file_ws.cell(1, col).value == "이름":
            STUDENT_NAME_COLUMN = col
            break
    else:
        OmikronLog.error(f"{data_file_ws.title} 시트에 '이름' 열이 없습니다.")
        return False, None, None, None, None, None, None

    for col in range(1, data_file_ws.max_column+1):
        if data_file_ws.cell(1, col).value == "학생 평균":
            AVERAGE_SCORE_COLUMN = col
            break
    else:
        OmikronLog.error(f"{data_file_ws.title} 시트에 '학생 평균' 열이 없습니다.")
        return False, None, None, None, None, None, None
    
    return True, TEST_TIME_COLUMN, CLASS_WEEKDAY_COLUMN, CLASS_NAME_COLUMN, TEACHER_NAME_COLUMN, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN

def is_cell_empty(row:int, col:int) -> bool:
    """
    데이터 파일이 열려있지 않을 때 특정 셀의 값이 비어있는 지 확인

    데일리 테스트에 대해서만 지원
    """

    data_file_wb = open(True)
    data_file_ws = data_file_wb["데일리테스트"]

    if data_file_ws.cell(row, col).value is None:
        return True, None

    value = data_file_ws.cell(row, col).value
    close(data_file_wb)

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

    data_file_wb = open()

    if "데일리테스트" not in data_file_wb.sheetnames:
        return False, None
    if "모의고사" not in data_file_wb.sheetnames:
        return False, None

    for sheet_name in data_file_wb.sheetnames:
        data_file_ws = data_file_wb[sheet_name]
        if sheet_name == "데일리테스트":
            TEST_NAME_COLUMN    = DataForm.DAILYTEST_NAME_COLUMN
            TEST_SCORE_COLUMN   = DataForm.DAILYTEST_SCORE_COLUMN
            TEST_AVERAGE_COLUMN = DataForm.DAILYTEST_AVERAGE_COLUMN
        elif sheet_name == "모의고사":
            TEST_NAME_COLUMN    = DataForm.MOCKTEST_NAME_COLUMN
            TEST_SCORE_COLUMN   = DataForm.MOCKTEST_SCORE_COLUMN
            TEST_AVERAGE_COLUMN = DataForm.MOCKTEST_AVERAGE_COLUMN
        else:
            continue

        complete, _, _, CLASS_NAME_COLUMN, _, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(data_file_ws)
        if not complete: return False, None

        for i in range(2, form_ws.max_row+1): # 데일리데이터 기록 양식 루프
            # 반 필터링
            if (form_ws.cell(i, DataForm.CLASS_NAME_COLUMN).value is not None) and (form_ws.cell(i, TEST_NAME_COLUMN).value is not None):
                class_name   = form_ws.cell(i, DataForm.CLASS_NAME_COLUMN).value
                test_name    = form_ws.cell(i, TEST_NAME_COLUMN).value
                test_average = form_ws.cell(i, TEST_AVERAGE_COLUMN).value
                
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
                
                # 데이터 작성 열 찾기
                for col in range(AVERAGE_SCORE_COLUMN+1, data_file_ws.max_column+2):
                    test_date = data_file_ws.cell(CLASS_START, col).value
                    if test_date is None:
                        WRITE_COLUMN = col
                        break
                    if type(test_date) != datetime:
                        continue
                    if test_date.strftime("%y%m%d") == datetime.today().strftime("%y%m%d"):
                        WRITE_COLUMN = col
                        break
                
                # 입력 틀 작성
                AVERAGE_FORMULA = f"=ROUND(AVERAGE({gcl(WRITE_COLUMN)+str(CLASS_START + 2)}:{gcl(WRITE_COLUMN)+str(CLASS_END - 1)}), 0)"
                data_file_ws.column_dimensions[gcl(WRITE_COLUMN)].width    = 14
                data_file_ws.cell(CLASS_START, WRITE_COLUMN).value         = datetime.today().date()
                data_file_ws.cell(CLASS_START, WRITE_COLUMN).number_format = "yyyy.mm.dd(aaa)"
                data_file_ws.cell(CLASS_START, WRITE_COLUMN).alignment     = Alignment(horizontal="center", vertical="center")
                data_file_ws.cell(CLASS_START, WRITE_COLUMN).border        = Border(top = Side(border_style="medium", color="000000"))

                data_file_ws.cell(CLASS_START + 1, WRITE_COLUMN).value     = test_name
                data_file_ws.cell(CLASS_START + 1, WRITE_COLUMN).alignment = Alignment(horizontal="center", vertical="center", wrapText=True)
                data_file_ws.cell(CLASS_START + 1, WRITE_COLUMN).border    = Border(bottom = Side(border_style="thin", color="909090"))

                data_file_ws.cell(CLASS_END, WRITE_COLUMN).value           = AVERAGE_FORMULA
                data_file_ws.cell(CLASS_END, WRITE_COLUMN).font            = Font(bold=True)
                data_file_ws.cell(CLASS_END, WRITE_COLUMN).alignment       = Alignment(horizontal="center", vertical="center")
                data_file_ws.cell(CLASS_END, WRITE_COLUMN).border          = Border(top = Side(border_style="thin", color="909090"), bottom = Side(border_style="medium", color="000000"))
                
                if type(test_average) in (int, float):
                    data_file_ws.cell(CLASS_END, WRITE_COLUMN).fill = class_average_color(test_average)
            
            test_score   = form_ws.cell(i, TEST_SCORE_COLUMN).value
            student_name = form_ws.cell(i, DataForm.STUDENT_NAME_COLUMN).value

            if test_score is None:
                continue # 점수 없으면 미응시 처리

            # 학생 찾기
            for row in range(CLASS_START + 2, CLASS_END):
                if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == student_name:
                    data_file_ws.cell(row, WRITE_COLUMN).value = test_score
                    if type(test_score) in (int, float):
                        data_file_ws.cell(row, WRITE_COLUMN).fill = test_score_color(test_score)

                    data_file_ws.cell(row, WRITE_COLUMN).alignment = Alignment(horizontal="center", vertical="center")
                    break
            else:
                OmikronLog.warning(f"{sheet_name} 시트: {class_name} 반에 {student_name} 학생이 존재하지 않습니다.")

    data_file_ws = data_file_wb["데일리테스트"]
    save_to_temp(data_file_wb)

    # 조건부 서식 수식 로딩
    pythoncom.CoInitialize()
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(f"{os.getcwd()}\\data\\{DataFile.TEMP_FILE_NAME}.xlsx")
    wb.Save()
    wb.Close()
    excel.Quit()
    pythoncom.CoUninitialize()

    data_file_wb       = open_temp()
    data_file_color_wb = open_temp(True)

    for sheet_name in data_file_wb.sheetnames:
        if sheet_name not in ("데일리테스트", "모의고사"):
            continue

        data_file_ws       = data_file_wb[sheet_name]
        data_file_color_ws = data_file_color_wb[sheet_name]

        complete, _, _, _, _, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(data_file_ws)
        if not complete: return False, None

        for row in range(2, data_file_color_ws.max_row+1):
            if data_file_color_ws.cell(row, STUDENT_NAME_COLUMN).value is None:
                break

            # 학생 별 평균 점수에 대한 조건부 서식
            student_average = data_file_color_ws.cell(row, AVERAGE_SCORE_COLUMN).value
            if type(student_average) in (int, float):
                if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "시험 평균":
                    data_file_ws.cell(row, AVERAGE_SCORE_COLUMN).fill = class_average_color(student_average)
                else:
                    data_file_ws.cell(row, AVERAGE_SCORE_COLUMN).fill = student_average_color(student_average)

            # 신규생 하이라이트
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value in ("날짜", "시험명", "시험 평균"):
                continue
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).font.strike:
                continue
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).font.color is not None and data_file_ws.cell(row, STUDENT_NAME_COLUMN).font.color.rgb == "FFFF0000":
                continue

            complete, _, _, new_student = omikron.studentinfo.get_student_info(student_ws, data_file_ws.cell(row, STUDENT_NAME_COLUMN).value)
            if complete:
                if new_student:
                    data_file_ws.cell(row, STUDENT_NAME_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("FFFF00"))
                else:
                    data_file_ws.cell(row, STUDENT_NAME_COLUMN).fill = PatternFill(fill_type=None)
            else:
                data_file_ws.cell(row, STUDENT_NAME_COLUMN).fill = PatternFill(fill_type=None)
                OmikronLog.warning(f"{data_file_ws.cell(row, STUDENT_NAME_COLUMN).value} 학생 정보가 존재하지 않습니다.")

    data_file_ws = data_file_wb["데일리테스트"]
    close(data_file_color_wb)
    omikron.dataform.close(form_wb)

    return True, data_file_wb

def save_individual_test_data(target_row:int, target_col:int, test_score:int|float):
    # 백업 생성
    make_backup_file()

    data_file_wb = open()

    if "데일리테스트" not in data_file_wb.sheetnames:
        return False, None, None

    data_file_ws = data_file_wb["데일리테스트"]

    data_file_ws.cell(target_row, target_col).value = test_score
    data_file_ws.cell(target_row, target_col).fill = test_score_color(test_score)
    data_file_ws.cell(target_row, target_col).alignment = Alignment(horizontal="center", vertical="center")

    save_to_temp(data_file_wb)

    pythoncom.CoInitialize()
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(f"{os.getcwd()}\\data\\{DataFile.TEMP_FILE_NAME}.xlsx")
    wb.Save()
    wb.Close()
    excel.Quit()
    pythoncom.CoUninitialize()

    data_file_wb       = open_temp()
    data_file_color_wb = open_temp(True)

    data_file_ws       = data_file_wb["데일리테스트"]
    data_file_color_ws = data_file_color_wb["데일리테스트"]

    complete, _, _, _, _, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(data_file_ws)
    if not complete: return False, None, None

    student_average = data_file_color_ws.cell(target_row, AVERAGE_SCORE_COLUMN).value
    if type(student_average)  in (int, float):
        data_file_ws.cell(target_row, AVERAGE_SCORE_COLUMN).fill = student_average_color(student_average)

    test_average_row = target_row
    while data_file_color_ws.cell(test_average_row, STUDENT_NAME_COLUMN).value != "시험 평균":
        test_average_row += 1

    test_average = data_file_color_ws.cell(test_average_row, target_col).value
    if type(test_average) in (int, float):
        data_file_ws.cell(test_average_row, target_col).fill = class_average_color(test_average)

    class_average = data_file_color_ws.cell(test_average_row, AVERAGE_SCORE_COLUMN).value
    if type(class_average) in (int, float):
        data_file_ws.cell(test_average_row, target_col).fill = class_average_color(test_average)

    close(data_file_color_wb)

    return True, test_average, data_file_wb

def conditional_formatting():
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(f"{os.getcwd()}\\data\\{DATA_FILE_NAME}.xlsx")
    wb.Save()
    wb.Close()
    excel.Quit()
    pythoncom.CoUninitialize()

    data_file_wb         = open()
    data_file_color_wb   = open(True)
    student_wb           = omikron.studentinfo.open()
    complete, student_ws = omikron.studentinfo.open_worksheet(student_wb)
    if not complete: return False

    for sheet_name in data_file_wb.sheetnames:
        data_file_ws       = data_file_wb[sheet_name]
        data_file_color_ws = data_file_color_wb[sheet_name]

        complete, _, _, _, _, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(data_file_ws)
        if not complete: return False

        for row in range(2, data_file_ws.max_row+1):
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value is None:
                break
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "날짜":
                DATE_ROW = row
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value != "시험명":
                data_file_ws.row_dimensions[row].height = 18

            # 데이터 조건부 서식
            for col in range(1, data_file_color_ws.max_column+1):
                if col > AVERAGE_SCORE_COLUMN and data_file_ws.cell(DATE_ROW, col).value is None:
                    break

                data_file_ws.column_dimensions[gcl(col)].width = 14
                if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "날짜":
                    data_file_ws.cell(row, col).border = Border(top = Side(border_style="medium", color="000000"))
                elif data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "시험명":
                    data_file_ws.cell(row, col).border = Border(bottom = Side(border_style="thin", color="909090"))
                elif data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "시험 평균":
                    data_file_ws.cell(row, col).border = Border(top = Side(border_style="thin", color="909090"), bottom = Side(border_style="medium", color="000000"))
                else:
                    data_file_ws.cell(row, col).border = None

                # 학생 평균 점수 열 기준 분기   
                if col <= AVERAGE_SCORE_COLUMN:
                    continue

                if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "시험명":
                    data_file_ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center", wrapText=True)
                elif data_file_color_ws.cell(row, STUDENT_NAME_COLUMN).value == "시험 평균":
                    data_file_ws.cell(row, col).font = Font(bold=True)
                    if type(data_file_color_ws.cell(row, col).value) in (int, float):
                        data_file_ws.cell(row, col).fill = class_average_color(data_file_color_ws.cell(row, col).value)
                elif type(data_file_color_ws.cell(row, col).value) in (int, float):
                    data_file_ws.cell(row, col).fill = test_score_color(data_file_color_ws.cell(row, col).value)
                else:
                    data_file_ws.cell(row, col).fill = PatternFill(fill_type=None)

            # 학생별 평균 조건부 서식
            if type(data_file_color_ws.cell(row, AVERAGE_SCORE_COLUMN).value) in (int, float):
                if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "시험 평균":
                    data_file_ws.cell(row, AVERAGE_SCORE_COLUMN).fill = class_average_color(data_file_color_ws.cell(row, AVERAGE_SCORE_COLUMN).value)
                else:
                    data_file_ws.cell(row, AVERAGE_SCORE_COLUMN).fill = student_average_color(data_file_color_ws.cell(row, AVERAGE_SCORE_COLUMN).value)
            else:
                data_file_ws.cell(row, col).fill = PatternFill(fill_type=None)

            # 학생별 평균 폰트 설정
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value in ("날짜", "시험명", "시험 평균"):
                data_file_ws.cell(row, AVERAGE_SCORE_COLUMN).font = Font(bold=True)
                continue
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).font.strike:
                data_file_ws.cell(row, AVERAGE_SCORE_COLUMN).font = Font(bold=True, strike=True)
                continue
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).font.color is not None and data_file_ws.cell(row, STUDENT_NAME_COLUMN).font.color.rgb == "FFFF0000":
                data_file_ws.cell(row, AVERAGE_SCORE_COLUMN).font = Font(bold=True, color="FFFF0000")
                continue

            # 신규생 하이라이트
            complete, _, _, new_student = omikron.studentinfo.get_student_info(student_ws, data_file_ws.cell(row, STUDENT_NAME_COLUMN).value)
            if complete:
                if new_student:
                    data_file_ws.cell(row, STUDENT_NAME_COLUMN).fill = PatternFill(fill_type="solid", fgColor=Color("FFFF00"))
                else:
                    data_file_ws.cell(row, STUDENT_NAME_COLUMN).fill = PatternFill(fill_type=None)
            else:
                data_file_ws.cell(row, STUDENT_NAME_COLUMN).fill = PatternFill(fill_type=None)
                OmikronLog.warning(f"{data_file_ws.cell(row, STUDENT_NAME_COLUMN).value} 학생 정보가 존재하지 않습니다.")

    omikron.studentinfo.close(student_wb)
    close(data_file_color_wb)
    save(data_file_wb)
    return True

def add_student(student_name:str, target_class_name:str, data_file_wb:xl.Workbook=None):
    if data_file_wb is None:
        data_file_wb = open()

    for data_file_ws in data_file_wb.worksheets:
        complete, TEST_TIME_COLUMN, CLASS_WEEKDAY_COLUMN, CLASS_NAME_COLUMN, TEACHER_NAME_COLUMN, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(data_file_ws)
        if not complete: return False, None

        # 목표 반에 학생 추가
        for row in range(2, data_file_ws.max_row+1):
            if data_file_ws.cell(row, CLASS_NAME_COLUMN).value == target_class_name:
                class_index = row+2
                break
        else:
            OmikronLog.warning(f"{data_file_ws.title} 시트에 {target_class_name} 반이 존재하지 않습니다.")
            continue

        while data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).value != "시험 평균":
            if data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).value > student_name:
                break
            elif data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).font.strike:
                class_index += 1
            elif data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).font.color is not None and data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).font.color.rgb == "FFFF0000":
                class_index += 1
            elif data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).value == student_name:
                OmikronLog.error(f"{student_name} 학생이 이미 존재합니다.")
                return False, None
            else:
                class_index += 1

        data_file_ws.insert_rows(class_index)
        data_file_ws.cell(class_index, TEST_TIME_COLUMN).value         = data_file_ws.cell(class_index-1, TEST_TIME_COLUMN).value
        data_file_ws.cell(class_index, CLASS_WEEKDAY_COLUMN).value     = data_file_ws.cell(class_index-1, CLASS_WEEKDAY_COLUMN).value
        data_file_ws.cell(class_index, CLASS_NAME_COLUMN).value        = data_file_ws.cell(class_index-1, CLASS_NAME_COLUMN).value
        data_file_ws.cell(class_index, TEACHER_NAME_COLUMN).value      = data_file_ws.cell(class_index-1, TEACHER_NAME_COLUMN).value
        data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).value      = student_name

        data_file_ws.cell(class_index, TEST_TIME_COLUMN).alignment     = Alignment(horizontal="center", vertical="center")
        data_file_ws.cell(class_index, CLASS_WEEKDAY_COLUMN).alignment = Alignment(horizontal="center", vertical="center")
        data_file_ws.cell(class_index, CLASS_NAME_COLUMN).alignment    = Alignment(horizontal="center", vertical="center")
        data_file_ws.cell(class_index, TEACHER_NAME_COLUMN).alignment  = Alignment(horizontal="center", vertical="center")
        data_file_ws.cell(class_index, STUDENT_NAME_COLUMN).alignment  = Alignment(horizontal="center", vertical="center")

        data_file_ws.cell(class_index, AVERAGE_SCORE_COLUMN).alignment = Alignment(horizontal="center", vertical="center")
        data_file_ws.cell(class_index, AVERAGE_SCORE_COLUMN).font      = Font(bold=True)

    # save(data_file_wb)

    return rescoping_formula(data_file_wb)

def move_student(student_name:str, target_class_name:str, current_class_name:str):
    data_file_wb = open()

    for data_file_ws in data_file_wb.worksheets:
        complete, _, _, CLASS_NAME_COLUMN, _, STUDENT_NAME_COLUMN, _ = find_dynamic_columns(data_file_ws)
        if not complete: return False, None

        # 기존 반 데이터 빨간색 처리
        for row in range(2, data_file_ws.max_row+1):
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == student_name and data_file_ws.cell(row, CLASS_NAME_COLUMN).value == current_class_name:
                for col in range(1, data_file_ws.max_column+1):
                    if data_file_ws.cell(row, col).font.bold:
                        data_file_ws.cell(row, col).font = Font(bold=True, color="FFFF0000")
                    else:
                        data_file_ws.cell(row, col).font = Font(color="FFFF0000")
                break

    # save(data_file_wb)

    return add_student(student_name, target_class_name, data_file_wb)

def delete_student(student_name:str):
    """
    학생 퇴원 처리
    
    퇴원 처리된 학생은 모든 데이터에 취소선 적용
    """

    data_file_wb = open()
    for data_file_ws in data_file_wb.worksheets:
        complete, _, _, _, _, STUDENT_NAME_COLUMN, _ = find_dynamic_columns(data_file_ws)
        if not complete: return False, None

        for row in range(2, data_file_ws.max_row+1):
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == student_name:
                for col in range(1, data_file_ws.max_column+1):
                    data_file_ws.cell(row, col).font = Font(strike=True)
    
    # save(data_file_wb)

    return True, data_file_wb

def rescoping_formula(data_file_wb:xl.Workbook):
    """
    데이터 파일 내 평균 산출 수식의 범위 재조정
    """

    # data_file_wb = open()
    for data_file_ws in data_file_wb.worksheets:
        complete, _, _, _, _, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(data_file_ws)
        if not complete: return False, None

        # 평균 범위 재지정
        for row in range(2, data_file_ws.max_row+1):
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value is None:
                break
            striked = False
            colored = False
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).font.strike:
                striked = True
            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).font.color is not None:
                if data_file_ws.cell(row, STUDENT_NAME_COLUMN).font.color.rgb == "FFFF0000":
                    colored = True

            if data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "날짜":
                DATE_ROW = row
                CLASS_START = row+2
            elif data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "시험 평균":
                CLASS_END = row-1
                data_file_ws[f"{gcl(AVERAGE_SCORE_COLUMN)}{str(row)}"] = ArrayFormula(f"{gcl(AVERAGE_SCORE_COLUMN)}{str(row)}", f"=ROUND(AVERAGE(IFERROR({gcl(AVERAGE_SCORE_COLUMN)}{str(CLASS_START)}:{gcl(AVERAGE_SCORE_COLUMN)}{str(CLASS_END)}, \"\")), 0)")
                if CLASS_START >= CLASS_END:
                    continue
                for col in range(AVERAGE_SCORE_COLUMN+1, data_file_ws.max_column+1):
                    if data_file_ws.cell(DATE_ROW, col).value is None:
                        break
                    data_file_ws.cell(row, col).value = f"=ROUND(AVERAGE({gcl(col)}{str(CLASS_START)}:{gcl(col)}{str(CLASS_END)}), 0)"
                    data_file_ws.cell(row, col).font  = Font(bold=True)
            elif data_file_ws.cell(row, STUDENT_NAME_COLUMN).value == "시험명":
                continue
            else:
                data_file_ws.cell(row, AVERAGE_SCORE_COLUMN).value = f"=ROUND(AVERAGE({gcl(AVERAGE_SCORE_COLUMN+1)}{str(row)}:XFD{str(row)}), 0)"

            if striked:
                data_file_ws.cell(row, AVERAGE_SCORE_COLUMN).font = Font(bold=True, strike=True)
            elif colored:
                data_file_ws.cell(row, AVERAGE_SCORE_COLUMN).font = Font(bold=True, color="FFFF0000")
            else:
                data_file_ws.cell(row, AVERAGE_SCORE_COLUMN).font = Font(bold=True)

    return True, data_file_wb

def update_class():
    make_backup_file()

    complete, deleted_class_names, unregistered_class_names = omikron.classinfo.check_difference_between()
    if not complete: return False, None

    print(deleted_class_names, unregistered_class_names)
    # 조건부 서식 수식 로딩
    pythoncom.CoInitialize()
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(f"{os.getcwd()}\\data\\{DATA_FILE_NAME}.xlsx")
    wb.Save()
    wb.Close()
    excel.Quit()
    pythoncom.CoUninitialize()

    data_file_wb = open()

    if len(deleted_class_names) > 0:
        if not os.path.isfile("./data/지난 데이터.xlsx"):
            ini_wb = xl.Workbook()
            ini_ws = ini_wb.worksheets[0]
            ini_ws.title = "데일리테스트"
            ini_ws[gcl(DataFile.TEST_TIME_COLUMN)+"1"]     = "시간"
            ini_ws[gcl(DataFile.CLASS_WEEKDAY_COLUMN)+"1"] = "요일"
            ini_ws[gcl(DataFile.CLASS_NAME_COLUMN)+"1"]    = "반"
            ini_ws[gcl(DataFile.TEACHER_NAME_COLUMN)+"1"]  = "담당"
            ini_ws[gcl(DataFile.STUDENT_NAME_COLUMN)+"1"]  = "이름"
            ini_ws[gcl(DataFile.AVERAGE_SCORE_COLUMN)+"1"] = "학생 평균"
            ini_ws.freeze_panes    = f"{gcl(DataFile.DATA_COLUMN)}2"
            ini_ws.auto_filter.ref = f"A:{gcl(DataFile.MAX)}"

            for col in range(1, DataFile.DATA_COLUMN):
                ini_ws.cell(1, col).alignment = Alignment(horizontal="center", vertical="center")
                ini_ws.cell(1, col).border    = Border(bottom = Side(border_style="medium", color="000000"))
            
            # 모의고사 sheet 생성
            copy_ws                 = ini_wb.copy_worksheet(ini_wb["데일리테스트"])
            copy_ws.title           = "모의고사"
            copy_ws.freeze_panes    = f"{gcl(DataFile.DATA_COLUMN)}2"
            copy_ws.auto_filter.ref = f"A:{gcl(DataFile.MAX)}"

            ini_wb.save("./data/지난 데이터.xlsx")

        data_only_wb = open(True)
        post_data_wb = xl.load_workbook("./data/지난 데이터.xlsx")
        for sheet_name in data_file_wb.sheetnames:
            data_only_ws = data_only_wb[sheet_name]
            post_data_ws = post_data_wb[sheet_name]
            data_file_ws = data_file_wb[sheet_name]

            complete, TEST_TIME_COLUMN, CLASS_WEEKDAY_COLUMN, CLASS_NAME_COLUMN, TEACHER_NAME_COLUMN, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(data_file_ws)
            if not complete: return False, None

            for row in range(2, data_file_ws.max_row+1):
                while data_file_ws.cell(row, CLASS_NAME_COLUMN).value in deleted_class_names:
                    data_file_ws.delete_rows(row)
            
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
                        data_file_ws.column_dimensions[gcl(POST_DATA_WRITE_COLUMN)].width    = 14
                        POST_DATA_WRITE_COLUMN += 1
            
            data_file_ws.auto_filter.ref = f"A:{gcl(AVERAGE_SCORE_COLUMN)}"
        
        post_data_wb.save("./data/지난 데이터.xlsx")

    if len(unregistered_class_names) > 0:
        class_wb = omikron.classinfo.open_temp()
        complete, class_ws = omikron.classinfo.open_worksheet(class_wb)
        if not complete: return False

        class_student_dict = omikron.chrome.get_class_student_dict()

        for sheet_name in data_file_wb.sheetnames:
            post_data_ws = post_data_wb[sheet_name]
            data_file_ws = data_file_wb[sheet_name]

            complete, TEST_TIME_COLUMN, CLASS_WEEKDAY_COLUMN, CLASS_NAME_COLUMN, TEACHER_NAME_COLUMN, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(data_file_ws)
            if not complete: return False, None

            for row in range(data_file_ws.max_row+1, 1, -1):
                if data_file_ws.cell(row-1, DataFile.STUDENT_NAME_COLUMN).value is not None:
                    WRITE_RANGE = WRITE_LOCATION = row
                    break

            for class_name in unregistered_class_names:
                complete, teacher_name, class_weekday, test_time = omikron.classinfo.get_class_info(class_ws, class_name)
                if not complete: continue

                # 시험명
                data_file_ws.cell(WRITE_LOCATION, TEST_TIME_COLUMN).value     = test_time
                data_file_ws.cell(WRITE_LOCATION, CLASS_WEEKDAY_COLUMN).value = class_weekday
                data_file_ws.cell(WRITE_LOCATION, CLASS_NAME_COLUMN).value    = class_name
                data_file_ws.cell(WRITE_LOCATION, TEACHER_NAME_COLUMN).value  = teacher_name
                data_file_ws.cell(WRITE_LOCATION, STUDENT_NAME_COLUMN).value  = "날짜"
                WRITE_LOCATION += 1
                
                data_file_ws.cell(WRITE_LOCATION, TEST_TIME_COLUMN).value     = test_time
                data_file_ws.cell(WRITE_LOCATION, CLASS_WEEKDAY_COLUMN).value = class_weekday
                data_file_ws.cell(WRITE_LOCATION, CLASS_NAME_COLUMN).value    = class_name
                data_file_ws.cell(WRITE_LOCATION, TEACHER_NAME_COLUMN).value  = teacher_name
                data_file_ws.cell(WRITE_LOCATION, STUDENT_NAME_COLUMN).value  = "시험명"

                for col in range(1, AVERAGE_SCORE_COLUMN + 1):
                    data_file_ws.cell(WRITE_LOCATION, col).border = Border(bottom = Side(border_style="thin", color="909090"))

                WRITE_LOCATION += 1

                # 학생 루프
                for studnet_name in class_student_dict[class_name]:
                    data_file_ws.cell(WRITE_LOCATION, TEST_TIME_COLUMN).value     = test_time
                    data_file_ws.cell(WRITE_LOCATION, CLASS_WEEKDAY_COLUMN).value = class_weekday
                    data_file_ws.cell(WRITE_LOCATION, CLASS_NAME_COLUMN).value    = class_name
                    data_file_ws.cell(WRITE_LOCATION, TEACHER_NAME_COLUMN).value  = teacher_name
                    data_file_ws.cell(WRITE_LOCATION, STUDENT_NAME_COLUMN).value  = studnet_name
                    WRITE_LOCATION += 1
                
                # 시험별 평균
                data_file_ws.cell(WRITE_LOCATION, TEST_TIME_COLUMN).value     = test_time
                data_file_ws.cell(WRITE_LOCATION, CLASS_WEEKDAY_COLUMN).value = class_weekday
                data_file_ws.cell(WRITE_LOCATION, CLASS_NAME_COLUMN).value    = class_name
                data_file_ws.cell(WRITE_LOCATION, TEACHER_NAME_COLUMN).value  = teacher_name
                data_file_ws.cell(WRITE_LOCATION, STUDENT_NAME_COLUMN).value  = "시험 평균"

                for col in range(1, AVERAGE_SCORE_COLUMN+1):
                    data_file_ws.cell(WRITE_LOCATION, col).border = Border(top = Side(border_style="thin", color="909090"), bottom = Side(border_style="medium", color="000000"))

                WRITE_LOCATION += 1

            # 정렬
            for row in range(WRITE_RANGE, data_file_ws.max_row + 1):
                for col in range(1, AVERAGE_SCORE_COLUMN + 1):
                    data_file_ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")

            # 필터 범위 재지정
            data_file_ws.auto_filter.ref = f"A:{gcl(AVERAGE_SCORE_COLUMN)}"

    return rescoping_formula(data_file_wb)