import os
import openpyxl as xl

from openpyxl.utils.cell import get_column_letter as gcl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment, Border, Side

import omikron.chrome

from omikron.defs import ClassInfo
from omikron.log import OmikronLog

# 파일 기본 작업
def make_file() -> bool:
    ini_wb = xl.Workbook()
    ini_ws = ini_wb.worksheets[0]
    ini_ws.title = ClassInfo.DEFAULT_NAME
    ini_ws[gcl(ClassInfo.CLASS_NAME_COLUMN)+"1"]    = "반명"
    ini_ws[gcl(ClassInfo.TEACHER_NAME_COLUMN)+"1"]  = "선생님명"
    ini_ws[gcl(ClassInfo.CLASS_WEEKDAY_COLUMN)+"1"] = "요일"
    ini_ws[gcl(ClassInfo.TEST_TIME_COLUMN)+"1"]     = "시간"

    ini_ws.freeze_panes = "A2"

    # 반 루프
    for class_name in omikron.chrome.get_class_names():
        WRITE_LOCATION = ini_ws.max_row + 1
        ini_ws.cell(WRITE_LOCATION, 1).value = class_name

    # 정렬 및 테두리
    for row in range(1, ini_ws.max_row + 1):
        for col in range(1, ini_ws.max_column + 1):
            ini_ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")
            ini_ws.cell(row, col).border    = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    ini_wb.save(f"./{ClassInfo.DEFAULT_NAME}.xlsx")

    return True

def open(data_only:bool=True) -> xl.Workbook:
    return xl.load_workbook(f"./{ClassInfo.DEFAULT_NAME}.xlsx", data_only=data_only)

def open_temp(data_only:bool=True) -> xl.Workbook:
    return xl.load_workbook(f"./{ClassInfo.TEMP_FILE_NAME}.xlsx", data_only=data_only)

def open_worksheet(class_wb:xl.Workbook):
    try:
        return True, class_wb[ClassInfo.DEFAULT_NAME]
    except:
        OmikronLog.error(r"'반 정보.xlsx'의 시트명을 '반 정보'로 변경해 주세요.")
        return False, None

def save(class_wb:xl.Workbook):
    class_wb.save(f"./{ClassInfo.DEFAULT_NAME}.xlsx")
    class_wb.close()

def save_to_temp(class_wb:xl.Workbook):
    class_wb.save(f"./{ClassInfo.TEMP_FILE_NAME}.xlsx")
    class_wb.close()

def delete_temp():
    os.remove(f"./{ClassInfo.TEMP_FILE_NAME}.xlsx")

def close(class_wb:xl.Workbook):
    class_wb.close()

def isopen() -> bool:
    return os.path.isfile(f"./data/~${ClassInfo.DEFAULT_NAME}.xlsx")

# 파일 유틸리티
def get_class_info(class_ws:Worksheet, class_name:str):
    """
    반 정보 파일로부터 특정 반의 정보 추출

    return 존재 여부, 담당 선생님, 수업 요일, 테스트 응시 시간
    """
    for row in range(2, class_ws.max_row + 1):
        if class_ws.cell(row, ClassInfo.CLASS_NAME_COLUMN).value == class_name:
            teacher_name  = class_ws.cell(row, ClassInfo.TEACHER_NAME_COLUMN).value
            class_weekday = class_ws.cell(row, ClassInfo.CLASS_WEEKDAY_COLUMN).value
            test_time     = class_ws.cell(row, ClassInfo.TEST_TIME_COLUMN).value
            break
    else:
        return False, None, None, None
    
    return True, teacher_name, class_weekday, test_time

def get_class_names(class_ws:Worksheet) -> list[str]:
    """
    반 정보 기준 반 이름 리스트 추출
    """
    class_names = []
    for row in range(2, class_ws.max_row + 1):
        class_name = class_ws.cell(row, ClassInfo.CLASS_NAME_COLUMN).value
        if class_name is not None:
            class_names.append(class_name)

    return sorted(class_names)

def check_updated_class(class_ws:Worksheet):
    latest_class_names = omikron.chrome.get_class_names()
    class_names        = get_class_names(class_ws)

    unregistered_class_names = list(set(latest_class_names).difference(class_names))

    if len(unregistered_class_names) == 0:
        return False

    return True, unregistered_class_names

def check_difference_between():
    class_wb = open()
    temp_wb  = open_temp()

    complete, class_ws = open_worksheet(class_wb)
    if not complete: return False, None, None

    complete, temp_ws = open_worksheet(temp_wb)
    if not complete: return False, None, None

    class_names        = get_class_names(class_ws)
    latest_class_names = get_class_names(temp_ws)

    deleted_class_names      = list(set(class_names).difference(latest_class_names))
    unregistered_class_names = list(set(latest_class_names).difference(class_names))

    close(class_wb)
    close(temp_wb)

    return True, deleted_class_names, unregistered_class_names

# 파일 작업
def make_temp_file_for_update():
    class_wb = open()
    complete, class_ws = open_worksheet(class_wb)
    if not complete: return False

    complete, unregistered_class_names = check_updated_class(class_ws)
    if not complete: return False

    for row in range(class_ws.max_row+1, 1, -1):
        if class_ws.cell(row-1, ClassInfo.CLASS_NAME_COLUMN).value is not None:
            WRITE_RANGE = WRITE_ROW = row
            break

    for row, class_name in enumerate(unregistered_class_names, start=WRITE_ROW):
        class_ws.cell(row, ClassInfo.CLASS_NAME_COLUMN).value = class_name

    for row in range(WRITE_RANGE, class_ws.max_row + 1):
        for col in range(1, class_ws.max_column + 1):
            class_ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")
            class_ws.cell(row, col).border    = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    save_to_temp(class_wb)

    return True
