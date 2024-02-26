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
    wb = xl.Workbook()
    ws = wb.worksheets[0]
    ws.title = ClassInfo.DEFAULT_NAME
    ws[gcl(ClassInfo.CLASS_NAME_COLUMN)+"1"]    = "반명"
    ws[gcl(ClassInfo.TEACHER_NAME_COLUMN)+"1"]  = "선생님명"
    ws[gcl(ClassInfo.CLASS_WEEKDAY_COLUMN)+"1"] = "요일"
    ws[gcl(ClassInfo.TEST_TIME_COLUMN)+"1"]     = "시간"

    ws.freeze_panes = "A2"

    # 반 루프
    for class_name in omikron.chrome.get_class_names():
        WRITE_LOCATION = ws.max_row + 1
        ws.cell(WRITE_LOCATION, 1).value = class_name

    # 정렬 및 테두리
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row, col).border    = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    save(wb)

    return True

def open(data_only:bool=True) -> xl.Workbook:
    return xl.load_workbook(f"./{ClassInfo.DEFAULT_NAME}.xlsx", data_only=data_only)

def open_temp(data_only:bool=True) -> xl.Workbook:
    return xl.load_workbook(f"./{ClassInfo.TEMP_FILE_NAME}.xlsx", data_only=data_only)

def open_worksheet(wb:xl.Workbook):
    try:
        return True, wb[ClassInfo.DEFAULT_NAME]
    except:
        OmikronLog.error(r"'반 정보.xlsx'의 시트명을 '반 정보'로 변경해 주세요.")
        return False, None

def save(wb:xl.Workbook):
    wb.save(f"./{ClassInfo.DEFAULT_NAME}.xlsx")

def save_to_temp(wb:xl.Workbook):
    wb.save(f"./{ClassInfo.TEMP_FILE_NAME}.xlsx")

def delete_temp():
    os.remove(f"./{ClassInfo.TEMP_FILE_NAME}.xlsx")

def isopen() -> bool:
    return os.path.isfile(f"./data/~${ClassInfo.DEFAULT_NAME}.xlsx")

# 파일 유틸리티
def get_class_info(ws:Worksheet, class_name:str):
    """
    반 정보 파일로부터 특정 반의 정보 추출

    return `반 정보 존재 여부`, `담당 선생님`, `수업 요일`, `테스트 응시 시간`
    """
    for row in range(2, ws.max_row + 1):
        if ws.cell(row, ClassInfo.CLASS_NAME_COLUMN).value == class_name:
            teacher_name  = ws.cell(row, ClassInfo.TEACHER_NAME_COLUMN).value
            class_weekday = ws.cell(row, ClassInfo.CLASS_WEEKDAY_COLUMN).value
            test_time     = ws.cell(row, ClassInfo.TEST_TIME_COLUMN).value
            break
    else:
        return False, None, None, None
    
    return True, teacher_name, class_weekday, test_time

def get_class_names(ws:Worksheet) -> list[str]:
    """
    반 정보 기준 반 이름 리스트 추출
    """
    class_names = []
    for row in range(2, ws.max_row + 1):
        class_name = ws.cell(row, ClassInfo.CLASS_NAME_COLUMN).value
        if class_name is not None:
            class_names.append(class_name)

    return sorted(class_names)

def check_updated_class(ws:Worksheet):
    """
    아이소식에 존재하지만 반 정보 파일에 없는 반 목록 리턴
    """
    latest_class_names = omikron.chrome.get_class_names()
    class_names        = get_class_names(ws)

    unregistered_class_names = list(set(latest_class_names).difference(class_names))

    return unregistered_class_names

def check_difference_between():
    """
    반 정보 파일과 임시 반 정보 파일의 반 목록을 비교

    return `성공 여부`, `반 정보 파일에만 존재하는 반 리스트`, `임시 반 정보 파일에만 존재하는 반 리스트`
    """
    wb       = open()
    temp_wb  = open_temp()

    complete, ws = open_worksheet(wb)
    if not complete: return False, None, None

    complete, temp_ws = open_worksheet(temp_wb)
    if not complete: return False, None, None

    class_names        = get_class_names(ws)
    latest_class_names = get_class_names(temp_ws)

    deleted_class_names      = list(set(class_names).difference(latest_class_names))
    unregistered_class_names = list(set(latest_class_names).difference(class_names))

    return True, deleted_class_names, unregistered_class_names

# 파일 작업
def make_temp_file_for_update() -> bool:
    """
    반 업데이트 작업에 필요한 임시 반 정보 파일 생성

    등록되지 않은 반을 반 리스트 최하단에 작성
    """
    wb = open()
    complete, ws = open_worksheet(wb)
    if not complete: return False

    unregistered_class_names = check_updated_class(ws)
    if len(unregistered_class_names) == 0:
        save_to_temp(wb)
        return True

    for row in range(ws.max_row+1, 1, -1):
        if ws.cell(row-1, ClassInfo.CLASS_NAME_COLUMN).value is not None:
            WRITE_RANGE = WRITE_ROW = row
            break

    for row, class_name in enumerate(unregistered_class_names, start=WRITE_ROW):
        ws.cell(row, ClassInfo.CLASS_NAME_COLUMN).value = class_name

    for row in range(WRITE_RANGE, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row, col).border    = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    save_to_temp(wb)

    return True
