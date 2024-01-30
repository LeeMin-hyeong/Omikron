import os.path
import openpyxl as xl

from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils.cell import get_column_letter as gcl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.datavalidation import DataValidation

import omikron.chrome

from omikron.defs import StudentInfo
from omikron.log import OmikronLog


def make_file() -> bool:
    ini_wb = xl.Workbook()
    ini_ws = ini_wb[ini_wb.sheetnames[0]]
    ini_ws.title = StudentInfo.DEFAULT_NAME

    ini_ws[gcl(StudentInfo.STUDENT_NAME_COLUMN)+"1"]       = "이름"
    # ini_ws[gcl(StudentInfo.CLASS_NAME_COLUMN)+"1"]       = "반명"
    # ini_ws[gcl(StudentInfo.TEACHER_NAME_COLUMN)+"1"]     = "담당"
    ini_ws[gcl(StudentInfo.MAKEUPTEST_WEEKDAY_COLUMN)+"1"] = "재시험 응시 요일"
    ini_ws[gcl(StudentInfo.MAKEUPTEST_TIME_COLUMN)+"1"]    = "재시험 응시 시간"
    ini_ws[gcl(StudentInfo.NEW_STUDENT_CHECK_COLUMN)+"1"]  = "기수 신규생"

    ini_ws["Z1"] = "N"
    ini_ws.auto_filter.ref = "A:"+gcl(StudentInfo.MAX)
    ini_ws.freeze_panes    = "A2"
    ini_ws.column_dimensions.group("Z", hidden=True)

    # 첫 행 정렬 및 자동 줄 바꿈
    for col in range(1, StudentInfo.MAX+1):
        ini_ws.cell(1, col).alignment = Alignment(horizontal="center", vertical="center", wrapText=True)
        ini_ws.cell(1, col).border    = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    # 학생 정보 불러오기 및 작성
    for studnet_name in omikron.chrome.get_student_names():
        WRITE_LOCATION = ini_ws.max_row + 1
        ini_ws.cell(WRITE_LOCATION, StudentInfo.STUDENT_NAME_COLUMN).value = studnet_name
        
        # 데이터 유효성 검사
        dv = DataValidation(type="list", formula1="=Z1", allow_blank=True, errorStyle="stop", showErrorMessage=True)
        dv.error = "이 셀의 값은 'N'이어야 합니다."
        ini_ws.add_data_validation(dv)
        dv.add(ini_ws.cell(WRITE_LOCATION, StudentInfo.NEW_STUDENT_CHECK_COLUMN))

    # 정렬 및 테두리
    for row in range(2, ini_ws.max_row + 1):
        for col in range(1, StudentInfo.MAX + 1):
            ini_ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")
            ini_ws.cell(row, col).border    = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    ini_wb.save(f"./{StudentInfo.DEFAULT_NAME}.xlsx")

    return True

def open(data_only:bool=False) -> xl.Workbook:
    return xl.load_workbook(f"./{StudentInfo.DEFAULT_NAME}.xlsx", data_only=data_only)

def open_worksheet(student_wb:xl.Workbook):
    try:
        return True, student_wb[StudentInfo.DEFAULT_NAME]
    except:
        OmikronLog.error(r"'학생 정보.xlsx'의 시트명을 '학생 정보'으로 변경해 주세요.")
        return False

def save(student_wb:xl.Workbook):
    student_wb.save(f"./{StudentInfo.DEFAULT_NAME}.xlsx")

def close(student_wb:xl.Workbook):
    student_wb.close()

def isopen() -> bool:
    return os.path.isfile(f"./data/~${StudentInfo.DEFAULT_NAME}.xlsx")

def get_student_info(student_ws:Worksheet, student_name:str):
    """
    학생 정보 파일로부터 학생 정보 추출

    return 파일 내 학생 존재 여부, 재시험 요일, 재시험 시간, 신규생 여부
    """

    for row in range(2, student_ws.max_row+1):
        if student_ws.cell(row, StudentInfo.STUDENT_NAME_COLUMN).value == student_name:
            makeup_test_weekday = student_ws.cell(row, StudentInfo.MAKEUPTEST_WEEKDAY_COLUMN).value
            makeup_test_time    = student_ws.cell(row, StudentInfo.MAKEUPTEST_TIME_COLUMN).value
            new_studnet         = student_ws.cell(row, StudentInfo.NEW_STUDENT_CHECK_COLUMN).value
            break
    else:
        return False, None, None, False
    
    return True, makeup_test_weekday, makeup_test_time, new_studnet == 'N'

def add_student(target_student_name:str) -> bool:
    """
    학생 정보 파일 내 신규생 추가

    return 작업 성공 여부
    """
    student_wb = open()
    complete, student_ws = open_worksheet(student_wb)
    if not complete:
        return False, None

    for row in range(student_ws.max_row+1, 1, -1):
        if student_ws.cell(row-1, StudentInfo.STUDENT_NAME_COLUMN).value is not None:
            student_ws.cell(row, StudentInfo.STUDENT_NAME_COLUMN).value      = target_student_name
            # student_ws.cell(row, StudentInfo.CLASS_NAME_COLUMN).value      = target_class_name
            student_ws.cell(row, StudentInfo.NEW_STUDENT_CHECK_COLUMN).value = "N"
            for col in range(1, StudentInfo.MAX+1):
                student_ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")
                student_ws.cell(row, col).border    = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
            break

    return True, student_wb

def delete_student(target_student_name:str):
    """
    학생 정보 파일에서 학생 정보 삭제
    """

    student_wb = open()
    complete, student_ws = open_worksheet(student_wb)
    if not complete: return False, None

    for row in range(2, student_ws.max_row+1):
        if student_ws.cell(row, StudentInfo.STUDENT_NAME_COLUMN).value == target_student_name:
            student_ws.delete_rows(row)

    return True, student_wb
