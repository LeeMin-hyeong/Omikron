"""Read-only queries over the main data workbook."""

from datetime import datetime

from openpyxl.worksheet.worksheet import Worksheet

import tdm.excel.class_info
from tdm.domain.models import DataFile
from tdm.excel.data_file_errors import NoReservedColumnError
from tdm.excel.data_file_storage import open

def get_data_sorted_dict(mocktest = False):
    """
    데이터 파일의 대략적 정보를 `dict` 형태로 추출

    return `dict[반:학생]`, `dict[반:시험명]`
    """
    wb = open()

    ws = wb[DataFile.DEFAULT_SHEET_NAME]

    CLASS_NAME_COLUMN, _, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN = find_dynamic_columns(ws)

    class_wb = tdm.excel.class_info.open()
    class_ws = tdm.excel.class_info.open_worksheet(class_wb)

    class_student_dict = {}
    class_test_dict    = {}

    for class_name in tdm.excel.class_info.get_class_names(class_ws, mocktest=mocktest):
        student_index_dict = {}
        test_index_dict    = {}
        for row in range(2, ws.max_row+1):
            if ws.cell(row, CLASS_NAME_COLUMN).value != class_name:
                continue
            if ws.cell(row, STUDENT_NAME_COLUMN).value == "날짜":
                for col in range(AVERAGE_SCORE_COLUMN+1, ws.max_column+1):
                    test_date = ws.cell(row, col).value
                    test_name = ws.cell(row+1, col).value
                    if test_date is None and test_name is None:
                        break
                    if type(test_date) == datetime:
                        test_date = test_date.strftime("%y.%m.%d")
                    else:
                        test_date = str(test_date).split()[0][2:10].replace("-", ".").replace(",", ".").replace("/", ".")
                    test_index_dict[f"[{test_date}] {test_name}"] = col
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

    return class_student_dict, class_test_dict

def find_dynamic_columns(ws:Worksheet):
    """
    파일 열(column) 정보 동적 탐색

    '반' 열, '담당' 열, '이름' 열, '학생 평균' 열

    return `CLASS_NAME_COLUMN`, `TEACHER_NAME_COLUMN`, `STUDENT_NAME_COLUMN`, `AVERAGE_SCORE_COLUMN`
    """

    for col in range(1, ws.max_column+1):
        if ws.cell(1, col).value == "반":
            CLASS_NAME_COLUMN = col
            break
    else:
        raise NoReservedColumnError(f"{ws.title} 시트에 '반' 열이 없습니다.")

    for col in range(1, ws.max_column+1):
        if ws.cell(1, col).value == "담당":
            TEACHER_NAME_COLUMN = col
            break
    else:
        raise NoReservedColumnError(f"{ws.title} 시트에 '담당' 열이 없습니다.")

    for col in range(1, ws.max_column+1):
        if ws.cell(1, col).value == "이름":
            STUDENT_NAME_COLUMN = col
            break
    else:
        raise NoReservedColumnError(f"{ws.title} 시트에 '이름' 열이 없습니다.")

    for col in range(1, ws.max_column+1):
        if ws.cell(1, col).value == "학생 평균":
            AVERAGE_SCORE_COLUMN = col
            break
    else:
        raise NoReservedColumnError(f"{ws.title} 시트에 '학생 평균' 열이 없습니다.")
    
    return CLASS_NAME_COLUMN, TEACHER_NAME_COLUMN, STUDENT_NAME_COLUMN, AVERAGE_SCORE_COLUMN

def is_cell_empty(row:int, col:int) -> bool:
    """
    데이터 파일이 열려있지 않을 때 특정 셀의 값이 비어있는 지 확인

    데일리테스트 시트 한정 기능
    """
    wb = open(data_only=True, read_only=True)
    ws = wb[DataFile.DEFAULT_SHEET_NAME]

    if ws.cell(row, col).value is None:
        return True, None

    value = ws.cell(row, col).value

    return False, value

def get_class_names(ws:Worksheet):
    class_names = []

    CLASS_NAME_COLUMN, _, _, _ = find_dynamic_columns(ws)

    for row in range(2, ws.max_row+1):
        if ws.cell(row, CLASS_NAME_COLUMN).value  not in class_names:
            class_names.append(ws.cell(row, CLASS_NAME_COLUMN).value)

    return class_names

def check_student_exist(student_name: str) -> bool:
    """데이터 파일의 어느 반에든 활성 상태인 학생이 있는지 확인."""
    wb = open(data_only=True, read_only=True)
    try:
        ws = wb[DataFile.DEFAULT_SHEET_NAME]
        _, _, STUDENT_NAME_COLUMN, _ = find_dynamic_columns(ws)

        for row in range(2, ws.max_row + 1):
            student_cell = ws.cell(row, STUDENT_NAME_COLUMN)
            if student_cell.value != student_name:
                continue
            if student_cell.font.strike:
                continue
            if student_cell.font.color is not None and student_cell.font.color.rgb == "FFFF0000":
                continue
            return True
        return False
    finally:
        wb.close()

# 파일 작업
