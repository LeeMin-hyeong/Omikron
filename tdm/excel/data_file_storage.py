"""Creation, opening, saving, and backup of the main data workbook."""

import os
import zipfile
from datetime import datetime

import openpyxl as xl
import pythoncom
import win32com.client
from openpyxl.utils.cell import get_column_letter as gcl
from openpyxl.worksheet.formula import ArrayFormula

import tdm.aisosik.reader
import tdm.config
import tdm.excel.class_info
from tdm.domain.errors import ExcelRequiredException, FileOpenException, NoMatchingSheetException, ReopenFileException
from tdm.domain.models import DataFile
from tdm.excel.styles import ALIGN_CENTER, BORDER_BOTTOM_MEDIUM_000, BORDER_BOTTOM_THIN_9090, BORDER_TOP_THIN_9090_BOTTOM_MEDIUM_000, FONT_BOLD

def _recalculate_with_excel(file_path: str):
    """Use Excel COM to force formula recalculation / conditional formatting evaluation."""
    pythoncom.CoInitialize()
    excel = None
    try:
        try:
            excel = win32com.client.Dispatch("Excel.Application")
        except pythoncom.com_error as e:
            raise ExcelRequiredException(
                "이 기능을 사용하려면 Microsoft Excel 이 설치되어 있어야 합니다."
            ) from e

        wb_com = excel.Workbooks.Open(os.path.abspath(file_path))
        wb_com.Save()
        wb_com.Close()
    finally:
        try:
            if excel is not None:
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

# 파일 기본 작업

def make_file():
    wb = xl.Workbook()
    ws = wb.worksheets[0]
    ws.title = DataFile.DEFAULT_SHEET_NAME
    ws[gcl(DataFile.CLASS_NAME_COLUMN)+"1"]    = "반"
    ws[gcl(DataFile.TEACHER_NAME_COLUMN)+"1"]  = "담당"
    ws[gcl(DataFile.STUDENT_NAME_COLUMN)+"1"]  = "이름"
    ws[gcl(DataFile.AVERAGE_SCORE_COLUMN)+"1"] = "학생 평균"
    ws.freeze_panes    = f"{gcl(DataFile.DATA_COLUMN)}2"
    ws.auto_filter.ref = f"A:{gcl(DataFile.MAX)}"

    for col in range(1, DataFile.DATA_COLUMN):
        ws.cell(1, col).border = BORDER_BOTTOM_MEDIUM_000

    class_wb = tdm.excel.class_info.open(True)
    class_ws = tdm.excel.class_info.open_worksheet(class_wb)

    # 반 루프
    for class_name, student_list in tdm.aisosik.reader.get_class_student_dict().items():
        if len(student_list) == 0:
            continue

        exist, teacher_name, _, _, mock_test_check = tdm.excel.class_info.get_class_info(class_name, ws=class_ws)
        if not exist: continue

        for i in range(2):
            if i == 1 and not mock_test_check:
                continue

            if i == 1:
                class_name = class_name + " (모의고사)"

            WRITE_LOCATION = ws.max_row + 1

            # 시험명
            ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value    = class_name
            ws.cell(WRITE_LOCATION, DataFile.TEACHER_NAME_COLUMN).value  = teacher_name
            ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value  = "날짜"
            
            WRITE_LOCATION = ws.max_row + 1
            ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value    = class_name
            ws.cell(WRITE_LOCATION, DataFile.TEACHER_NAME_COLUMN).value  = teacher_name
            ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value  = "시험명"

            for col in range(1, DataFile.DATA_COLUMN):
                ws.cell(WRITE_LOCATION, col).border = BORDER_BOTTOM_THIN_9090

            class_start = WRITE_LOCATION + 1

            # 학생 루프
            for student_name in student_list:
                WRITE_LOCATION = ws.max_row + 1
                ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value    = class_name
                ws.cell(WRITE_LOCATION, DataFile.TEACHER_NAME_COLUMN).value  = teacher_name
                ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value  = student_name
                ws.cell(WRITE_LOCATION, DataFile.AVERAGE_SCORE_COLUMN).value = f"=ROUND(AVERAGE({gcl(DataFile.DATA_COLUMN)}{WRITE_LOCATION}:XFD{WRITE_LOCATION}), 0)"
                ws.cell(WRITE_LOCATION, DataFile.AVERAGE_SCORE_COLUMN).font  = FONT_BOLD
            
            # 시험별 평균
            class_end = WRITE_LOCATION
            WRITE_LOCATION = ws.max_row + 1
            ws.cell(WRITE_LOCATION, DataFile.CLASS_NAME_COLUMN).value    = class_name
            ws.cell(WRITE_LOCATION, DataFile.TEACHER_NAME_COLUMN).value  = teacher_name
            ws.cell(WRITE_LOCATION, DataFile.STUDENT_NAME_COLUMN).value  = "시험 평균"
            ws[f"{gcl(DataFile.AVERAGE_SCORE_COLUMN)}{WRITE_LOCATION}"] = ArrayFormula(
                f"{gcl(DataFile.AVERAGE_SCORE_COLUMN)}{WRITE_LOCATION}",
                f"=ROUND(SUM(IFERROR({gcl(DataFile.AVERAGE_SCORE_COLUMN)}{class_start}:{gcl(DataFile.AVERAGE_SCORE_COLUMN)}{class_end},0))/COUNT({gcl(DataFile.AVERAGE_SCORE_COLUMN)}{class_start}:{gcl(DataFile.AVERAGE_SCORE_COLUMN)}{class_end}),0)",
            )
            ws.cell(WRITE_LOCATION, DataFile.AVERAGE_SCORE_COLUMN).font = FONT_BOLD

            for col in range(1, DataFile.DATA_COLUMN):
                ws.cell(WRITE_LOCATION, col).border = BORDER_TOP_THIN_9090_BOTTOM_MEDIUM_000

    # 정렬
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            ws.cell(row, col).alignment = ALIGN_CENTER

    save(wb)

def open(data_only:bool=False, read_only:bool=False) -> xl.Workbook:
    try:
        return xl.load_workbook(f"{tdm.config.DATA_DIR}/data/{tdm.config.DATA_FILE_NAME}.xlsx", data_only=data_only, read_only=read_only)
    except PermissionError:
        raise ReopenFileException(f"{tdm.config.DATA_FILE_NAME} 파일에 접근할 수 없습니다.\n파일을 직접 연 후 닫으면 문제가 해결될 수 있습니다.")
    except zipfile.BadZipFile:
        raise ReopenFileException(f"{tdm.config.DATA_FILE_NAME} 파일을 직접 연 후 닫으면 문제가 해결될 수 있습니다.")

def open_temp(data_only:bool=False, read_only:bool=False) -> xl.Workbook:
    return xl.load_workbook(f"{tdm.config.DATA_DIR}/data/{DataFile.TEMP_FILE_NAME}.xlsx", data_only=data_only, read_only=read_only)

def save(wb:xl.Workbook):
    try:
        if not os.path.isdir(f"{tdm.config.DATA_DIR}/data"):
            os.mkdir(f"{tdm.config.DATA_DIR}/data")
        wb.save(f"{tdm.config.DATA_DIR}/data/{tdm.config.DATA_FILE_NAME}.xlsx")
    except:
        raise FileOpenException(f"{tdm.config.DATA_FILE_NAME} 파일을 닫은 뒤 다시 시도해주세요")

def save_to_temp(wb:xl.Workbook):
    if not os.path.isdir(f"{tdm.config.DATA_DIR}/data"):
        os.mkdir(f"{tdm.config.DATA_DIR}/data")
    wb.save(f"{tdm.config.DATA_DIR}/data/{DataFile.TEMP_FILE_NAME}.xlsx")
    os.system(f"attrib +h {tdm.config.DATA_DIR}/data/{DataFile.TEMP_FILE_NAME}.xlsx")

def delete_temp():
    try:
        os.remove(f"{tdm.config.DATA_DIR}/data/{DataFile.TEMP_FILE_NAME}.xlsx")
    except:
        pass

def file_validation():
    wb = open(read_only=True)

    if DataFile.DEFAULT_SHEET_NAME not in wb.sheetnames:
        raise NoMatchingSheetException(f"데이터 파일: {DataFile.DEFAULT_SHEET_NAME} 시트가 존재하지 않습니다.")

    wb.close()

# 파일 유틸리티

def make_backup_file():
    if not os.path.isdir(f"{tdm.config.DATA_DIR}/data"):
        os.mkdir(f"{tdm.config.DATA_DIR}/data")
    if not os.path.isdir(f"{tdm.config.DATA_DIR}/data/backup"):
        os.mkdir(f"{tdm.config.DATA_DIR}/data/backup")
    wb = open()
    wb.save(f"{tdm.config.DATA_DIR}/data/backup/{tdm.config.DATA_FILE_NAME}({datetime.today().strftime('%Y%m%d%H%M%S')}).xlsx")
