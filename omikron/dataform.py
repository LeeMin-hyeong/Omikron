import os
import openpyxl as xl

from datetime import datetime
from openpyxl.styles import Alignment, Border, Side, Protection
from openpyxl.utils.cell import get_column_letter as gcl
from openpyxl.worksheet.datavalidation import DataValidation

import omikron.chrome
import omikron.classinfo

from omikron.defs import DataForm
from omikron.log import OmikronLog

# 파일 기본 작업
def make_file() -> bool:
    """
    테스트 데이터 입력 양식 파일 생성
    """
    ini_wb = xl.Workbook()
    ini_ws = ini_wb.worksheets[0]
    ini_ws.title = DataForm.DEFAULT_NAME
    ini_ws[gcl(DataForm.CLASS_WEEKDAY_COLUMN)+"1"]     = "요일"
    ini_ws[gcl(DataForm.TEST_TIME_COLUMN)+"1"]         = "시간"
    ini_ws[gcl(DataForm.CLASS_NAME_COLUMN)+"1"]        = "반"
    ini_ws[gcl(DataForm.STUDENT_NAME_COLUMN)+"1"]      = "이름"
    ini_ws[gcl(DataForm.TEACHER_NAME_COLUMN)+"1"]      = "담당T"
    ini_ws[gcl(DataForm.DAILYTEST_NAME_COLUMN)+"1"]    = "시험명"
    ini_ws[gcl(DataForm.DAILYTEST_SCORE_COLUMN)+"1"]   = "점수"
    ini_ws[gcl(DataForm.DAILYTEST_AVERAGE_COLUMN)+"1"] = "평균"
    ini_ws[gcl(DataForm.MOCKTEST_NAME_COLUMN)+"1"]     = "모의고사 시험명"
    ini_ws[gcl(DataForm.MOCKTEST_SCORE_COLUMN)+"1"]    = "모의고사 점수"
    ini_ws[gcl(DataForm.MOCKTEST_AVERAGE_COLUMN)+"1"]  = "모의고사 평균"
    ini_ws[gcl(DataForm.MAKEUP_TEST_CHECK_COLUMN)+"1"] = "재시험 응시 여부"
    ini_ws["Y1"] = "X"
    ini_ws["Z1"] = "x"
    ini_ws.column_dimensions.group("Y", "Z", hidden=True)
    ini_ws.auto_filter.ref = "A:"+gcl(DataForm.TEST_TIME_COLUMN)
    ini_ws.freeze_panes    = "A2"
    
    for col in range(1, DataForm.MAX+1):
        ini_ws.cell(1, col).alignment = Alignment(horizontal="center", vertical="center", wrapText=True)
        ini_ws.cell(1, col).border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    class_wb = omikron.classinfo.open(True)
    complete, class_ws = omikron.classinfo.open_worksheet(class_wb)
    if not complete: return False

    for class_name, student_list in omikron.chrome.get_class_student_dict().items():
        if len(student_list) == 0:
            continue

        complete, teacher_name, class_weekday, test_time = omikron.classinfo.get_class_info(class_ws, class_name)
        if not complete: continue

        WRITE_LOCATION = start = ini_ws.max_row + 1

        ini_ws.cell(WRITE_LOCATION, DataForm.CLASS_NAME_COLUMN).value   = class_name
        ini_ws.cell(WRITE_LOCATION, DataForm.TEACHER_NAME_COLUMN).value = teacher_name

        #학생 루프
        for student_name in student_list:
            ini_ws.cell(WRITE_LOCATION, DataForm.CLASS_WEEKDAY_COLUMN).value = class_weekday
            ini_ws.cell(WRITE_LOCATION, DataForm.TEST_TIME_COLUMN).value     = test_time
            ini_ws.cell(WRITE_LOCATION, DataForm.STUDENT_NAME_COLUMN).value  = student_name
            dv = DataValidation(type="list", formula1="=Y1:Z1", showDropDown=True, allow_blank=True, showErrorMessage=True)
            dv.error = "이 셀의 값은 'x' 또는 'X'이어야 합니다."
            ini_ws.add_data_validation(dv)
            dv.add(ini_ws.cell(WRITE_LOCATION,DataForm.MAKEUP_TEST_CHECK_COLUMN))
            WRITE_LOCATION = ini_ws.max_row + 1
        
        end = WRITE_LOCATION - 1

        # 시험 평균
        ini_ws.cell(start, DataForm.DAILYTEST_AVERAGE_COLUMN).value = f"=ROUND(AVERAGE({gcl(DataForm.DAILYTEST_SCORE_COLUMN)+str(start)}:{gcl(DataForm.DAILYTEST_SCORE_COLUMN)+str(end)}), 0)"
        # 모의고사 평균
        ini_ws.cell(start, DataForm.MOCKTEST_AVERAGE_COLUMN).value = f"=ROUND(AVERAGE({gcl(DataForm.MOCKTEST_SCORE_COLUMN)+str(start)}:{gcl(DataForm.MOCKTEST_SCORE_COLUMN)+str(end)}), 0)"
        
        # 정렬 및 테두리
        for row in range(start, end + 1):
            for col in range(1, DataForm.MAX+1):
                ini_ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")
                ini_ws.cell(row, col).border    = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        
        # 셀 병합
        if start < end:
            ini_ws.merge_cells(f"{gcl(DataForm.CLASS_NAME_COLUMN)+str(start)}:{gcl(DataForm.CLASS_NAME_COLUMN)+str(end)}")
            ini_ws.merge_cells(f"{gcl(DataForm.TEACHER_NAME_COLUMN)+str(start)}:{gcl(DataForm.TEACHER_NAME_COLUMN)+str(end)}")
            ini_ws.merge_cells(f"{gcl(DataForm.DAILYTEST_NAME_COLUMN)+str(start)}:{gcl(DataForm.DAILYTEST_NAME_COLUMN)+str(end)}")
            ini_ws.merge_cells(f"{gcl(DataForm.DAILYTEST_AVERAGE_COLUMN)+str(start)}:{gcl(DataForm.DAILYTEST_AVERAGE_COLUMN)+str(end)}")
            ini_ws.merge_cells(f"{gcl(DataForm.MOCKTEST_NAME_COLUMN)+str(start)}:{gcl(DataForm.MOCKTEST_NAME_COLUMN)+str(end)}")
            ini_ws.merge_cells(f"{gcl(DataForm.MOCKTEST_AVERAGE_COLUMN)+str(start)}:{gcl(DataForm.MOCKTEST_AVERAGE_COLUMN)+str(end)}")
        
    ini_ws.protection.sheet         = True
    ini_ws.protection.autoFilter    = False
    ini_ws.protection.formatColumns = False
    for row in range(2, ini_ws.max_row + 1):
        ini_ws.cell(row, DataForm.CLASS_NAME_COLUMN).alignment         = Alignment(horizontal="center", vertical="center", wrapText=True)
        ini_ws.cell(row, DataForm.DAILYTEST_NAME_COLUMN).alignment     = Alignment(horizontal="center", vertical="center", wrapText=True)
        ini_ws.cell(row, DataForm.MOCKTEST_NAME_COLUMN).alignment      = Alignment(horizontal="center", vertical="center", wrapText=True)
        ini_ws.cell(row, DataForm.DAILYTEST_NAME_COLUMN).protection    = Protection(locked=False)
        ini_ws.cell(row, DataForm.DAILYTEST_SCORE_COLUMN).protection   = Protection(locked=False)
        ini_ws.cell(row, DataForm.MOCKTEST_NAME_COLUMN).protection     = Protection(locked=False)
        ini_ws.cell(row, DataForm.MOCKTEST_SCORE_COLUMN).protection    = Protection(locked=False)
        ini_ws.cell(row, DataForm.MAKEUP_TEST_CHECK_COLUMN).protection = Protection(locked=False)

    if os.path.isfile(f"./데일리테스트 기록 양식({datetime.today().strftime('%m.%d')}).xlsx"):
        i = 1
        while True:
            if not os.path.isfile(f"./데일리테스트 기록 양식({datetime.today().strftime('%m.%d')})({i}).xlsx"):
                ini_wb.save(f"./데일리테스트 기록 양식({datetime.today().strftime('%m.%d')})({i}).xlsx")
                break
            i += 1
    else:
        ini_wb.save(f"./데일리테스트 기록 양식({datetime.today().strftime('%m.%d')}).xlsx")

    return True

def open(filepath, data_only=True) -> xl.Workbook:
    return xl.load_workbook(filepath, data_only=data_only)

def open_worksheet(form_wb:xl.Workbook):
    try:
        return True, form_wb[DataForm.DEFAULT_NAME]
    except:
        OmikronLog.error(f"'{DataForm.DEFAULT_NAME}.xlsx'의 시트명을 '{DataForm.DEFAULT_NAME}'로 변경해 주세요.")
        return False, None

def close(form_wb:xl.Workbook):
    form_wb.close()

# 파일 유틸리티
def data_validation(filepath:str) -> bool:
    """
    데이터 입력 양식의 데이터가 올바르게 입력되었는지 확인
    """
    form_wb = open(filepath)
    complete, form_ws = open_worksheet(form_wb)
    if not complete: return False
    
    form_checked      = True
    dailytest_checked = False
    mocktest_checked  = False
    for i in range(1, form_ws.max_row+1):
        if form_ws.cell(i, DataForm.CLASS_NAME_COLUMN).value is not None:
            class_name        = form_ws.cell(i, DataForm.CLASS_NAME_COLUMN).value
            dailytest_checked = False
            mocktest_checked  = False
            dailytest_name    = form_ws.cell(i, DataForm.DAILYTEST_NAME_COLUMN).value
            mocktest_name     = form_ws.cell(i, DataForm.MOCKTEST_NAME_COLUMN).value
        
        if dailytest_checked and mocktest_checked: continue
        
        if not dailytest_checked and form_ws.cell(i, DataForm.DAILYTEST_SCORE_COLUMN).value is not None and dailytest_name is None:
            OmikronLog.error(f"{class_name}의 시험명이 작성되지 않았습니다.")
            dailytest_checked = True
            form_checked      = False
        if not mocktest_checked and form_ws.cell(i, DataForm.MOCKTEST_SCORE_COLUMN).value is not None and mocktest_name is None:
            OmikronLog.error(f"{class_name}의 모의고사명이 작성되지 않았습니다.")
            mocktest_checked = True
            form_checked     = False

    form_wb.close()

    return form_checked
