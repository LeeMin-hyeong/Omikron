import json
import os.path
import calendar
import tkinter as tk
import openpyxl as xl
import win32com.client # only works in Windows

from tkinter import filedialog
from datetime import date as DATE, datetime, timedelta
from dateutil.relativedelta import relativedelta
from openpyxl.utils.cell import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment, Border, Color, PatternFill, Side, Font
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from win32process import CREATE_NO_WINDOW
from webdriver_manager.chrome import ChromeDriverManager

config = json.load(open('config.json', encoding='UTF8'))
os.environ['WDM_PROGRESS_BAR'] = '0'
service = Service(ChromeDriverManager().install())
service.creation_flags = CREATE_NO_WINDOW

def makeDataFile(gui):
    gui.makeDataFileButton['state'] = tk.DISABLED
    if not os.path.isfile('./data/' + config['dataFileName'] + '.xlsx'):
        gui.appendLog('데이터파일 생성 중...')

        iniWb = xl.Workbook()
        iniWs = iniWb.active
        iniWs.title = 'DailyTest'
        iniWs['A1'] = '시간'
        iniWs['B1'] = '요일'
        iniWs['C1'] = '반'
        iniWs['D1'] = '담당'
        iniWs['E1'] = '이름'
        iniWs['F1'] = '학생 평균'
        iniWs.freeze_panes = 'G2'
        iniWs.auto_filter.ref = 'A:E'

        # 반 정보 확인
        if not os.path.isfile('./반 정보.xlsx'):
            gui.appendLog('[오류] 반 정보.xlsx 파일이 존재하지 않습니다.')
            return
        classWb = xl.load_workbook("./반 정보.xlsx")
        try:
            classWs = classWb['반 정보']
        except:
            gui.appendLog('[오류] \'반 정보.xlsx\'의 시트명을')
            gui.appendLog('\'반 정보\'로 변경해 주세요.')
            gui.makeDataFileButton['state'] = tk.NORMAL
            return

        options = webdriver.ChromeOptions()
        options.add_argument('headless')
        driver = webdriver.Chrome(service = service, options = options)

        # 아이소식 접속
        driver.get(config['url'])
        tableNames = driver.find_elements(By.CLASS_NAME, 'style1')

        # 반 루프
        for i in range(3, len(tableNames)):
            trs = driver.find_element(By.ID, 'table_' + str(i)).find_elements(By.CLASS_NAME, 'style12')
            writeLocation = iniWs.max_row + 1

            className = tableNames[i].text.rstrip()
            time = ''
            date = ''
            teacher = ''
            isClassExist = False
            for j in range(2, classWs.max_row + 1):
                if classWs.cell(j, 1).value == className:
                    teacher = classWs.cell(j, 2).value
                    date = classWs.cell(j, 3).value
                    time = classWs.cell(j, 4).value
                    isClassExist = True
            if not isClassExist:
                continue
            
            # 시험명
            iniWs.cell(writeLocation, 1).value = time
            iniWs.cell(writeLocation, 2).value = date
            iniWs.cell(writeLocation, 3).value = className
            iniWs.cell(writeLocation, 4).value = teacher
            iniWs.cell(writeLocation, 5).value = '날짜'
            
            writeLocation = iniWs.max_row + 1
            iniWs.cell(writeLocation, 1).value = time
            iniWs.cell(writeLocation, 2).value = date
            iniWs.cell(writeLocation, 3).value = className
            iniWs.cell(writeLocation, 4).value = teacher
            iniWs.cell(writeLocation, 5).value = '시험명'
            start = writeLocation + 1

            # 학생 루프
            for tr in trs:
                writeLocation = iniWs.max_row + 1
                iniWs.cell(writeLocation, 1).value = time
                iniWs.cell(writeLocation, 2).value = date
                iniWs.cell(writeLocation, 3).value = className
                iniWs.cell(writeLocation, 4).value = teacher
                iniWs.cell(writeLocation, 5).value = tr.find_element(By.CLASS_NAME, 'style9').text
                iniWs.cell(writeLocation, 6).value = '=ROUND(AVERAGE(G' + str(writeLocation) + ':XFD' + str(writeLocation) + '), 0)'
                iniWs.cell(writeLocation, 6).font = Font(bold=True)
            
            # 시험별 평균
            writeLocation = iniWs.max_row + 1
            end = writeLocation - 1
            iniWs.cell(writeLocation, 1).value = time
            iniWs.cell(writeLocation, 2).value = date
            iniWs.cell(writeLocation, 3).value = className
            iniWs.cell(writeLocation, 4).value = teacher
            iniWs.cell(writeLocation, 5).value = '시험 평균'
            iniWs.cell(writeLocation, 6).value = '=ROUND(AVERAGE(F' + str(start) + ':F' + str(end) + '), 0)'
            iniWs.cell(writeLocation, 6).font = Font(bold=True)

            for j in range(1, 7):
                iniWs.cell(writeLocation, j).border = Border(bottom = Side(border_style='medium', color='000000'))

        # 정렬
        for i in range(1, iniWs.max_row + 1):
                for j in range(1, iniWs.max_column + 1):
                    iniWs.cell(i, j).alignment = Alignment(horizontal='center', vertical='center')
        
        # 모의고사 sheet 생성
        copyWs = iniWb.copy_worksheet(iniWb['DailyTest'])
        copyWs.title = '모의고사'
        copyWs.freeze_panes = 'G2'
        copyWs.auto_filter.ref = 'A:E'

        iniWb.save('./data/' + config['dataFileName'] + '.xlsx')
        gui.appendLog('데이터 파일을 생성했습니다.')
        gui.ui.wm_attributes("-topmost", 1)
        gui.ui.wm_attributes("-topmost", 0)

    else:
        gui.appendLog('이미 파일이 존재합니다')

def makeDataForm(gui):
    gui.appendLog('데일리테스트 기록 양식 생성 중...')

    iniWb = xl.Workbook()
    iniWs = iniWb.active
    iniWs.title = '데일리테스트 기록 양식'
    iniWs['A1'] = '요일'
    iniWs['B1'] = '시간'
    iniWs['C1'] = '반'
    iniWs['D1'] = '이름'
    iniWs['E1'] = '담당T'
    iniWs['F1'] = '시험명'
    iniWs['G1'] = '점수'
    iniWs['H1'] = '평균'
    iniWs['I1'] = '시험대비 모의고사명'
    iniWs['J1'] = '모의고사 점수'
    iniWs['K1'] = '모의고사 평균'
    iniWs['L1'] = '재시문자 X'
    iniWs['M1'] = 'X'
    iniWs['N1'] = 'x'
    iniWs.column_dimensions.group('M', 'N', hidden=True)
    iniWs.auto_filter.ref = 'A:B'
    if not os.path.isfile('./반 정보.xlsx'):
        gui.appendLog('[오류] 반 정보.xlsx 파일이 존재하지 않습니다.')
        return
    
    classWb = xl.load_workbook("반 정보.xlsx")
    try:
        classWs = classWb['반 정보']
    except:
        gui.appendLog('[오류] \'반 정보.xlsx\'의 시트명을')
        gui.appendLog('\'반 정보\'로 변경해 주세요.')
        return

    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    driver = webdriver.Chrome(service = service, options = options)

    # 아이소식 접속
    driver.get(config['url'])
    tableNames = driver.find_elements(By.CLASS_NAME, 'style1')
    #반 루프
    for i in range(3, len(tableNames)):
        trs = driver.find_element(By.ID, 'table_' + str(i)).find_elements(By.CLASS_NAME, 'style12')
        writeLocation = start = iniWs.max_row + 1

        className = tableNames[i].text.rstrip()
        teacher = ''
        date = ''
        time = ''
        isClassExist = False

        for j in range(2, classWs.max_row + 1):
            if classWs.cell(j, 1).value == className:
                teacher = classWs.cell(j, 2).value
                date = classWs.cell(j, 3).value
                time = classWs.cell(j, 4).value
                isClassExist = True
        if not isClassExist:
            continue
        iniWs.cell(writeLocation, 3).value = className
        iniWs.cell(writeLocation, 5).value = teacher

        #학생 루프
        for tr in trs:
            iniWs.cell(writeLocation, 1).value = date
            iniWs.cell(writeLocation, 2).value = time
            iniWs.cell(writeLocation, 4).value = tr.find_element(By.CLASS_NAME, 'style9').text
            dv = DataValidation(type='list', formula1='=M1:N1', showDropDown=True, allow_blank=True, showErrorMessage=True)
            iniWs.add_data_validation(dv)
            dv.add(iniWs.cell(writeLocation, 12))
            writeLocation = iniWs.max_row + 1
        
        end = writeLocation - 1

        # 시험 평균
        iniWs.cell(start, 8).value = '=ROUND(AVERAGE(G' + str(start) + ':G' + str(end) + '), 0)'
        # 모의고사 평균
        iniWs.cell(start, 11).value = '=ROUND(AVERAGE(J' + str(start) + ':J' + str(end) + '), 0)'
        
        # 정렬 및 테두리
        for j in range(1, iniWs.max_row + 1):
            for k in range(1, iniWs.max_column + 1):
                iniWs.cell(j, k).alignment = Alignment(horizontal='center', vertical='center')
                iniWs.cell(j, k).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        if start < end:
            # 셀 병합
            iniWs.merge_cells('C' + str(start) + ':C' + str(end))
            iniWs.merge_cells('E' + str(start) + ':E' + str(end))
            iniWs.merge_cells('F' + str(start) + ':F' + str(end))
            iniWs.merge_cells('H' + str(start) + ':H' + str(end))
            iniWs.merge_cells('I' + str(start) + ':I' + str(end))
            iniWs.merge_cells('K' + str(start) + ':K' + str(end))
        
    if os.path.isfile('./데일리테스트 기록 양식.xlsx'):
        i = 1
        while True:
            if not os.path.isfile('./데일리테스트 기록 양식(' + str(i) +').xlsx'):
                iniWb.save('./데일리테스트 기록 양식(' + str(i) +').xlsx')
                break;
            i += 1
    else:
        iniWb.save('./데일리테스트 기록 양식.xlsx')
    gui.appendLog('데일리테스트 기록 양식 생성을 완료했습니다.')
    gui.ui.wm_attributes("-topmost", 1)
    gui.ui.wm_attributes("-topmost", 0)

def saveData(gui):
    def quitEvent():
        window.destroy()
        gui.saveDataButton['state'] = tk.NORMAL
        gui.ui.wm_attributes("-disabled", False)
        gui.ui.wm_attributes("-topmost", 1)
        gui.ui.wm_attributes("-topmost", 0)
    gui.saveDataButton['state'] = tk.DISABLED
    gui.ui.wm_attributes("-disabled", True)
    window=tk.Tk()
    window.geometry('200x300+500+500')
    window.resizable(False, False)
    window.protocol("WM_DELETE_WINDOW", quitEvent)

    today = DATE.today()

    makeupDate={}
    if today == today + relativedelta(weekday=calendar.MONDAY):
        makeupDate['월'] = today + timedelta(days=7)
    else:
        makeupDate['월'] = today + relativedelta(weekday=calendar.MONDAY)

    if today == today + relativedelta(weekday=calendar.TUESDAY):
        makeupDate['화'] = today + timedelta(days=7)
    else:
        makeupDate['화'] = today + relativedelta(weekday=calendar.TUESDAY)

    if today == today + relativedelta(weekday=calendar.WEDNESDAY):
        makeupDate['수'] = today + timedelta(days=7)
    else:
        makeupDate['수'] = today + relativedelta(weekday=calendar.WEDNESDAY)

    if today == today + relativedelta(weekday=calendar.THURSDAY):
        makeupDate['목'] = today + timedelta(days=7)
    else:
        makeupDate['목'] = today + relativedelta(weekday=calendar.THURSDAY)

    if today == today + relativedelta(weekday=calendar.FRIDAY):
        makeupDate['금'] = today + timedelta(days=7)
    else:
        makeupDate['금'] = today + relativedelta(weekday=calendar.FRIDAY)

    if today == today + relativedelta(weekday=calendar.SATURDAY):
        makeupDate['토'] = today + timedelta(days=7)
    else:
        makeupDate['토'] = today + relativedelta(weekday=calendar.SATURDAY)

    if today == today + relativedelta(weekday=calendar.SUNDAY):
        makeupDate['일'] = today + timedelta(days=7)
    else:
        makeupDate['일'] = today + relativedelta(weekday=calendar.SUNDAY)

    mon = tk.IntVar()
    tue = tk.IntVar()
    wed = tk.IntVar()
    thu = tk.IntVar()
    fri = tk.IntVar()
    sat = tk.IntVar()
    sun = tk.IntVar()

    tk.Label(window, text='\n다음 중 휴일을 선택해주세요\n').pack()
    dateCalc = DATE.today() + timedelta(days=1)
    for i in range(0, 8):
        for j in range(0, 8):
            if dateCalc == makeupDate['월']:
                tk.Checkbutton(window, text=str(makeupDate['월'])+' (월)', variable=mon).pack()
                dateCalc += timedelta(days=1)
        for j in range(0, 8):
            if dateCalc == makeupDate['화']:
                tk.Checkbutton(window, text=str(makeupDate['화'])+' (화)', variable=tue).pack()
                dateCalc += timedelta(days=1)
        for j in range(0, 8):
            if dateCalc == makeupDate['수']:
                tk.Checkbutton(window, text=str(makeupDate['수'])+' (수)', variable=wed).pack()
                dateCalc += timedelta(days=1)
        for j in range(0, 8):
            if dateCalc == makeupDate['목']:
                tk.Checkbutton(window, text=str(makeupDate['목'])+' (목)', variable=thu).pack()
                dateCalc += timedelta(days=1)
        for j in range(0, 8):
            if dateCalc == makeupDate['금']:
                tk.Checkbutton(window, text=str(makeupDate['금'])+' (금)', variable=fri).pack()
                dateCalc += timedelta(days=1)
        for j in range(0, 8):
            if dateCalc == makeupDate['토']:
                tk.Checkbutton(window, text=str(makeupDate['토'])+' (토)', variable=sat).pack()
                dateCalc += timedelta(days=1)
        for j in range(0, 8):
            if dateCalc == makeupDate['일']:
                tk.Checkbutton(window, text=str(makeupDate['일'])+' (일)', variable=sun).pack()
                dateCalc += timedelta(days=1)
    tk.Label(window, text='\n').pack()
    tk.Button(window, text="확인", width=10 , command=window.destroy).pack()
    
    window.mainloop()

    if mon.get() == 1:
        makeupDate['월'] += timedelta(days=7)
    if tue.get() == 1:
        makeupDate['화'] += timedelta(days=7)
    if wed.get() == 1:
        makeupDate['수'] += timedelta(days=7)
    if thu.get() == 1:
        makeupDate['목'] += timedelta(days=7)
    if fri.get() == 1:
        makeupDate['금'] += timedelta(days=7)
    if sat.get() == 1:
        makeupDate['토'] += timedelta(days=7)
    if sun.get() == 1:
        makeupDate['일'] += timedelta(days=7)
    gui.ui.wm_attributes("-disabled", False)

    if gui.saveDataButton['state'] == tk.NORMAL: return
    # 입력 양식 엑셀
    filepath = filedialog.askopenfilename(initialdir='./', title='데일리테스트 기록 양식 선택', filetypes=(('Excel files', '*.xlsx'),('all files', '*.*')))
    if filepath == '':
        gui.saveDataButton['state'] = tk.NORMAL
        return
    
    formWb = xl.load_workbook(filepath, data_only=True)
    formWs = formWb['데일리테스트 기록 양식']

    # 올바른 양식이 아닙니다.
    if (formWs.title != '데일리테스트 기록 양식') or (formWs['A1'].value != '요일') or (formWs['B1'].value != '시간') or (formWs['C1'].value != '반') or (formWs['D1'].value != '이름') or (formWs['E1'].value != '담당T') or (formWs['F1'].value != '시험명') or (formWs['G1'].value != '점수') or (formWs['H1'].value != '평균') or (formWs['I1'].value != '시험대비 모의고사명') or (formWs['J1'].value != '모의고사 점수') or (formWs['K1'].value != '모의고사 평균') or (formWs['L1'].value != '재시문자 X'):
        gui.appendLog('올바른 기록 양식이 아닙니다.')
        gui.saveDataButton['state'] = tk.NORMAL
        return
    

    # 데이터 저장 엑셀
    if not os.path.isfile('./data/' + config['dataFileName'] + '.xlsx'):
        gui.appendLog('[오류] ' + config['dataFileName'] + '.xlsx' + '파일이 존재하지 않습니다.')
        gui.saveDataButton['state'] = tk.NORMAL
        return

    dataFileWb = xl.load_workbook('./data/' + config['dataFileName'] + '.xlsx')
    dataFileWs = dataFileWb['DailyTest']
    
    for i in range(1, dataFileWs.max_column+1):
        if dataFileWs.cell(1, i).value == '반':
            classColumn = i
            break
    for i in range(1, dataFileWs.max_column+1):
        if dataFileWs.cell(1, i).value == '이름':
            nameColumn = i
            break
    for i in range(1, dataFileWs.max_column+1):
        if dataFileWs.cell(1, i).value == '학생 평균':
            scoreColumn = i
            break


    if not os.path.isfile('./재시험 정보.xlsx'):
        gui.appendLog('[오류] 재시험 정보.xlsx 파일이 존재하지 않습니다.')
        return
    makeupWb = xl.load_workbook("./재시험 정보.xlsx")
    try:
        makeupWs = makeupWb['재시험 정보']
    except:
        gui.appendLog('[오류] \'재시험 정보.xlsx\'의 시트명을')
        gui.appendLog('\'재시험 정보\'로 변경해 주세요.')
        gui.saveDataButton['state'] = tk.NORMAL
        return

    if not os.path.isfile('./data/재시험 명단.xlsx'):
        gui.appendLog('재시험 명단 파일 생성 중...')

        iniWb = xl.Workbook()
        iniWs = iniWb.active
        iniWs.title = '재시험 명단'
        iniWs['A1'] = '응시일'
        iniWs['B1'] = '반'
        iniWs['C1'] = '담당T'
        iniWs['D1'] = '이름'
        iniWs['E1'] = '시험명'
        iniWs['F1'] = '시험 점수'
        iniWs['G1'] = '재시 요일'
        iniWs['H1'] = '재시 시간'
        iniWs['I1'] = '재시 날짜'
        iniWs['J1'] = '재시 점수'
        iniWs['K1'] = '비고'
        iniWs.auto_filter.ref = 'A:K'
        iniWb.save('./data/재시험 명단.xlsx')
    
    makeupListWb = xl.load_workbook('./data/재시험 명단.xlsx')
    try:
        makeupListWs = makeupListWb['재시험 명단']
    except:
        gui.appendLog('[오류] \'재시험 명단.xlsx\'의 시트명을')
        gui.appendLog('\'재시험 명단\'으로 변경해 주세요.')
        gui.saveDataButton['state'] = tk.NORMAL
        return
    # try:
    gui.appendLog('데이터 저장 중...')

    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(os.getcwd() + '\\data\\' + config['dataFileName'] + '.xlsx')
    wb.Save()
    wb.Close()

    dailyWriteColumn = 7
    mockWriteColumn = 7
    for i in range(2, formWs.max_row+1):
        dailyTestScore = formWs.cell(i, 7).value
        mockTestScore = formWs.cell(i, 10).value
        # 파일 끝 검사
        if formWs.cell(i, 4).value is None:
            break
        
        # 반 필터링
        if formWs.cell(i, 3).value is not None: # form className is not None
            className = formWs.cell(i, 3).value
            dailyTestName = formWs.cell(i, 6).value
            teacher = formWs.cell(i, 5).value
            mockTestName = formWs.cell(i, 9).value
            if dailyTestName is None and mockTestName is None:
                continue # 시험 안 본 반 건너뛰기

            #반 시작 찾기
            for j in range(2, dataFileWs.max_row+1):
                if dataFileWs.cell(j, classColumn).value == className: # data className == form className
                    start = j # 데이터파일에서 반이 시작하는 행 번호
                    break
            # 반 끝 찾기
            for j in range(start, dataFileWs.max_row+1):
                if dataFileWs.cell(j, nameColumn).value == '시험 평균': # data name is 시험 평균
                    end = j #데이터파일에서 반이 끝나는 행 번호
                    break
            
            if dailyTestName is not None:
                # 데일리테스트 작성 열 위치 찾기
                dataFileWs = dataFileWb['DailyTest']
                for j in range(scoreColumn+1, dataFileWs.max_column+2):
                    if dataFileWs.cell(start, j).value is None:
                        dailyWriteColumn = j
                        break
                    if dataFileWs.cell(start, j).value.strftime('%y.%m.%d') == DATE.today().strftime('%y.%m.%d'):
                        dailyWriteColumn = j
                        break
                # 입력 틀 작성
                average = '=ROUND(AVERAGE(' + get_column_letter(dailyWriteColumn) + str(start + 2) + ':' + get_column_letter(dailyWriteColumn) + str(end - 1) + '), 0)'
                dataFileWs.cell(start, dailyWriteColumn).value = DATE.today()
                dataFileWs.cell(start, dailyWriteColumn).number_format = 'yyyy.mm.dd(aaa)'
                dataFileWs.cell(start, dailyWriteColumn).alignment = Alignment(horizontal='center', vertical='center')

                dataFileWs.cell(start + 1, dailyWriteColumn).value = dailyTestName
                dataFileWs.cell(start + 1, dailyWriteColumn).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

                dataFileWs.cell(end, dailyWriteColumn).value = average
                dataFileWs.cell(end, dailyWriteColumn).alignment = Alignment(horizontal='center', vertical='center')
                dataFileWs.cell(end, dailyWriteColumn).border = Border(bottom=Side(border_style='medium', color='000000'))
            
            if mockTestName is not None:
                # 모의고사 작성 열 위치 찾기
                dataFileWs = dataFileWb['모의고사']
                for j in range(scoreColumn+1, dataFileWs.max_column+2):
                    if dataFileWs.cell(start, j).value is None:
                        mockWriteColumn = j
                        break
                    if dataFileWs.cell(start, j).value.strftime('%y.%m.%d') == DATE.today().strftime('%y.%m.%d'):
                        mockWriteColumn = j
                        break
                # 입력 틀 작성
                average = '=ROUND(AVERAGE(' + get_column_letter(mockWriteColumn) + str(start + 2) + ':' + get_column_letter(mockWriteColumn) + str(end - 1) + '), 0)'
                dataFileWs.cell(start, mockWriteColumn).value = DATE.today()
                dataFileWs.cell(start, mockWriteColumn).number_format = 'yyyy.mm.dd(aaa)'
                dataFileWs.cell(start, mockWriteColumn).alignment = Alignment(horizontal='center', vertical='center')

                dataFileWs.cell(start + 1, mockWriteColumn).value = mockTestName
                dataFileWs.cell(start + 1, mockWriteColumn).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

                dataFileWs.cell(end, mockWriteColumn).value = average
                dataFileWs.cell(end, mockWriteColumn).alignment = Alignment(horizontal='center', vertical='center')
                dataFileWs.cell(end, mockWriteColumn).border = Border(bottom=Side(border_style='medium', color='000000'))

        if dailyTestScore is not None:
            dataFileWs = dataFileWb['DailyTest']
            testName = dailyTestName
            score = dailyTestScore
            writeColumn = dailyWriteColumn
        elif mockTestScore is not None:
            dataFileWs = dataFileWb['모의고사']
            testName = mockTestName
            score = mockTestScore
            writeColumn = mockWriteColumn
        else:
            continue # 점수 없으면 미응시 처리
        
        for j in range(start + 2, end):
            if dataFileWs.cell(j, nameColumn).value == formWs.cell(i, 4).value: # data name == form name
                dataFileWs.cell(j, writeColumn).value = score
                dataFileWs.cell(j, writeColumn).alignment = Alignment(horizontal='center', vertical='center')
                break
        
        dataFileWs = dataFileWb['DailyTest']

        # 재시험 작성
        if score < 80:
            for j in range(2, makeupWs.max_row+1):
                if makeupWs.cell(j, 1).value == formWs.cell(i, 4).value:
                    dates = makeupWs.cell(j, 4).value
                    time = makeupWs.cell(j, 5).value
                    newFace = makeupWs.cell(j, 6).value
                    break
            for j in range(2, makeupListWs.max_row + 2):
                if makeupListWs.cell(j, 1).value is None:
                    writeRow = j
                    break
            makeupListWs.cell(writeRow, 1).value = DATE.today()
            makeupListWs.cell(writeRow, 2).value = className
            makeupListWs.cell(writeRow, 3).value = teacher
            makeupListWs.cell(writeRow, 4).value = formWs.cell(i, 4).value
            if (newFace is not None) and (newFace == 'N'):
                makeupListWs.cell(writeRow, 4).fill = PatternFill(fill_type='solid', fgColor=Color('FFFF00'))
            makeupListWs.cell(writeRow, 5).value = testName
            makeupListWs.cell(writeRow, 6).value = score
            if dates is not None:
                makeupListWs.cell(writeRow, 7).value = dates
                dateList = dates.split('/')
                result = makeupDate[dateList[0].replace(' ', '')]
                for d in dateList:
                    if result > makeupDate[d.replace(' ', '')]:
                        result = makeupDate[d.replace(' ', '')]
                if time is not None:
                    makeupListWs.cell(writeRow, 8).value = time
                makeupListWs.cell(writeRow, 9).value = result
                makeupListWs.cell(writeRow, 9).number_format = 'mm월 dd일(aaa)'

    # except:
    #     gui.appendLog('이 데이터 파일 양식에는 작성할 수 없습니다.')
    #     gui.saveDataButton['state'] = tk.NORMAL
    #     return

    for j in range(1, makeupListWs.max_row + 1):
        for k in range(1, makeupListWs.max_column + 1):
            makeupListWs.cell(j, k).alignment = Alignment(horizontal='center', vertical='center')
            makeupListWs.cell(j, k).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    gui.appendLog('재시험 명단 작성 중...')
    
    gui.appendLog('백업 파일 생성중...')
    formWb = xl.load_workbook(filepath)
    formWs = formWb['데일리테스트 기록 양식']
    formWb.save('./data/backup/데일리테스트 기록 양식(' + datetime.today().strftime('%Y%m%d') + ').xlsx')
    # 양식 지우기
    # for i in range(2, formWs.max_row + 1):
    #     formWs.cell(i, 6).value = ''
    #     formWs.cell(i, 7).value = ''
    #     formWs.cell(i, 9).value = ''
    #     formWs.cell(i, 10).value = ''
    # formWb.save('./데일리테스트 기록 양식.xlsx')
    # except:
    #     gui.appendLog('재시험명단 파일 창을 끄고 다시 실행해 주세요.')
    #     gui.saveDataButton['state'] = tk.NORMAL
    #     return
    
    try:
        dataFileWb.save('./data/' + config['dataFileName'] + '.xlsx')
    except:
        gui.appendLog('데이터 파일 창을 끄고 다시 실행해 주세요.')
        gui.saveDataButton['state'] = tk.NORMAL
        return

    # 조건부서식
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(os.getcwd() + '\\data\\' + config['dataFileName'] + '.xlsx')
    wb.Save()
    wb.Close()

    dataFileWb = xl.load_workbook('./data/' + config['dataFileName'] + '.xlsx')
    dataFileColorWb = xl.load_workbook('./data/' + config['dataFileName'] + '.xlsx', data_only=True)
    sheetNames = dataFileWb.sheetnames
    for i in range(0, 2):
        dataFileWs = dataFileWb[sheetNames[i]]
        dataFileColorWs = dataFileColorWb[sheetNames[i]]
        for j in range(2, dataFileColorWs.max_row+1):
            for k in range(scoreColumn+1, dataFileColorWs.max_column+1):
                if k > scoreColumn:
                    dataFileWs.column_dimensions[get_column_letter(j)].width = 14

                if type(dataFileColorWs.cell(j, k).value) == int:
                    if dataFileColorWs.cell(j, k).value < 60:
                        dataFileWs.cell(j, k).fill = PatternFill(fill_type='solid', fgColor=Color('EC7E31'))
                    elif dataFileColorWs.cell(j, k).value < 70:
                        dataFileWs.cell(j, k).fill = PatternFill(fill_type='solid', fgColor=Color('F5AF85'))
                    elif dataFileColorWs.cell(j, k).value < 80:
                        dataFileWs.cell(j, k).fill = PatternFill(fill_type='solid', fgColor=Color('FCE4D6'))
                    elif dataFileColorWs.cell(j, nameColumn).value == '시험 평균':
                        dataFileWs.cell(j, k).fill = PatternFill(fill_type='solid', fgColor=Color('DDEBF7'))

            # 학생별 평균 조건부 서식
            if type(dataFileColorWs.cell(j, scoreColumn).value) == int:
                if dataFileColorWs.cell(j, scoreColumn).value < 60:
                    dataFileWs.cell(j, scoreColumn).fill = PatternFill(fill_type='solid', fgColor=Color('EC7E31'))
                elif dataFileColorWs.cell(j, scoreColumn).value < 70:
                    dataFileWs.cell(j, scoreColumn).fill = PatternFill(fill_type='solid', fgColor=Color('F5AF85'))
                elif dataFileColorWs.cell(j, scoreColumn).value < 80:
                    dataFileWs.cell(j, scoreColumn).fill = PatternFill(fill_type='solid', fgColor=Color('FCE4D6'))
                elif dataFileColorWs.cell(j, nameColumn).value == '시험 평균':
                    dataFileWs.cell(j, scoreColumn).fill = PatternFill(fill_type='solid', fgColor=Color('DDEBF7'))
                else:
                    dataFileWs.cell(j, scoreColumn).fill = PatternFill(fill_type='solid', fgColor=Color('E2EFDA'))
            name = dataFileWs.cell(j, nameColumn)
            if (name.value is not None) or (name.value != '날짜') or (name.value != '시험명') or (name.value != '시험 평균'):
                for k in range(2, makeupWs.max_row+1):
                    if (name.value == makeupWs.cell(k, 1).value) and (makeupWs.cell(k, 6).value == 'N'):
                        name.fill = PatternFill(fill_type='solid', fgColor=Color('FFFF00'))
                        break
    dataFileWs = dataFileWb['DailyTest']
    dataFileWb.save('./data/' + config['dataFileName'] + '.xlsx')
    try:
        makeupListWb.save('./data/재시험 명단.xlsx')
    except:
        gui.appendLog('재시험 명단 파일 창을 끄고 다시 실행해 주세요.')
        gui.saveDataButton['state'] = tk.NORMAL
        return
    
    gui.appendLog('데이터 저장을 완료했습니다.')
    gui.ui.wm_attributes("-topmost", 1)
    gui.ui.wm_attributes("-topmost", 0)
    gui.saveDataButton['state'] = tk.NORMAL
    excel.Visible = True
    wb = excel.Workbooks.Open(os.getcwd() + '\\data\\' + config['dataFileName'] + '.xlsx')

def classInfo(gui):
    if not os.path.isfile('반 정보.xlsx'):
        gui.appendLog('반 정보 입력 파일 생성 중...')

        iniWb = xl.Workbook()
        iniWs = iniWb.active
        iniWs.title = '반 정보'
        iniWs['A1'] = '반명'
        iniWs['B1'] = '선생님명'
        iniWs['C1'] = '요일'
        iniWs['D1'] = '시간'

        options = webdriver.ChromeOptions()
        options.add_argument('headless')
        driver = webdriver.Chrome(service = service, options = options)

        # 아이소식 접속
        driver.get(config['url'])
        tableNames = driver.find_elements(By.CLASS_NAME, 'style1')

        # 반 루프
        for i in range(3, len(tableNames)):
            trs = driver.find_element(By.ID, 'table_' + str(i)).find_elements(By.CLASS_NAME, 'style12')
            writeLocation = start = iniWs.max_row + 1
            iniWs.cell(writeLocation, 1).value = tableNames[i].text.rstrip()

        # 정렬 및 테두리
        for j in range(1, iniWs.max_row + 1):
                for k in range(1, iniWs.max_column + 1):
                    iniWs.cell(j, k).alignment = Alignment(horizontal='center', vertical='center')
                    iniWs.cell(j, k).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        iniWb.save('./반 정보.xlsx')
        gui.classInfoButton['state'] = tk.DISABLED
        gui.appendLog('반 정보 입력 파일 생성을 완료했습니다.')
        gui.appendLog('반 정보를 입력해 주세요.')
        gui.ui.wm_attributes("-topmost", 1)
        gui.ui.wm_attributes("-topmost", 0)

    else:
        gui.appendLog('이미 파일이 존재합니다.')
        gui.classInfoButton['state'] = tk.DISABLED

def sendMessage(gui):
    def quitEvent():
        window.destroy()
        gui.sendMessageButton['state'] = tk.NORMAL
        gui.ui.wm_attributes("-disabled", False)
        gui.ui.wm_attributes("-topmost", 1)
        gui.ui.wm_attributes("-topmost", 0)
    gui.sendMessageButton['state'] = tk.DISABLED
    gui.ui.wm_attributes("-disabled", True)
    window=tk.Tk()
    window.geometry('200x300+500+500')
    window.resizable(False, False)
    window.protocol("WM_DELETE_WINDOW", quitEvent)

    today = DATE.today()

    makeupDate={}
    if today == today + relativedelta(weekday=calendar.MONDAY):
        makeupDate['월'] = today + timedelta(days=7)
    else:
        makeupDate['월'] = today + relativedelta(weekday=calendar.MONDAY)

    if today == today + relativedelta(weekday=calendar.TUESDAY):
        makeupDate['화'] = today + timedelta(days=7)
    else:
        makeupDate['화'] = today + relativedelta(weekday=calendar.TUESDAY)

    if today == today + relativedelta(weekday=calendar.WEDNESDAY):
        makeupDate['수'] = today + timedelta(days=7)
    else:
        makeupDate['수'] = today + relativedelta(weekday=calendar.WEDNESDAY)

    if today == today + relativedelta(weekday=calendar.THURSDAY):
        makeupDate['목'] = today + timedelta(days=7)
    else:
        makeupDate['목'] = today + relativedelta(weekday=calendar.THURSDAY)

    if today == today + relativedelta(weekday=calendar.FRIDAY):
        makeupDate['금'] = today + timedelta(days=7)
    else:
        makeupDate['금'] = today + relativedelta(weekday=calendar.FRIDAY)

    if today == today + relativedelta(weekday=calendar.SATURDAY):
        makeupDate['토'] = today + timedelta(days=7)
    else:
        makeupDate['토'] = today + relativedelta(weekday=calendar.SATURDAY)

    if today == today + relativedelta(weekday=calendar.SUNDAY):
        makeupDate['일'] = today + timedelta(days=7)
    else:
        makeupDate['일'] = today + relativedelta(weekday=calendar.SUNDAY)

    mon = tk.IntVar()
    tue = tk.IntVar()
    wed = tk.IntVar()
    thu = tk.IntVar()
    fri = tk.IntVar()
    sat = tk.IntVar()
    sun = tk.IntVar()

    tk.Label(window, text='\n다음 중 휴일을 선택해주세요\n').pack()
    dateCalc = DATE.today() + timedelta(days=1)
    for i in range(0, 8):
        for j in range(0, 8):
            if dateCalc == makeupDate['월']:
                tk.Checkbutton(window, text=str(makeupDate['월'])+' (월)', variable=mon).pack()
                dateCalc += timedelta(days=1)
        for j in range(0, 8):
            if dateCalc == makeupDate['화']:
                tk.Checkbutton(window, text=str(makeupDate['화'])+' (화)', variable=tue).pack()
                dateCalc += timedelta(days=1)
        for j in range(0, 8):
            if dateCalc == makeupDate['수']:
                tk.Checkbutton(window, text=str(makeupDate['수'])+' (수)', variable=wed).pack()
                dateCalc += timedelta(days=1)
        for j in range(0, 8):
            if dateCalc == makeupDate['목']:
                tk.Checkbutton(window, text=str(makeupDate['목'])+' (목)', variable=thu).pack()
                dateCalc += timedelta(days=1)
        for j in range(0, 8):
            if dateCalc == makeupDate['금']:
                tk.Checkbutton(window, text=str(makeupDate['금'])+' (금)', variable=fri).pack()
                dateCalc += timedelta(days=1)
        for j in range(0, 8):
            if dateCalc == makeupDate['토']:
                tk.Checkbutton(window, text=str(makeupDate['토'])+' (토)', variable=sat).pack()
                dateCalc += timedelta(days=1)
        for j in range(0, 8):
            if dateCalc == makeupDate['일']:
                tk.Checkbutton(window, text=str(makeupDate['일'])+' (일)', variable=sun).pack()
                dateCalc += timedelta(days=1)
    tk.Label(window, text='\n').pack()
    tk.Button(window, text="확인", width=10 , command=window.destroy).pack()
    
    window.mainloop()
    if mon.get() == 1:
        makeupDate['월'] += timedelta(days=7)
    if tue.get() == 1:
        makeupDate['화'] += timedelta(days=7)
    if wed.get() == 1:
        makeupDate['수'] += timedelta(days=7)
    if thu.get() == 1:
        makeupDate['목'] += timedelta(days=7)
    if fri.get() == 1:
        makeupDate['금'] += timedelta(days=7)
    if sat.get() == 1:
        makeupDate['토'] += timedelta(days=7)
    if sun.get() == 1:
        makeupDate['일'] += timedelta(days=7)
    gui.ui.wm_attributes("-disabled", False)
    
    if gui.sendMessageButton['state'] == tk.NORMAL: return

    filepath = filedialog.askopenfilename(initialdir='./', title='데일리테스트 기록 양식 선택', filetypes=(('Excel files', '*.xlsx'),('all files', '*.*')))
    if filepath == '':
        gui.sendMessageButton['state'] = tk.NORMAL
        return
    
    # 점수기록표 xlsx
    formWb = xl.load_workbook(filepath, data_only=True)
    formWs = formWb['데일리테스트 기록 양식']
    
    # 올바른 양식이 아닙니다.
    if (formWs.title != '데일리테스트 기록 양식') or (formWs['A1'].value != '요일') or (formWs['B1'].value != '시간') or (formWs['C1'].value != '반') or (formWs['D1'].value != '이름') or (formWs['E1'].value != '담당T') or (formWs['F1'].value != '시험명') or (formWs['G1'].value != '점수') or (formWs['H1'].value != '평균') or (formWs['I1'].value != '시험대비 모의고사명') or (formWs['J1'].value != '모의고사 점수') or (formWs['K1'].value != '모의고사 평균') or (formWs['L1'].value != '재시문자 X'):
        gui.appendLog('올바른 기록 양식이 아닙니다.')
        gui.sendMessageButton['state'] = tk.NORMAL
        return

    # 휴일 선택창
    try:
        gui.appendLog('크롬을 실행시키는 중...')
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        driver = webdriver.Chrome(service=service, options=options)

        if not os.path.isfile('./재시험 정보.xlsx'):
            gui.appendLog('[오류] 재시험 정보.xlsx 파일이 존재하지 않습니다.')
            return
        makeupWb = xl.load_workbook("./재시험 정보.xlsx")
        try:
            makeupWs = makeupWb['재시험 정보']
        except:
            gui.appendLog('[오류] \'재시험 정보.xlsx\'의 시트명을')
            gui.appendLog('\'재시험 정보\'로 변경해 주세요.')
            gui.sendMessageButton['state'] = tk.NORMAL
            return
        
        # 아이소식 접속
        driver.get(config['url'])
        driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(config['dailyTest'])
        
        driver.execute_script('window.open('');')
        driver.switch_to.window(driver.window_handles[1])
        driver.get(config['url'])
        driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(config['makeupTest'])

        driver.execute_script('window.open('');')
        driver.switch_to.window(driver.window_handles[2])
        driver.get(config['url'])
        driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(config['makeupTestDate'])

        gui.appendLog('메시지 작성 중...')
        for i in range(2, formWs.max_row+1):
            driver.switch_to.window(driver.window_handles[0])
            name = formWs.cell(i, 4).value
            dailyTestScore = formWs.cell(i, 7).value
            mockTestScore = formWs.cell(i, 10).value
            if formWs.cell(i, 3).value is not None:
                className = formWs.cell(i, 3).value
                dailyTestName = formWs.cell(i, 6).value
                mockTestName = formWs.cell(i, 9).value
                dailyTestAverage = formWs.cell(i, 8).value
                mockTestAverage = formWs.cell(i, 11).value

            # 시험 미응시시 건너뛰기
            if dailyTestScore is not None:
                testName = dailyTestName
                score = dailyTestScore
                average = dailyTestAverage
            elif mockTestScore is not None:
                testName = mockTestName
                score = mockTestScore
                average = mockTestAverage
            else:
                continue
            
            if testName is None:
                gui.appendLog(className + '반의 시험명이 없습니다.')
                gui.appendLog('시험 결과 전송을 중단합니다.')
                driver.quit()
            
            tableNames = driver.find_elements(By.CLASS_NAME, 'style1')
            for j in range(len(tableNames)):
                if className in tableNames[j].text:
                    index = j
                    break

            trs = driver.find_element(By.ID, 'table_' + str(index)).find_elements(By.CLASS_NAME, 'style12')
            for tr in trs:
                if tr.find_element(By.CLASS_NAME, 'style9').text == name:
                    tds = tr.find_elements(By.TAG_NAME, 'td')
                    tds[0].find_element(By.TAG_NAME, 'input').send_keys(testName)
                    tds[1].find_element(By.TAG_NAME, 'input').send_keys(score)
                    tds[2].find_element(By.TAG_NAME, 'input').send_keys(average)
                    break
            
            if score < 80 and ((formWs.cell(i, 12).value != 'x') or (formWs.cell(i, 12).value != 'X')):
                for j in range(2, makeupWs.max_row+1):
                    if makeupWs.cell(j, 1).value == name:
                        date = makeupWs.cell(j, 4).value
                        time = makeupWs.cell(j, 5).value
                        break
                if date is None:
                    driver.switch_to.window(driver.window_handles[1])
                    trs = driver.find_element(By.ID, 'table_' + str(index)).find_elements(By.CLASS_NAME, 'style12')
                    for tr in trs:
                        if tr.find_element(By.CLASS_NAME, 'style9').text == name:
                            tds = tr.find_elements(By.TAG_NAME, 'td')
                            tds[0].find_element(By.TAG_NAME, 'input').send_keys(testName)
                else:
                    dateList = date.split('/')
                    result = makeupDate[dateList[0].replace(' ', '')]
                    timeIndex = 0
                    for i in range(len(dateList)):
                        if result > makeupDate[dateList[i].replace(' ', '')]:
                            result = makeupDate[dateList[i].replace(' ', '')]
                            timeIndex = i
                    driver.switch_to.window(driver.window_handles[2])
                    trs = driver.find_element(By.ID, 'table_' + str(index)).find_elements(By.CLASS_NAME, 'style12')
                    for tr in trs:
                        if tr.find_element(By.CLASS_NAME, 'style9').text == name:
                            tds = tr.find_elements(By.TAG_NAME, 'td')
                            tds[0].find_element(By.TAG_NAME, 'input').send_keys(testName)
                            if time is not None:
                                tds[1].find_element(By.TAG_NAME, 'input').send_keys(result.strftime('%m월 %d일') + ' ' + str(time).split('/')[timeIndex] + '시')
                            else:
                                tds[1].find_element(By.TAG_NAME, 'input').send_keys(result.strftime('%m월 %d일'))
        gui.sendMessageButton['state'] = tk.NORMAL
        gui.appendLog('메시지 입력을 완료했습니다.')
        gui.appendLog('메시지 확인 후 전송해주세요.')
        gui.ui.wm_attributes("-topmost", 1)
        gui.ui.wm_attributes("-topmost", 0)
    except:
        gui.appendLog('중단되었습니다.')
        gui.sendMessageButton['state'] = tk.NORMAL
        return

def makeupTestInfo(gui):
    if not os.path.isfile('./재시험 정보.xlsx'):
        gui.appendLog('재시험 정보 파일 생성 중...')

        iniWb = xl.Workbook()
        iniWs = iniWb.active
        iniWs.title = '재시험 정보'
        iniWs['A1'] = '이름'
        iniWs['B1'] = '반명'
        iniWs['C1'] = '담당'
        iniWs['D1'] = '요일'
        iniWs['E1'] = '시간'
        iniWs['F1'] = '기수 신규생'
        iniWs['G1'] = 'N'
        iniWs.auto_filter.ref = 'A:F'
        iniWs.column_dimensions.group('G', hidden=True)

        # 반 정보 확인
        if not os.path.isfile('./반 정보.xlsx'):
            gui.appendLog('[오류] 반 정보.xlsx 파일이 존재하지 않습니다.')
            return
        classWb = xl.load_workbook("./반 정보.xlsx")
        try:
            classWs = classWb['반 정보']
        except:
            gui.appendLog('[오류] \'반 정보.xlsx\'의 시트명을')
            gui.appendLog('\'반 정보\'로 변경해 주세요.')
            gui.makeupTestInfoButton['state'] = tk.NORMAL
            return

        options = webdriver.ChromeOptions()
        options.add_argument('headless')
        driver = webdriver.Chrome(service = service, options = options)

        # 아이소식 접속
        driver.get(config['url'])
        tableNames = driver.find_elements(By.CLASS_NAME, 'style1')

        # 반 루프
        for i in range(3, len(tableNames)):
            trs = driver.find_element(By.ID, 'table_' + str(i)).find_elements(By.CLASS_NAME, 'style12')
            writeLocation = iniWs.max_row + 1
            teacher = ''

            className = tableNames[i].text.rstrip()
            isClassExist = False
            for j in range(2, classWs.max_row + 1):
                if classWs.cell(j, 1).value == className:
                    teacher = classWs.cell(j, 2).value
                    isClassExist = True
            if not isClassExist:
                continue

            # 학생 루프
            for tr in trs:
                writeLocation = iniWs.max_row + 1
                iniWs.cell(writeLocation, 1).value = tr.find_element(By.CLASS_NAME, 'style9').text
                iniWs.cell(writeLocation, 2).value = className
                iniWs.cell(writeLocation, 3).value = teacher
                dv = DataValidation(type='list', formula1='=G1',  allow_blank=True, errorStyle='stop', showErrorMessage=True)
                iniWs.add_data_validation(dv)
                dv.add(iniWs.cell(writeLocation, 6))

        # 정렬 및 테두리
        for j in range(1, iniWs.max_row + 1):
                for k in range(1, iniWs.max_column + 1):
                    iniWs.cell(j, k).alignment = Alignment(horizontal='center', vertical='center')
                    iniWs.cell(j, k).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        iniWb.save('./재시험 정보.xlsx')
        gui.appendLog('재시험 정보 파일을 생성했습니다.')
        gui.ui.wm_attributes("-topmost", 1)
        gui.ui.wm_attributes("-topmost", 0)

    else:
        gui.appendLog('이미 파일이 존재합니다')
        gui.makeupTestInfoButton['state'] = tk.DISABLED

def applyColor(gui):
    gui.applyColorButton['state'] = tk.DISABLED
    try:
        if not os.path.isfile('./data/' + config['dataFileName'] + '.xlsx'):
            gui.appendLog('[오류] '+ config['dataFileName'] +'.xlsx 파일이 존재하지 않습니다.')
            gui.applyColorButton['state'] = tk.NORMAL
            return
        
        if not os.path.isfile('./재시험 정보.xlsx'):
            gui.appendLog('[오류] 재시험 정보.xlsx 파일이 존재하지 않습니다.')
            return
        makeupWb = xl.load_workbook("./재시험 정보.xlsx")
        try:
            makeupWs = makeupWb['재시험 정보']
        except:
            gui.appendLog('[오류] \'재시험 정보.xlsx\'의 시트명을')
            gui.appendLog('\'재시험 정보\'로 변경해 주세요.')
            gui.applyColorButton['state'] = tk.NORMAL
            return
        
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = False
        wb = excel.Workbooks.Open(os.getcwd() + '\\data\\' + config['dataFileName'] + '.xlsx')
        wb.Save()
        wb.Close()

        gui.appendLog('조건부 서식 적용중...')

        dataFileWb = xl.load_workbook('./data/' + config['dataFileName'] + '.xlsx')
        dataFileColorWb = xl.load_workbook('./data/' + config['dataFileName'] + '.xlsx', data_only=True)
    except:
        gui.appendLog('데이터 파일 창을 끄고 다시 실행해 주세요.')
        gui.applyColorButton['state'] = tk.NORMAL
        return
    
    try:
        for sheetName in dataFileWb.sheetnames:
            dataFileWs = dataFileWb[sheetName]
            dataFileColorWs = dataFileColorWb[sheetName]

            for i in range(1, dataFileWs.max_column):
                if dataFileWs.cell(1, i).value == '이름':
                    nameColumn = i
                    break
            for i in range(1, dataFileWs.max_column):
                if dataFileWs.cell(1, i).value == '학생 평균':
                    scoreColumn = i
                    break
            
            for i in range(2, dataFileColorWs.max_row+1):
                if dataFileColorWs.cell(i, nameColumn).value is None:
                    break
                for j in range(scoreColumn+1, dataFileColorWs.max_column+1):
                    dataFileWs.column_dimensions[get_column_letter(j)].width = 14
                    if dataFileWs.cell(i, nameColumn).value == '시험 평균' and dataFileWs.cell(i, j).value is not None:
                        dataFileWs.cell(i, j).border = Border(bottom=Side(border_style='medium', color='000000'))
                    if dataFileWs.cell(i, nameColumn).value == '날짜' and dataFileWs.cell(i, j).value is not None:
                        dataFileWs.cell(i, j).border = Border(top=Side(border_style='medium', color='000000'))
                    if type(dataFileColorWs.cell(i, j).value) == int:
                        if dataFileColorWs.cell(i, j).value < 60:
                            dataFileWs.cell(i, j).fill = PatternFill(fill_type='solid', fgColor=Color('EC7E31'))
                        elif dataFileColorWs.cell(i, j).value < 70:
                            dataFileWs.cell(i, j).fill = PatternFill(fill_type='solid', fgColor=Color('F5AF85'))
                        elif dataFileColorWs.cell(i, j).value < 80:
                            dataFileWs.cell(i, j).fill = PatternFill(fill_type='solid', fgColor=Color('FCE4D6'))
                        elif dataFileColorWs.cell(i, nameColumn).value == '시험 평균':
                            dataFileWs.cell(i, j).fill = PatternFill(fill_type='solid', fgColor=Color('DDEBF7'))
                        else:
                            dataFileWs.cell(i, j).fill = PatternFill(fill_type=None, fgColor=Color('00FFFFFF'))
                    if dataFileColorWs.cell(i, nameColumn).value == '시험 평균':
                        dataFileWs.cell(i, j).font = Font(bold=True)

                # 학생별 평균 조건부 서식
                dataFileWs.cell(i, scoreColumn).font = Font(bold=True)
                if type(dataFileColorWs.cell(i, scoreColumn).value) == int:
                    if dataFileColorWs.cell(i, scoreColumn).value < 60:
                        dataFileWs.cell(i, scoreColumn).fill = PatternFill(fill_type='solid', fgColor=Color('EC7E31'))
                    elif dataFileColorWs.cell(i, scoreColumn).value < 70:
                        dataFileWs.cell(i, scoreColumn).fill = PatternFill(fill_type='solid', fgColor=Color('F5AF85'))
                    elif dataFileColorWs.cell(i, scoreColumn).value < 80:
                        dataFileWs.cell(i, scoreColumn).fill = PatternFill(fill_type='solid', fgColor=Color('FCE4D6'))
                    elif dataFileColorWs.cell(i, nameColumn).value == '시험 평균':
                        dataFileWs.cell(i, scoreColumn).fill = PatternFill(fill_type='solid', fgColor=Color('DDEBF7'))
                    else:
                        dataFileWs.cell(i, scoreColumn).fill = PatternFill(fill_type='solid', fgColor=Color('E2EFDA'))
                name = dataFileWs.cell(i, nameColumn)
                if (name.value is not None) or (name.value != '날짜') or (name.value != '시험명') or (name.value != '시험 평균'):
                    for j in range(2, makeupWs.max_row+1):
                        if (name.value == makeupWs.cell(j, 1).value) and (makeupWs.cell(j, 6).value == 'N'):
                            name.fill = PatternFill(fill_type='solid', fgColor=Color('FFFF00'))
                            break

    except:
        gui.appendLog('이 데이터 양식에는 조건부 서식을 지정할 수 없습니다.')
        gui.applyColorButton['state'] = tk.NORMAL
        return
    
    try:
        dataFileWb.save('./data/' + config['dataFileName'] + '.xlsx')
        gui.appendLog('조건부 서식 지정을 완료했습니다.')
        gui.ui.wm_attributes("-topmost", 1)
        gui.ui.wm_attributes("-topmost", 0)
        gui.applyColorButton['state'] = tk.NORMAL
        excel.Visible = True
        wb = excel.Workbooks.Open(os.getcwd() + '\\data\\' + config['dataFileName'] + '.xlsx')
        gui.ui.wm_attributes("-topmost", 1)
        gui.ui.wm_attributes("-topmost", 0)
    except:
        gui.appendLog('데이터 파일 창을 끄고 다시 실행해 주세요.')
        gui.applyColorButton['state'] = tk.NORMAL
        return