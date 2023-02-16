import json
import os.path
import calendar
import tkinter as tk
import openpyxl as xl
import win32com.client # only works in Windows

from tkinter import filedialog
from datetime import date, datetime, timedelta
from openpyxl.utils.cell import get_column_letter
from dateutil.relativedelta import relativedelta
from openpyxl.styles import Alignment, Border, Color, PatternFill, Side

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
            gui.appendLog('[오류]반 정보.xlsx 파일이 존재하지 않습니다.')
            return
        classWb = xl.load_workbook("./반 정보.xlsx")
        classWs = classWb.active

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

            className = tableNames[i].text.split('(')[0].rstrip()
            for j in range(2, classWs.max_row + 1):
                if classWs.cell(j, 1).value == className:
                    teacher = classWs.cell(j, 2).value
                    date = classWs.cell(j, 3).value
                    time = classWs.cell(j, 4).value
            
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

            # 학생 루프
            for tr in trs:
                writeLocation = iniWs.max_row + 1
                iniWs.cell(writeLocation, 1).value = time
                iniWs.cell(writeLocation, 2).value = date
                iniWs.cell(writeLocation, 3).value = className
                iniWs.cell(writeLocation, 4).value = teacher
                iniWs.cell(writeLocation, 5).value = tr.find_element(By.CLASS_NAME, 'style9').text
                iniWs.cell(writeLocation, 6).value = '=ROUND(AVERAGE(G' + str(writeLocation) + ':XFD' + str(writeLocation) + '), 0)'
            
            # 시험별 평균
            writeLocation = iniWs.max_row + 1
            iniWs.cell(writeLocation, 1).value = time
            iniWs.cell(writeLocation, 2).value = date
            iniWs.cell(writeLocation, 3).value = className
            iniWs.cell(writeLocation, 4).value = teacher
            iniWs.cell(writeLocation, 5).value = '시험 평균'
            iniWs.cell(writeLocation, 6).value = '=ROUND(AVERAGE(G' + str(writeLocation) + ':XFD' + str(writeLocation) + '), 0)'

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
    iniWs.auto_filter.ref = 'A:B'

    if not os.path.isfile('./반 정보.xlsx'):
        gui.appendLog('[오류]반 정보.xlsx 파일이 존재하지 않습니다.')
        return
    
    classWb = xl.load_workbook("반 정보.xlsx")
    classWs = classWb.active

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

        className = tableNames[i].text.split('(')[0].rstrip()

        iniWs.cell(writeLocation, 3).value = className
        for j in range(2, classWs.max_row + 1):
            if classWs.cell(j, 1).value == className:
                teacher = classWs.cell(j, 2).value
                date = classWs.cell(j, 3).value
                time = classWs.cell(j, 4).value
        iniWs.cell(writeLocation, 5).value = teacher

        #학생 루프
        for tr in trs:
            iniWs.cell(writeLocation, 1).value = date
            iniWs.cell(writeLocation, 2).value = time
            iniWs.cell(writeLocation, 4).value = tr.find_element(By.CLASS_NAME, 'style9').text
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

def saveData(gui):
    gui.saveDataButton['state'] = tk.DISABLED
    filepath = filedialog.askopenfilename(initialdir='./', title='데일리테스트 기록 양식 선택', filetypes=(('Excel files', '*.xlsx'),('all files', '*.*')))
    if filepath == '':
        gui.saveDataButton['state'] = tk.NORMAL
        return
    gui.appendLog('데이터 저장 중...')

    # 파일 위치 및 파일명 지정
    if not os.path.isfile('./data/' + config['dataFileName'] + '.xlsx'):
        gui.appendLog('[오류]' + config['dataFileName'] + '.xlsx' + '파일이 존재하지 않습니다.')
        return

    # 입력 양식 엑셀
    formWb = xl.load_workbook(filepath, data_only=True)
    formWs = formWb.active
    # 데이터 저장 엑셀
    dataFileWb = xl.load_workbook('./data/' + config['dataFileName'] + '.xlsx')
    dataFileWs = dataFileWb['DailyTest']
    # 재시험 엑셀
    if not os.path.isfile('./data/재시험명단.xlsx'):
        gui.appendLog('재시험 명단 파일 생성 중...')

        iniWb = xl.Workbook()
        iniWs = iniWb.active
        iniWs.title = 'Make-upTest'
        iniWs['A1'] = '응시일'
        iniWs['B1'] = '반'
        iniWs['C1'] = '담당T'
        iniWs['D1'] = '이름'
        iniWs['E1'] = '시험명'
        iniWs['F1'] = '시험 점수'
        iniWs['G1'] = '재시 응시 여부'
        iniWs.auto_filter.ref = 'A:G'
        iniWb.save('./data/재시험명단.xlsx')
    
    makeupListWb = xl.load_workbook('./data/재시험명단.xlsx')
    makeupListWs = makeupListWb.active

    # today = datetime.today().strftime('%Y-%m-%d')
    # 데이터 날짜 내림차순
    # writeColumn = 7
    # if str(dataFileWs.cell(1, writeColumn).value) != today:
    #     dataFileWs.insert_cols(writeColumn)
    #     for i in range(2, dataFileWs.max_row+1):
    #         if dataFileWs.cell(i, 5).value == '시험 평균':
    #             dataFileWs.cell(i, 7).border = Border(bottom=Side(border_style='medium', color='000000'))
    # dataFileWs.cell(1, writeColumn).value = today

    # for i in range(2, formWs.max_row+1):
    #     if formWs.cell(i, 9).value is not None:
    #         dataFileWs = dataFileWb['모의고사']
    #         if str(dataFileWs.cell(1, writeColumn).value) != today:
    #             dataFileWs.insert_cols(writeColumn)
    #             for i in range(2, dataFileWs.max_row+1):
    #                 if dataFileWs.cell(i, 5).value == '시험 평균':
    #                     dataFileWs.cell(i, 7).border = Border(bottom=Side(border_style='medium', color='000000'))
    #         dataFileWs.cell(1, writeColumn).value = today
    #         dataFileWs = dataFileWb['DailyTest']
    #         break

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
                continue
            #반 시작 찾기
            for j in range(2, dataFileWs.max_row+1):
                if dataFileWs.cell(j, 3).value == className: # data className == form className
                    start = j
                    break
            # 반 끝 찾기
            for j in range(start, dataFileWs.max_row+1):
                if dataFileWs.cell(j, 5).value == '시험 평균': # data name is 시험 평균
                    end = j
                    break
            
            
            if dailyTestName is not None:
                # 작성 위치 찾기
                dataFileWs = dataFileWb['DailyTest']
                for j in range(7, dataFileWs.max_column+2):
                    if dataFileWs.cell(start, j).value == date.today():
                        dailyWriteColumn = j
                        break
                    if dataFileWs.cell(start, j).value is None:
                        dailyWriteColumn = j
                        break
                average = '=ROUND(AVERAGE(' + get_column_letter(dailyWriteColumn) + str(start + 2) + ':' + get_column_letter(dailyWriteColumn) + str(end - 1) + '), 0)'
                dataFileWs.cell(start, dailyWriteColumn).value = date.today()
                dataFileWs.cell(start, dailyWriteColumn).number_format = 'yyyy.mm.dd(aaa)'
                dataFileWs.cell(start, dailyWriteColumn).alignment = Alignment(horizontal='center', vertical='center')
                dataFileWs.cell(start + 1, dailyWriteColumn).value = dailyTestName
                dataFileWs.cell(start + 1, dailyWriteColumn).alignment = Alignment(horizontal='center', vertical='center')
                dataFileWs.cell(end, dailyWriteColumn).value = average
                dataFileWs.cell(end, dailyWriteColumn).alignment = Alignment(horizontal='center', vertical='center')
                dataFileWs.cell(end, dailyWriteColumn).border = Border(bottom=Side(border_style='medium', color='000000'))
            if mockTestName is not None:
                # 작성 위치 찾기
                dataFileWs = dataFileWb['모의고사']
                for j in range(7, dataFileWs.max_column+2):
                    if dataFileWs.cell(start, j).value == date.today():
                        makeupWriteColumn = j
                        break
                    if dataFileWs.cell(start, j).value is None:
                        makeupWriteColumn = j
                        break
                average = '=ROUND(AVERAGE(' + get_column_letter(makeupWriteColumn) + str(start + 2) + ':' + get_column_letter(makeupWriteColumn) + str(end - 1) + '), 0)'
                dataFileWs.cell(start, makeupWriteColumn).value = date.today()
                dataFileWs.cell(start, makeupWriteColumn).number_format = 'yyyy.mm.dd(aaa)'
                dataFileWs.cell(start, makeupWriteColumn).alignment = Alignment(horizontal='center', vertical='center')
                dataFileWs.cell(start + 1, makeupWriteColumn).value = mockTestName
                dataFileWs.cell(start + 1, makeupWriteColumn).alignment = Alignment(horizontal='center', vertical='center')
                dataFileWs.cell(end, makeupWriteColumn).value = average
                dataFileWs.cell(end, makeupWriteColumn).alignment = Alignment(horizontal='center', vertical='center')
                dataFileWs.cell(end, makeupWriteColumn).border = Border(bottom=Side(border_style='medium', color='000000'))

        if dailyTestScore is not None:
            dataFileWs = dataFileWb['DailyTest']
            testName = dailyTestName
            score = dailyTestScore
            writeColumn = dailyWriteColumn
        elif mockTestScore is not None:
            dataFileWs = dataFileWb['모의고사']
            testName = mockTestName
            score = mockTestScore
            writeColumn = makeupWriteColumn
        else:
            continue # 점수 없으면 미응시 처리
        
        for j in range(start + 2, end):
            if dataFileWs.cell(j, 5).value == formWs.cell(i, 4).value: # data name == form name
                dataFileWs.cell(j, writeColumn).value = score
                dataFileWs.cell(j, writeColumn).alignment = Alignment(horizontal='center', vertical='center')
                break
        
        dataFileWs = dataFileWb['DailyTest']

        if score < 80:
            writeRow = makeupListWs.max_row + 1
            makeupListWs.cell(writeRow, 1).value = date.today()
            makeupListWs.cell(writeRow, 2).value = className
            makeupListWs.cell(writeRow, 3).value = teacher
            makeupListWs.cell(writeRow, 4).value = formWs.cell(i, 4).value
            makeupListWs.cell(writeRow, 5).value = testName
            makeupListWs.cell(writeRow, 6).value = score

    gui.appendLog('재시험 명단 작성 중...')
    for j in range(1, makeupListWs.max_row + 1):
            for k in range(1, makeupListWs.max_column + 1):
                makeupListWs.cell(j, k).alignment = Alignment(horizontal='center', vertical='center')
                makeupListWs.cell(j, k).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    makeupListWb.save('./data/재시험명단.xlsx')

    gui.appendLog('백업 파일 생성중...')
    formWb = xl.load_workbook(filepath)
    formWs = formWb.active
    formWb.save('./data/backup/데일리테스트 기록 양식(' + datetime.today().strftime('%Y%m%d') + ').xlsx')
    # for i in range(2, formWs.max_row + 1):
    #     formWs.cell(i, 6).value = ''
    #     formWs.cell(i, 7).value = ''
    #     formWs.cell(i, 9).value = ''
    #     formWs.cell(i, 10).value = ''
    # formWb.save('./데일리테스트 기록 양식.xlsx')

    dataFileWb.save('./data/' + config['dataFileName'] + '.xlsx')

    # 조건부서식
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(os.getcwd() + '\\data\\' + config['dataFileName'] + '.xlsx')
    wb.Save()
    wb.Close()
    # excel.Quit()

    dataFileWb = xl.load_workbook('./data/' + config['dataFileName'] + '.xlsx')
    dataFileColorWb = xl.load_workbook('./data/' + config['dataFileName'] + '.xlsx', data_only=True)
    sheetNames = dataFileWb.sheetnames
    for i in range(0, 2):
        dataFileWs = dataFileWb[sheetNames[i]]
        dataFileColorWs = dataFileColorWb[sheetNames[i]]
        if i == 0:
            writeColumn = dailyWriteColumn
        else:
            writeColumn = makeupWriteColumn
        
        for j in range(2, dataFileColorWs.max_row+1):
            # 입력 데이터 조건부 서식
            if j > 6:
                dataFileWs.column_dimensions[get_column_letter(j)].width = 13

            if type(dataFileColorWs.cell(j, writeColumn).value) == int:
                if dataFileColorWs.cell(j, writeColumn).value < 60:
                    dataFileWs.cell(j, writeColumn).fill = PatternFill(fill_type='solid', fgColor=Color('EC7E31'))
                elif dataFileColorWs.cell(j, writeColumn).value < 70:
                    dataFileWs.cell(j, writeColumn).fill = PatternFill(fill_type='solid', fgColor=Color('F5AF85'))
                elif dataFileColorWs.cell(j, writeColumn).value < 80:
                    dataFileWs.cell(j, writeColumn).fill = PatternFill(fill_type='solid', fgColor=Color('FCE4D6'))
                elif dataFileColorWs.cell(j, 5).value == '시험 평균':
                    dataFileWs.cell(j, writeColumn).fill = PatternFill(fill_type='solid', fgColor=Color('DDEBF7'))

            # 학생별 평균 조건부 서식
            if type(dataFileColorWs.cell(j, 6).value) == int:
                if dataFileColorWs.cell(j, 6).value < 60:
                    dataFileWs.cell(j, 6).fill = PatternFill(fill_type='solid', fgColor=Color('EC7E31'))
                elif dataFileColorWs.cell(j, 6).value < 70:
                    dataFileWs.cell(j, 6).fill = PatternFill(fill_type='solid', fgColor=Color('F5AF85'))
                elif dataFileColorWs.cell(j, 6).value < 80:
                    dataFileWs.cell(j, 6).fill = PatternFill(fill_type='solid', fgColor=Color('FCE4D6'))
                elif dataFileColorWs.cell(j, 5).value == '시험 평균':
                    dataFileWs.cell(j, 6).fill = PatternFill(fill_type='solid', fgColor=Color('DDEBF7'))
                else:
                    dataFileWs.cell(j, 6).fill = PatternFill(fill_type='solid', fgColor=Color('E2EFDA'))

    dataFileWs = dataFileWb['DailyTest']
    dataFileWb.save('./data/' + config['dataFileName'] + '.xlsx')

    gui.appendLog('데이터 저장을 완료했습니다.')
    gui.saveDataButton['state'] = tk.NORMAL
    excel.Visible = True
    wb = excel.Workbooks.Open(os.getcwd() + '\\data\\' + config['dataFileName'] + '.xlsx')

def classInfo(gui):
    if not os.path.isfile('반 정보.xlsx'):
        gui.appendLog('반 정보 입력 파일 생성 중...')

        iniWb = xl.Workbook()
        iniWs = iniWb.active
        iniWs.title = '데일리테스트 기록 양식'
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
            iniWs.cell(writeLocation, 1).value = tableNames[i].text.split('(')[0].rstrip()

        # 정렬 및 테두리
        for j in range(1, iniWs.max_row + 1):
                for k in range(1, iniWs.max_column + 1):
                    iniWs.cell(j, k).alignment = Alignment(horizontal='center', vertical='center')
                    iniWs.cell(j, k).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        iniWb.save('./반 정보.xlsx')
        gui.classInfoButton['state'] = tk. DISABLED
        gui.appendLog('반 정보 입력 파일 생성을 완료했습니다.')
        gui.appendLog('반 정보를 입력해 주세요.')

    else:
        gui.appendLog('이미 파일이 존재합니다.')
        gui.classInfoButton['state'] = tk. DISABLED

def sendMessage(gui):
    gui.sendMessageButton['state'] = tk.DISABLED
    filepath = filedialog.askopenfilename(initialdir='./', title='데일리테스트 기록 양식 선택', filetypes=(('Excel files', '*.xlsx'),('all files', '*.*')))
    if filepath == '':
        gui.sendMessageButton['state'] = tk.NORMAL
        return

    today = datetime.today()
    
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
        
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=service, options=options)

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

    # 점수기록표 xlsx
    formWb = xl.load_workbook(filepath, data_only=True)
    formWs = formWb.active

    if not os.path.isfile('./재시험 정보.xlsx'):
        gui.appendLog('[오류]재시험 정보.xlsx 파일이 존재하지 않습니다.')
        return
    makeupWb = xl.load_workbook("./재시험 정보.xlsx")
    makeupWs = makeupWb.active

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
        
        if score < 80:
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
                driver.switch_to.window(driver.window_handles[2])
                trs = driver.find_element(By.ID, 'table_' + str(index)).find_elements(By.CLASS_NAME, 'style12')
                for tr in trs:
                    if tr.find_element(By.CLASS_NAME, 'style9').text == name:
                        tds = tr.find_elements(By.TAG_NAME, 'td')
                        tds[0].find_element(By.TAG_NAME, 'input').send_keys(testName)
                        tds[1].find_element(By.TAG_NAME, 'input').send_keys(makeupDate[date].strftime('%m월 %d일') + ' ' + str(time) + '시')
    
    gui.sendMessageButton['state'] = tk.NORMAL
    gui.appendLog('메시지 입력을 완료했습니다.')
    gui.appendLog('메시지 확인 후 전송해주세요.')

def makeupTestInfo(gui):
    if not os.path.isfile('./재시험 정보.xlsx'):
        gui.appendLog('재시험 정보 파일 생성 중...')

        iniWb = xl.Workbook()
        iniWs = iniWb.active
        iniWs.title = 'DailyTest'
        iniWs['A1'] = '이름'
        iniWs['B1'] = '반명'
        iniWs['C1'] = '담당'
        iniWs['D1'] = '요일'
        iniWs['E1'] = '시간'
        iniWs.auto_filter.ref = 'A:C'

        # 반 정보 확인
        if not os.path.isfile('./반 정보.xlsx'):
            gui.appendLog('[오류]반 정보.xlsx 파일이 존재하지 않습니다.')
            return
        classWb = xl.load_workbook("./반 정보.xlsx")
        classWs = classWb.active

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

            className = tableNames[i].text.split('(')[0].rstrip()
            for j in range(2, classWs.max_row + 1):
                if classWs.cell(j, 1).value == className:
                    teacher = classWs.cell(j, 2).value

            # 학생 루프
            for tr in trs:
                writeLocation = iniWs.max_row + 1
                iniWs.cell(writeLocation, 1).value = tr.find_element(By.CLASS_NAME, 'style9').text
                iniWs.cell(writeLocation, 2).value = className
                iniWs.cell(writeLocation, 3).value = teacher

        # 정렬 및 테두리
        for j in range(1, iniWs.max_row + 1):
                for k in range(1, iniWs.max_column + 1):
                    iniWs.cell(j, k).alignment = Alignment(horizontal='center', vertical='center')
                    iniWs.cell(j, k).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        iniWb.save('./재시험 정보.xlsx')
        gui.appendLog('재시험 정보 파일을 생성했습니다.')

    else:
        gui.appendLog('이미 파일이 존재합니다')
        gui.makeDataFileButton['state'] = tk.DISABLED

def applyColor(gui):
    gui.applyColorButton['state'] = tk.DISABLED
    if not os.path.isfile('./data/' + config['dataFileName'] + '.xlsx'):
        gui.appendLog('[오류]'+ config['dataFileName'] +'.xlsx 파일이 존재하지 않습니다.')
        gui.applyColorButton['state'] = tk.NORMAL
        return
    
    gui.appendLog('사용자 서식 적용중...')
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(os.getcwd() + '\\data\\' + config['dataFileName'] + '.xlsx')
    wb.Save()
    wb.Close()

    dataFileWb = xl.load_workbook('./data/' + config['dataFileName'] + '.xlsx')
    dataFileColorWb = xl.load_workbook('./data/' + config['dataFileName'] + '.xlsx', data_only=True)
    for sheetName in dataFileWb.sheetnames:
        dataFileWs = dataFileWb[sheetName]
        dataFileColorWs = dataFileColorWb[sheetName]
        for i in range(2, dataFileColorWs.max_row+1):
            if i > 6:
                dataFileWs.column_dimensions[get_column_letter(i)].width = 14
            if dataFileColorWs.cell(i, 5).value is None:
                break
            for j in range(7, dataFileColorWs.max_column+1):
                if dataFileWs.cell(i, 5).value == '시험 평균' and dataFileWs.cell(i, j).value is not None:
                    dataFileWs.cell(i, j).border = Border(bottom=Side(border_style='medium', color='000000'))
                if dataFileWs.cell(i, 5).value == '날짜' and dataFileWs.cell(i, j).value is not None:
                    dataFileWs.cell(i, j).border = Border(top=Side(border_style='medium', color='000000'))
                # 입력 데이터 조건부 서식
                if type(dataFileColorWs.cell(i, j).value) == int:
                    if dataFileColorWs.cell(i, j).value < 60:
                        dataFileWs.cell(i, j).fill = PatternFill(fill_type='solid', fgColor=Color('EC7E31'))
                    elif dataFileColorWs.cell(i, j).value < 70:
                        dataFileWs.cell(i, j).fill = PatternFill(fill_type='solid', fgColor=Color('F5AF85'))
                    elif dataFileColorWs.cell(i, j).value < 80:
                        dataFileWs.cell(i, j).fill = PatternFill(fill_type='solid', fgColor=Color('FCE4D6'))
                    elif dataFileColorWs.cell(i, 5).value == '시험 평균':
                        dataFileWs.cell(i, j).fill = PatternFill(fill_type='solid', fgColor=Color('DDEBF7'))
                    else:
                        dataFileWs.cell(i, j).fill = PatternFill(fill_type=None, fgColor=Color('00FFFFFF'))


                # 학생별 평균 조건부 서식
                if type(dataFileColorWs.cell(i, 6).value) == int:
                    if dataFileColorWs.cell(i, 6).value < 60:
                        dataFileWs.cell(i, 6).fill = PatternFill(fill_type='solid', fgColor=Color('EC7E31'))
                    elif dataFileColorWs.cell(i, 6).value < 70:
                        dataFileWs.cell(i, 6).fill = PatternFill(fill_type='solid', fgColor=Color('F5AF85'))
                    elif dataFileColorWs.cell(i, 6).value < 80:
                        dataFileWs.cell(i, 6).fill = PatternFill(fill_type='solid', fgColor=Color('FCE4D6'))
                    elif dataFileColorWs.cell(i, 5).value == '시험 평균':
                        dataFileWs.cell(i, 6).fill = PatternFill(fill_type='solid', fgColor=Color('DDEBF7'))
                    else:
                        dataFileWs.cell(i, 6).fill = PatternFill(fill_type='solid', fgColor=Color('E2EFDA'))

    dataFileWb.save('./data/' + config['dataFileName'] + '.xlsx')
    gui.appendLog('사용자 서식 지정을 완료했습니다.')
    gui.applyColorButton['state'] = tk.NORMAL
    excel.Visible = True
    wb = excel.Workbooks.Open(os.getcwd() + '\\data\\' + config['dataFileName'] + '.xlsx')