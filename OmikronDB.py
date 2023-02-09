import json
import os.path
import openpyxl as xl
import win32com.client # only works in Windows

from datetime import datetime
from tkinter import filedialog
from openpyxl.utils.cell import get_column_letter
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
        if not os.path.isfile('./class.xlsx'):
            gui.appendLog('[오류]class.xlsx 파일이 존재하지 않습니다.')
            return
        classWb = xl.load_workbook("./class.xlsx")
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
    iniWs.title = 'DailyTestForm'
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

    if not os.path.isfile('./class.xlsx'):
        gui.appendLog('[오류]class.xlsx 파일이 존재하지 않습니다.')
        return
    
    classWb = xl.load_workbook("class.xlsx")
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
        
    if os.path.isfile('./dailyTestForm.xlsx'):
        i = 1
        while True:
            if not os.path.isfile('./dailyTestForm(' + str(i) +').xlsx'):
                iniWb.save('./dailyTestForm(' + str(i) +').xlsx')
                break;
            i += 1
    else:
        iniWb.save('./dailyTestForm.xlsx')
    gui.appendLog('데일리테스트 기록 양식 생성을 완료했습니다.')

def saveData(gui):
    filepath = filedialog.askopenfilename(initialdir='./', title='데일리테스트 기록 양식 선택', filetypes=(('Excel files', '*.xlsx'),('all files', '*.*')))
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

    today = datetime.today().strftime('%Y.%m.%d')
    # 데이터 날짜 내림차순
    writeColumn = 7
    if str(dataFileWs.cell(1, writeColumn).value) != today:
        dataFileWs.insert_cols(writeColumn)
        for i in range(2, dataFileWs.max_row+1):
            if dataFileWs.cell(i, 5).value == '시험 평균':
                dataFileWs.cell(i, 7).border = Border(bottom=Side(border_style='medium', color='000000'))
    dataFileWs.cell(1, writeColumn).value = today

    for i in range(2, formWs.max_row+1):
        if formWs.cell(i, 9).value is not None:
            dataFileWs = dataFileWb['모의고사']
            if str(dataFileWs.cell(1, writeColumn).value) != today:
                dataFileWs.insert_cols(writeColumn)
                for i in range(2, dataFileWs.max_row+1):
                    if dataFileWs.cell(i, 5).value == '시험 평균':
                        dataFileWs.cell(i, 7).border = Border(bottom=Side(border_style='medium', color='000000'))
            dataFileWs.cell(1, writeColumn).value = today
            dataFileWs = dataFileWb['DailyTest']
            break

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
        
            for j in range(2, dataFileWs.max_row+1):
                if dataFileWs.cell(j, 3).value == className: # data className == form className
                    start = j + 1
                    break
            
            for j in range(start, dataFileWs.max_row+1):
                if dataFileWs.cell(j, 5).value == '시험 평균': # data name is 시험 평균
                    end = j - 1
                    break
            average = '=ROUND(AVERAGE(' + get_column_letter(writeColumn) + str(start) + ':' + get_column_letter(writeColumn) + str(end) + '), 0)'
            
            if dailyTestName is not None:
                dataFileWs.cell(start - 1, writeColumn).value = dailyTestName
                dataFileWs.cell(end + 1, writeColumn).value = average
            if mockTestName is not None:
                dataFileWs = dataFileWb['모의고사']
                dataFileWs.cell(start - 1, writeColumn).value = mockTestName
                dataFileWs.cell(end + 1, writeColumn).value = average
                dataFileWs = dataFileWb['DailyTest']

        if mockTestScore is not None:
            dataFileWs = dataFileWb['모의고사']
            testName = mockTestName
            score = mockTestScore
        elif dailyTestScore is not None:
            testName = dailyTestName
            score = dailyTestScore
        else:
            continue
        
        for j in range(start, end):
            if dataFileWs.cell(j, 5).value == formWs.cell(i, 4).value: # data name == form name
                dataFileWs.cell(j, writeColumn).value = score
                break
        
        dataFileWs = dataFileWb['DailyTest']

        if score < 80:
            writeRow = makeupListWs.max_row + 1
            makeupListWs.cell(writeRow, 1).value = today
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
    formWb.save('./data/backup/dailyTestForm(' + datetime.today().strftime('%Y%m%d') + ').xlsx')
    # for i in range(2, formWs.max_row + 1):
    #     formWs.cell(i, 6).value = ''
    #     formWs.cell(i, 7).value = ''
    #     formWs.cell(i, 9).value = ''
    #     formWs.cell(i, 10).value = ''
    # formWb.save('./dailyTestForm.xlsx')

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
    for sheetName in dataFileWb.sheetnames:
        dataFileWs = dataFileWb[sheetName]
        dataFileColorWs = dataFileColorWb[sheetName]
        for i in range(2, dataFileColorWs.max_row+1):
            # 입력 데이터 조건부 서식
            if type(dataFileColorWs.cell(i, writeColumn).value) == int:
                if dataFileColorWs.cell(i, writeColumn).value < 60:
                    dataFileWs.cell(i, writeColumn).fill = PatternFill(fill_type='solid', fgColor=Color('EC7E31'))
                elif dataFileColorWs.cell(i, writeColumn).value < 70:
                    dataFileWs.cell(i, writeColumn).fill = PatternFill(fill_type='solid', fgColor=Color('F5AF85'))
                elif dataFileColorWs.cell(i, writeColumn).value < 80:
                    dataFileWs.cell(i, writeColumn).fill = PatternFill(fill_type='solid', fgColor=Color('FCE4D6'))
                elif dataFileColorWs.cell(i, 5).value == '시험 평균':
                    dataFileWs.cell(i, 7).fill = PatternFill(fill_type='solid', fgColor=Color('DDEBF7'))

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

    gui.appendLog('데이터 저장을 완료했습니다.')
    excel.Visible = True
    wb = excel.Workbooks.Open(os.getcwd() + '\\data\\' + config['dataFileName'] + '.xlsx')

def classInfo(gui):
    if not os.path.isfile('class.xlsx'):
        gui.appendLog('반 정보 입력 파일 생성 중...')

        iniWb = xl.Workbook()
        iniWs = iniWb.active
        iniWs.title = 'DailyTestForm'
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

        iniWb.save('./class.xlsx')
        gui.appendLog('반 정보 입력 파일 생성을 완료했습니다.')
        gui.appendLog('반 정보를 입력해 주세요.')

    else:
        gui.appendLog('이미 파일이 존재합니다.')

def sendMessage(gui):
    filepath = filedialog.askopenfilename(initialdir='./', title='데일리테스트 기록 양식 선택', filetypes=(('Excel files', '*.xlsx'),('all files', '*.*')))
    
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=service, options=options)

    # 아이소식 접속
    driver.get(config['url'])
    driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(config['dailyTest'])

    # 점수기록표 xlsx 입력
    formWb = xl.load_workbook(filepath, data_only=True)
    formWs = formWb.active

    for i in range(2, formWs.max_row+1):
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
        for i in range(len(tableNames)):
            if className in tableNames[i].text:
                index = i
                break

        trs = driver.find_element(By.ID, 'table_' + str(index)).find_elements(By.CLASS_NAME, 'style12')
        for tr in trs:
            if tr.find_element(By.CLASS_NAME, 'style9').text == name:
                tds = tr.find_elements(By.TAG_NAME, 'td')
                tds[0].find_element(By.TAG_NAME, 'input').send_keys(testName)
                tds[1].find_element(By.TAG_NAME, 'input').send_keys(score)
                tds[2].find_element(By.TAG_NAME, 'input').send_keys(average)
                break
    gui.appendLog('메시지 입력을 완료했습니다.')
    gui.appendLog('메시지 확인 후 전송해주세요.')