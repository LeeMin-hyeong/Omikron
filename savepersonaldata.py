import json
import os.path
import openpyxl as xl

from datetime import datetime

config = json.load(open('config.json'))

formWb = xl.load_workbook("test.xlsx", data_only=True)
formWs = formWb.active

className = ''
teacher = ''
testName = ''
average = 0

for i in range(2, formWs.max_row+1):
    if formWs.cell(row = i, column = 2).value is not None:
        className = str(formWs.cell(row = i, column = 2).value)
        teacher = str(formWs.cell(row = i, column = 4).value)
        testName = str(formWs.cell(row = i, column = 5).value)
        average = formWs.cell(row = i, column = 7).value
    # 시험 미응시시 건너뛰기
    if formWs.cell(row = i, column = 6).value is None: continue
    
    # 해당 학생 엑셀 파일이 존재하지 않으면 생성
    if not os.path.isfile(config['dailyTestPersonalFilePath'] + str(formWs.cell(row = i, column = 3).value) + '.xlsx'):
        iniWb = xl.Workbook()
        iniWs = iniWb.active
        iniWs['A1']='응시일'
        iniWs['B1']='반'
        iniWs['C1']='담당T'
        iniWs['D1']='시험명'
        iniWs['E1']='점수'
        iniWs['F1']='반평균'
        iniWb.save(config['dailyTestPersonalFilePath'] + str(formWs.cell(row = i, column = 3).value) + '.xlsx')

    # 해당 학생 파일에 응시 결과 입력
    studentWb = xl.load_workbook(config['dailyTestPersonalFilePath'] + str(formWs.cell(row = i, column = 3).value) + '.xlsx', data_only=True)
    studentWs = studentWb.active
    writeLocation = studentWs.max_row + 1

    #중복방지
    if str(studentWs.cell(row = writeLocation-1, column = 4).value) == testName: continue

    studentWs.cell(row = writeLocation, column = 1, value = datetime.today().strftime('%Y-%m-%d'))
    studentWs.cell(row = writeLocation, column = 2, value = className)
    studentWs.cell(row = writeLocation, column = 3, value = teacher)
    studentWs.cell(row = writeLocation, column = 4, value = testName)
    studentWs.cell(row = writeLocation, column = 5, value = formWs.cell(row = i, column = 6).value)
    studentWs.cell(row = writeLocation, column = 6, value = average)
    studentWb.save(config['dailyTestPersonalFilePath'] + str(formWs.cell(row = i, column = 3).value) + '.xlsx')