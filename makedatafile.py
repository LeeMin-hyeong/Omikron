import json
import os.path
import openpyxl as xl

from openpyxl.styles import Border, Side

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service

from webdriver_manager.chrome import ChromeDriverManager

index = 0

config = json.load(open('config.json', encoding='UTF8'))

if True:
# if not os.path.isfile('./data/'+config['dailyTestFileName']):
    print('Processing...')

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

    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    driver = webdriver.Chrome(service = Service(ChromeDriverManager().install()), options = options)

    # 아이소식 접속
    driver.get(config['url'])
    tableNames = driver.find_elements(By.CLASS_NAME, 'style1')

    #반 루프
    for i in range(3, len(tableNames)):
        trs = driver.find_element(By.ID, 'table_' + str(i)).find_elements(By.CLASS_NAME, 'style12')
        writeLocation = iniWs.max_row + 1

        className = tableNames[i].text.split('(')[0].rstrip()
        iniWs.cell(row = writeLocation, column = 3, value = className)
        iniWs.cell(row = writeLocation, column = 5, value = '시험명')

        #학생 루프
        for tr in trs:
            writeLocation = iniWs.max_row + 1
            iniWs.cell(row = writeLocation, column = 3, value = className)
            iniWs.cell(row = writeLocation, column = 5, value = tr.find_element(By.CLASS_NAME, 'style9').text)
            iniWs.cell(row = writeLocation, column = 6, value = '=ROUND(AVERAGE(G' + str(writeLocation) + ':XFD' + str(writeLocation) + '), 0)')
        
        writeLocation = iniWs.max_row + 1
        iniWs.cell(row = writeLocation, column = 3, value = className)
        iniWs.cell(row = writeLocation, column = 5, value = '시험 평균')
        iniWs.cell(row = writeLocation, column = 6, value = '=ROUND(AVERAGE(G' + str(writeLocation) + ':XFD' + str(writeLocation) + '), 0)')

        for j in range(1, 7):
            iniWs.cell(row = writeLocation, column = j).border = Border(bottom = Side(border_style='medium', color='000000'))



    #모의고사 sheet 생성
    copyWs = iniWb.copy_worksheet(iniWb['DailyTest'])
    copyWs.title = '모의고사'
    copyWs.freeze_panes = 'G2'
    copyWs.auto_filter.ref = 'A:E'

    iniWb.save('./data/'+config['dailyTestFileName'])
    print('done')

else:
    print('이미 파일이 존재합니다')