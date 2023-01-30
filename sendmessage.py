import json
import openpyxl as xl

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service

from webdriver_manager.chrome import ChromeDriverManager

#변수
className = ''
name = ''
teacher = ''
testName = ''
score = 0
average = 0
index = 0

# 아이소식 key url
config = json.load(open('config.json'))

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# 아이소식 접속
driver.get(config['url'])
driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(config['dailyTest'])

# 점수기록표 xlsx 입력
# GUI를 통해 엑셀 파일 받도록 변경 예정
wb = xl.load_workbook("test.xlsx", data_only=True)
ws = wb.active

for i in range(2, ws.max_row+1):
    name = str(ws.cell(row = i, column = 3).value)
    score = ws.cell(row = i, column = 6).value
    if ws.cell(row = i, column = 2).value is not None:
        className = str(ws.cell(row = i, column = 2).value)
        teacher = str(ws.cell(row = i, column = 4).value)
        testName = str(ws.cell(row = i, column = 5).value)
        average = ws.cell(row = i, column = 7).value

    # 시험 미응시시 건너뛰기
    if ws.cell(row = i, column = 6).value is None: continue

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

print('done')
