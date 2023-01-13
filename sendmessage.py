import time
import json
import openpyxl as xl

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
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

dailyTest = '<오미크론 일일테스트 결과공지>\n안녕하세요 %NAME 학부모님.\n%NAME 학생 당일 일일테스트 (%VAL1) 결과는 %VAL2점 입니다.\n테스트 결과 반 평균은 %VAL3점 입니다.\n감사합니다.'

# 아이소식 key
config = json.load(open('config.json'))

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# 아이소식 접속
driver.get('https://dbserver2.iday-b2.com/pc/')
driver.switch_to.frame('main')
driver.find_element(By.XPATH, '//*[@id="keyid"]').send_keys(config['key'])
driver.find_element(By.XPATH, '//*[@id="keyid"]').send_keys(Keys.ENTER)
driver.find_element(By.XPATH, '//*[@id="keyid"]').send_keys(Keys.ENTER)
driver.find_element(By.XPATH, '//*[@id="firepanel"]').click()
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="leftpanel3"]/div/div[2]/div[8]/a/img').click()
time.sleep(4)
driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(dailyTest)

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
