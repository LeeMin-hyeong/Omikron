from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from win32process import CREATE_NO_WINDOW # only works in Windows
from webdriver_manager.chrome import ChromeDriverManager

import omikron.dataform
import omikron.studentinfo

from omikron.config import URL, TEST_RESULT_MESSAGE, MAKEUP_TEST_NO_SCHEDULE_MESSAGE, MAKEUP_TEST_SCHEDULE_MESSAGE
from omikron.defs import Chrome, DataForm
from omikron.errorui import chrome_driver_version_error
from omikron.log import OmikronLog
from omikron.util import calculate_makeup_test_schedule, date_to_kor_date

try:
    service = Service(ChromeDriverManager().install().replace("THIRD_PARTY_NOTICES.chromedriver", "chromedriver.exe"))
    service.creation_flags = CREATE_NO_WINDOW
    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    test = webdriver.Chrome(service = service, options = options)
    test.get("about:blank")
    test.quit()
except:
    chrome_driver_version_error()

# 크롬 기본 작업
def open_web_background() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    driver = webdriver.Chrome(service = service, options = options)

    # 아이소식 접속
    driver.get(URL)

    return driver

def close_web_background(driver:webdriver.Chrome):
    driver.quit()

# 크롬 유틸리티
def get_class_names() -> list[str]:
    """
    실제 반 정보를 담고 있는 테이블부터 모든 반 이름 리스트를 생성
    """
    driver = open_web_background()

    table_names = driver.find_elements(By.CLASS_NAME, "style1")[Chrome.ACTUAL_CLASS_START_INDEX:]
    class_names = [table_name.text.strip() for table_name in table_names]

    close_web_background(driver)

    return class_names

def get_student_names() -> list[str]:
    """
    실제 반 정보를 담고 있는 테이블부터 모든 학생의 이름 리스트를 생성

    중복 제거된 리스트
    """
    student_names = []

    driver = open_web_background()

    table_counts = len(driver.find_elements(By.CLASS_NAME, "style1"))
    for i in range(Chrome.ACTUAL_CLASS_START_INDEX, table_counts):
        trs = driver.find_element(By.ID, f"table_{i}").find_elements(By.CLASS_NAME, "style12")
        for tr in trs:
            student_names.append(tr.find_element(By.CLASS_NAME, "style9").text.strip())

    close_web_background(driver)

    return sorted(list(set(student_names)))

def get_class_student_dict() -> dict[str:list[str]]:
    """
    실제 반 정보를 담고 있는 테이블부터 '반 : 학생 리스트'의 dict 생성
    """
    class_student_dict = {}

    driver = open_web_background()

    table_names = driver.find_elements(By.CLASS_NAME, "style1")[Chrome.ACTUAL_CLASS_START_INDEX:]
    for i, table_name in enumerate(table_names, start=Chrome.ACTUAL_CLASS_START_INDEX):
        trs = driver.find_element(By.ID, f"table_{i}").find_elements(By.CLASS_NAME, "style12")
        student_list = [tr.find_element(By.CLASS_NAME, "style9").text.strip() for tr in trs]

        class_student_dict[table_name.text.strip()] = student_list

    close_web_background(driver)

    return class_student_dict

def check_student_exists(student_name:str, target_class_name:str) -> bool:
    """
    아이소식의 스크립트 정보에 기반하여

    특정 반에 특정 학생이 존재하는 지 확인
    """
    driver = open_web_background()

    table_names = driver.find_elements(By.CLASS_NAME, "style1")[Chrome.ACTUAL_CLASS_START_INDEX:]
    for i, table_name in enumerate(table_names, start=Chrome.ACTUAL_CLASS_START_INDEX):
        if target_class_name == table_name.text.strip():
            trs = driver.find_element(By.ID, f"table_{i}").find_elements(By.CLASS_NAME, "style12")
            for tr in trs:
                if student_name == tr.find_element(By.CLASS_NAME, "style9").text.strip():
                    close_web_background(driver)
                    return True

    close_web_background(driver)

    return False

# 크롬 작업
def send_test_result_message(filepath:str, makeup_test_date:dict) -> bool:
    """
    기록 양식의 데이터를 추출하여 아이소식 스크립트 작성
    """
    form_wb = omikron.dataform.open(filepath)
    complete, form_ws = omikron.dataform.open_worksheet(form_wb)
    if not complete: return False

    student_wb = omikron.studentinfo.open()
    complete, student_ws = omikron.studentinfo.open_worksheet(student_wb)
    if not complete: return False

    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=service, options=options)
    
    # 아이소식 접속
    driver.get(URL)
    driver.execute_script(f"arguments[0].value = '{TEST_RESULT_MESSAGE}'", driver.find_element(By.XPATH, '//*[@id="ctitle"]'))
    driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(' \b')

    driver.execute_script(f"window.open('{URL}')")
    driver.switch_to.window(driver.window_handles[Chrome.MAKEUPTEST_NO_SCHEDULE_TAB])
    driver.execute_script(f"arguments[0].value = '{MAKEUP_TEST_NO_SCHEDULE_MESSAGE}'", driver.find_element(By.XPATH, '//*[@id="ctitle"]'))
    driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(' \b')

    driver.execute_script(f"window.open('{URL}')")
    driver.switch_to.window(driver.window_handles[Chrome.MAKEUPTEST_SCHEDULE_TAB])
    driver.execute_script(f"arguments[0].value = '{MAKEUP_TEST_SCHEDULE_MESSAGE}'", driver.find_element(By.XPATH, '//*[@id="ctitle"]'))
    driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(' \b')

    driver.switch_to.window(driver.window_handles[Chrome.DAILYTEST_RESULT_TAB])

    # 반 인덱스 dict
    table_names = driver.find_elements(By.CLASS_NAME, "style1")
    table_index_dict = {table_name.text.strip() : i for i, table_name in enumerate(table_names)}

    for i in range(2, form_ws.max_row + 1):
        # 반 필터링
        if form_ws.cell(i, DataForm.CLASS_NAME_COLUMN).value is not None:
            class_name         = form_ws.cell(i, DataForm.CLASS_NAME_COLUMN).value
            daily_test_name    = form_ws.cell(i, DataForm.DAILYTEST_NAME_COLUMN).value
            mock_test_name     = form_ws.cell(i, DataForm.MOCKTEST_NAME_COLUMN).value
            daily_test_average = form_ws.cell(i, DataForm.DAILYTEST_AVERAGE_COLUMN).value
            mock_test_average  = form_ws.cell(i, DataForm.MOCKTEST_AVERAGE_COLUMN).value

            # 반 전체가 시험을 응시하지 않은 경우
            keep_continue = False
            if daily_test_name is None and mock_test_name is None:
                keep_continue = True
                continue

            # 테이블 인덱스
            class_index = table_index_dict[class_name]

            # 학생 인덱스 dict
            trs = driver.find_element(By.ID, f"table_{str(class_index)}").find_elements(By.CLASS_NAME, "style12")
            student_index_dict = {tr.find_element(By.CLASS_NAME, "style9").text.strip() : i for i, tr in enumerate(trs)}

        # 반 전체가 시험을 응시하지 않은 경우
        if keep_continue:
            continue

        student_name     = form_ws.cell(i, DataForm.STUDENT_NAME_COLUMN).value
        daily_test_score = form_ws.cell(i, DataForm.DAILYTEST_SCORE_COLUMN).value
        mock_test_score  = form_ws.cell(i, DataForm.MOCKTEST_SCORE_COLUMN).value

        # 시험 미응시 시 건너뛰기
        if daily_test_score is not None:
            test_name    = daily_test_name
            test_score   = daily_test_score
            test_average = daily_test_average
        elif mock_test_score is not None:
            test_name    = mock_test_name
            test_score   = mock_test_score
            test_average = mock_test_average
        else:
            continue

        if type(test_score) not in (int, float):
            continue

        # 시험 결과 메시지 작성
        driver.switch_to.window(driver.window_handles[Chrome.DAILYTEST_RESULT_TAB])

        try:
            student_index = student_index_dict[student_name]
        except KeyError:
            OmikronLog.warning(f"아이소식의 {class_name} 내 {student_name} 학생이 존재하지 않습니다.")
            continue

        trs = driver.find_element(By.ID, f"table_{str(class_index)}").find_elements(By.CLASS_NAME, "style12")
        tds = trs[student_index].find_elements(By.TAG_NAME, "td")
        driver.execute_script(f"arguments[0].value = '{test_name}'",  tds[0].find_element(By.TAG_NAME, "input"))
        driver.execute_script(f"arguments[0].value = '{test_score}'", tds[1].find_element(By.TAG_NAME, "input"))
        tds[2].find_element(By.TAG_NAME, "input").send_keys(test_average)

        # 재시험 메시지 작성
        if test_score >= 80:
            continue
        if form_ws.cell(i, DataForm.MAKEUP_TEST_CHECK_COLUMN).value in ("x", "X"):
            continue

        # 학생 정보 검색
        info_exists, makeup_test_weekday, makeup_test_time, _ = omikron.studentinfo.get_student_info(student_ws, student_name)
        if not info_exists:
            OmikronLog.warning(f"{student_name}의 학생 정보가 존재하지 않습니다.")

        if info_exists and makeup_test_weekday is not None:
            # 재시험 일정 계산
            complete, calculated_schedule, time_index = calculate_makeup_test_schedule(makeup_test_weekday, makeup_test_date)
            if complete:
                # 재시험 일정 계산 성공
                calculated_schedule_str = date_to_kor_date(calculated_schedule)

                driver.switch_to.window(driver.window_handles[Chrome.MAKEUPTEST_SCHEDULE_TAB])
                trs = driver.find_element(By.ID, f"table_{str(class_index)}").find_elements(By.CLASS_NAME, "style12")
                tds = trs[student_index].find_elements(By.TAG_NAME, "td")
                driver.execute_script(f"arguments[0].value = '{test_name}'", tds[0].find_element(By.TAG_NAME, "input"))

                if makeup_test_time is not None:
                    # 날짜 지정 / 시간 지정
                    if "/" in str(makeup_test_time):
                        # 다중 재시험 시간
                        if len(makeup_test_weekday.split("/")) == len(makeup_test_time.split("/")):
                            # 재시험 요일 의 구분 개수는 재시험 시간의 구분 개수와 동일
                            driver.execute_script(f"arguments[0].value = '{calculated_schedule_str} {str(makeup_test_time).split('/')[time_index]}시'", tds[1].find_element(By.TAG_NAME, "input"))
                            tds[2].find_element(By.TAG_NAME, "input").send_keys(' \b')

                            continue
                        else:
                            OmikronLog.warning(f"{student_name}의 재시험 시간이 올바른 양식이 아닙니다.")
                    else:
                        # 단일 재시험 시간
                        driver.execute_script(f"arguments[0].value = '{calculated_schedule_str} {str(makeup_test_time)}시'", tds[1].find_element(By.TAG_NAME, "input"))
                        tds[2].find_element(By.TAG_NAME, "input").send_keys(' \b')

                        continue

                # 날짜 지정 / 시간 미지정
                driver.execute_script(f"arguments[0].value = '{calculated_schedule_str}'", tds[1].find_element(By.TAG_NAME, "input"))
                tds[2].find_element(By.TAG_NAME, "input").send_keys(' \b')

                continue
            else:
                # 재시험 일정 계산 중 오류
                OmikronLog.warning(f"{student_name}의 재시험 요일이 올바른 양식이 아닙니다.")

        # 재시험 일정 없음
        driver.switch_to.window(driver.window_handles[Chrome.MAKEUPTEST_NO_SCHEDULE_TAB])

        trs = driver.find_element(By.ID, f"table_{str(class_index)}").find_elements(By.CLASS_NAME, "style12")
        tds = trs[student_index].find_elements(By.TAG_NAME, "td")
        driver.execute_script(f"arguments[0].value = '{test_name}'", tds[0].find_element(By.TAG_NAME, "input"))

        tds[1].find_element(By.TAG_NAME, "input").send_keys(' \b')

    return True

def send_individual_test_message(student_name:str, class_name:int, test_name:int, test_score:int, test_average:int, makeup_test_check:bool, makeup_test_date:dict) -> bool:
    """
    개별 시험에 대한 결과 메시지 전송
    """
    student_wb = omikron.studentinfo.open()
    complete, student_ws = omikron.studentinfo.open_worksheet(student_wb)
    if not complete: return False

    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=service, options=options)
    
    # 아이소식 접속
    driver.get(URL)
    driver.execute_script(f"arguments[0].value = '{TEST_RESULT_MESSAGE}'", driver.find_element(By.XPATH, '//*[@id="ctitle"]'))
    driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(' \b')

    # 반 인덱스 dict
    table_names = driver.find_elements(By.CLASS_NAME, "style1")
    table_index_dict = {table_name.text.strip() : i for i, table_name in enumerate(table_names)}

    class_index = table_index_dict[class_name]

    # 학생 인덱스 dict
    trs = driver.find_element(By.ID, f"table_{str(class_index)}").find_elements(By.CLASS_NAME, "style12")
    student_index_dict = {tr.find_element(By.CLASS_NAME, "style9").text.strip() : i for i, tr in enumerate(trs)}
    try:
        student_index = student_index_dict[student_name]
    except KeyError:
        OmikronLog.warning(f"아이소식의 {class_name} 내 {student_name} 학생이 존재하지 않습니다.")
        return False

    trs = driver.find_element(By.ID, f"table_{str(class_index)}").find_elements(By.CLASS_NAME, "style12")
    tds = trs[student_index].find_elements(By.TAG_NAME, "td")
    driver.execute_script(f"arguments[0].value = '{test_name}'",  tds[0].find_element(By.TAG_NAME, "input"))
    driver.execute_script(f"arguments[0].value = '{test_score}'", tds[1].find_element(By.TAG_NAME, "input"))
    tds[2].find_element(By.TAG_NAME, "input").send_keys(test_average)

    if test_score >= 80:
        return True
    if makeup_test_check:
        return True

    driver.execute_script(f"window.open('{URL}')")
    driver.switch_to.window(driver.window_handles[Chrome.INDIVIDUAL_MAKEUPTEST_TAB])

    # 학생 정보 검색
    info_exists, makeup_test_weekday, makeup_test_time, _ = omikron.studentinfo.get_student_info(student_ws, student_name)
    if not info_exists:
        OmikronLog.warning(f"{student_name}의 학생 정보가 존재하지 않습니다.")

    if info_exists and makeup_test_weekday is not None:
        # 재시험 일정 계산
        complete, calculated_schedule, time_index = calculate_makeup_test_schedule(makeup_test_weekday, makeup_test_date)
        if complete:
            # 재시험 일정 계산 성공
            driver.execute_script(f"arguments[0].value = '{MAKEUP_TEST_SCHEDULE_MESSAGE}'", driver.find_element(By.XPATH, '//*[@id="ctitle"]'))
            driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(' \b')
            calculated_schedule_str = date_to_kor_date(calculated_schedule)

            trs = driver.find_element(By.ID, f"table_{str(class_index)}").find_elements(By.CLASS_NAME, "style12")
            tds = trs[student_index].find_elements(By.TAG_NAME, "td")
            driver.execute_script(f"arguments[0].value = '{test_name}'", tds[0].find_element(By.TAG_NAME, "input"))

            if makeup_test_time is not None:
                if "/" in str(makeup_test_time):
                    # 다중 재시험 시간
                    if len(makeup_test_weekday.split("/")) == len(makeup_test_time.split("/")):
                        # 재시험 요일 의 구분 개수는 재시험 시간의 구분 개수와 동일
                        driver.execute_script(f"arguments[0].value = '{calculated_schedule_str} {str(makeup_test_time).split('/')[time_index]}시'", tds[1].find_element(By.TAG_NAME, "input"))
                        tds[2].find_element(By.TAG_NAME, "input").send_keys(' \b')

                        driver.switch_to.window(driver.window_handles[Chrome.DAILYTEST_RESULT_TAB])

                        return True
                    else:
                        OmikronLog.warning(f"{student_name}의 재시험 시간이 올바른 양식이 아닙니다.")

                else:
                    # 단일 재시험 시간
                    driver.execute_script(f"arguments[0].value = '{calculated_schedule_str} {str(makeup_test_time)}시'", tds[1].find_element(By.TAG_NAME, "input"))
                    tds[2].find_element(By.TAG_NAME, "input").send_keys(' \b')

                    driver.switch_to.window(driver.window_handles[Chrome.DAILYTEST_RESULT_TAB])

                    return True

            # 날짜 지정 / 시간 미지정
            driver.execute_script(f"arguments[0].value = '{calculated_schedule_str}'", tds[1].find_element(By.TAG_NAME, "input"))
            tds[2].find_element(By.TAG_NAME, "input").send_keys(' \b')

            driver.switch_to.window(driver.window_handles[Chrome.DAILYTEST_RESULT_TAB])

            return True
        else:
            # 재시험 일정 계산 중 오류
            OmikronLog.warning(f"{student_name}의 재시험 요일이 올바른 양식이 아닙니다.")

    # 재시험 일정 없음
    driver.execute_script(f"arguments[0].value = '{MAKEUP_TEST_NO_SCHEDULE_MESSAGE}'", driver.find_element(By.XPATH, '//*[@id="ctitle"]'))
    driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(' \b')
    trs = driver.find_element(By.ID, f"table_{str(class_index)}").find_elements(By.CLASS_NAME, "style12")
    tds = trs[student_index].find_elements(By.TAG_NAME, "td")
    driver.execute_script(f"arguments[0].value = '{test_name}'", tds[0].find_element(By.TAG_NAME, "input"))

    tds[1].find_element(By.TAG_NAME, "input").send_keys(' \b')

    driver.switch_to.window(driver.window_handles[Chrome.DAILYTEST_RESULT_TAB])

    return True
