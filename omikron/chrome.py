from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from win32process import CREATE_NO_WINDOW # only works in Windows

import requests
from bs4 import BeautifulSoup

import omikron.dataform
import omikron.studentinfo

from omikron.config import URL, TEST_RESULT_MESSAGE, MAKEUP_TEST_NO_SCHEDULE_MESSAGE, MAKEUP_TEST_SCHEDULE_MESSAGE
from omikron.defs import Chrome, DataForm
from omikron.util import calculate_makeup_test_schedule, date_to_kor_date
from omikron.progress import Progress

def _fetch_aisosik_soup() -> BeautifulSoup:
    """
    아이소식 페이지 HTML을 가져와 BeautifulSoup로 반환.
    - 로그인/쿠키가 필요한 페이지면, 여기에서 세션/쿠키 처리하도록 확장하면 됨.
    """
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
    }

    with requests.Session() as s:
        r = s.get(URL, headers=headers, timeout=10)
        r.raise_for_status()

        # 인코딩이 애매한 사이트면 아래 라인이 도움 될 수 있음
        # r.encoding = r.apparent_encoding

        return BeautifulSoup(r.text, "html.parser")

# 크롬 유틸리티
def get_class_names() -> list[str]:
    """
    실제 반 정보를 담고 있는 테이블부터 모든 반 이름 리스트를 생성
    """
    soup = _fetch_aisosik_soup()

    elems = soup.select(".style1")[Chrome.ACTUAL_CLASS_START_INDEX:]
    class_names = [e.get_text(strip=True) for e in elems]
    # 빈 문자열 제거
    return [name for name in class_names if name]

def get_student_names() -> list[str]:
    """
    실제 반 정보를 담고 있는 테이블부터 모든 학생의 이름 리스트를 생성 (중복 제거)
    """
    soup = _fetch_aisosik_soup()

    # Selenium 코드에서는 style1 개수로 table_{i} 범위를 잡았음
    table_count = len(soup.select(".style1"))

    student_set: set[str] = set()

    for i in range(Chrome.ACTUAL_CLASS_START_INDEX, table_count):
        table = soup.find(id=f"table_{i}")
        if table is None:
            continue

        # table 안에서 style12 행들 찾고, 각 행에서 style9(이름) 텍스트 추출
        for tr in table.select(".style12"):
            name_el = tr.select_one(".style9")
            if not name_el:
                continue
            name = name_el.get_text(strip=True)
            if name:
                student_set.add(name)

    return sorted(student_set)

def get_class_student_dict() -> dict[str, list[str]]:
    """
    실제 반 정보를 담고 있는 테이블부터 '반 : 학생 리스트' dict 생성
    """
    soup = _fetch_aisosik_soup()

    class_student_dict: dict[str, list[str]] = {}

    table_names = soup.select(".style1")[Chrome.ACTUAL_CLASS_START_INDEX:]
    for offset, table_name_el in enumerate(table_names):
        i = Chrome.ACTUAL_CLASS_START_INDEX + offset

        class_name = table_name_el.get_text(strip=True)
        if not class_name:
            continue

        table = soup.find(id=f"table_{i}")
        if table is None:
            class_student_dict[class_name] = []
            continue

        student_list: list[str] = []
        for tr in table.select(".style12"):
            name_el = tr.select_one(".style9")
            if not name_el:
                continue
            name = name_el.get_text(strip=True)
            if name:
                student_list.append(name)

        class_student_dict[class_name] = student_list

    return class_student_dict

def check_student_exists(student_name: str, target_class_name: str) -> bool:
    """
    특정 반에 특정 학생이 존재하는지 확인
    """
    soup = _fetch_aisosik_soup()

    table_names = soup.select(".style1")[Chrome.ACTUAL_CLASS_START_INDEX:]

    for offset, table_name_el in enumerate(table_names):
        class_name = table_name_el.get_text(strip=True)
        if class_name != target_class_name:
            continue

        i = Chrome.ACTUAL_CLASS_START_INDEX + offset
        table = soup.find(id=f"table_{i}")
        if table is None:
            return False

        # 같은 반 테이블에서 학생 이름만 검사
        for tr in table.select(".style12"):
            name_el = tr.select_one(".style9")
            if not name_el:
                continue
            if name_el.get_text(strip=True) == student_name:
                return True

        return False

    return False

# 크롬 작업
def _set_input(driver, input_el, value):
    driver.execute_script("arguments[0].value = arguments[1]", input_el, str(value))

def _set_value_with_events(driver, el, value):
    driver.execute_script("""
        const el = arguments[0];
        const val = arguments[1];
        el.focus();
        el.value = val;

        // React/Vue 같은 경우 input 이벤트가 핵심인 경우 많음
        el.dispatchEvent(new Event('input',  { bubbles: true }));
        // 폼 검증/계산 트리거가 change에 걸린 경우도 많음
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.blur();
    """, el, value)

def _cache_table_inputs(driver:webdriver.Chrome, class_index: int) -> dict[str, tuple]:
    """
    table_{class_index}에서
    학생이름 -> (시험명 input, 점수 input, 평균 input) 캐싱
    """
    table = driver.find_element(By.ID, f"table_{class_index}")
    rows = table.find_elements(By.CLASS_NAME, "style12")

    name_to_inputs = {}
    for row in rows:
        name = row.find_element(By.CLASS_NAME, "style9").text.strip()
        if not name:
            continue

        tds = row.find_elements(By.TAG_NAME, "td")
        in0 = tds[0].find_element(By.TAG_NAME, "input")
        in1 = tds[1].find_element(By.TAG_NAME, "input")
        in2 = tds[2].find_element(By.TAG_NAME, "input")
        name_to_inputs[name] = (in0, in1, in2)

    return name_to_inputs

def send_test_result_message(filepath:str, makeup_test_date:dict, prog:Progress) -> bool:
    """
    기록 양식의 데이터를 추출하여 아이소식 스크립트 작성
    """
    try:
        service = Service()
        service.creation_flags = CREATE_NO_WINDOW
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--blink-settings=imagesEnabled=false")
        options.page_load_strategy = "eager"
        options.add_experimental_option("detach", True)

        form_wb = omikron.dataform.open(filepath)
        form_ws = omikron.dataform.open_worksheet(form_wb)

        student_wb = omikron.studentinfo.open()
        student_ws = omikron.studentinfo.open_worksheet(student_wb)

        driver = webdriver.Chrome(service=service, options=options)
        
        # 아이소식 접속
        driver.get(URL)
        driver.execute_script(f"""
                              arguments[0].value = '{TEST_RESULT_MESSAGE}';
                              const el = arguments[0];
                              el.dispatchEvent(new Event('input',  {"{ bubbles: true }"}));
                              el.dispatchEvent(new Event('change', {"{ bubbles: true }"}));
                              el.blur();
                              document.title = '시험 결과 전송';""", 
                              driver.find_element(By.XPATH, '//*[@id="ctitle"]')
                              )

        driver.execute_script(f"window.open('{URL}')")
        driver.switch_to.window(driver.window_handles[Chrome.MAKEUPTEST_NO_SCHEDULE_TAB])
        driver.execute_script(f"""
                              arguments[0].value = '{MAKEUP_TEST_NO_SCHEDULE_MESSAGE}';
                              const el = arguments[0];
                              el.dispatchEvent(new Event('input',  {"{ bubbles: true }"}));
                              el.dispatchEvent(new Event('change', {"{ bubbles: true }"}));
                              el.blur();
                              document.title = '재시험 일정 없는 학생';""", 
                              driver.find_element(By.XPATH, '//*[@id="ctitle"]')
                              )

        driver.execute_script(f"window.open('{URL}')")
        driver.switch_to.window(driver.window_handles[Chrome.MAKEUPTEST_SCHEDULE_TAB])
        driver.execute_script(f"""
                              arguments[0].value = '{MAKEUP_TEST_SCHEDULE_MESSAGE}';
                              const el = arguments[0];
                              el.dispatchEvent(new Event('input',  {"{ bubbles: true }"}));
                              el.dispatchEvent(new Event('change', {"{ bubbles: true }"}));
                              el.blur();
                              document.title = '재시험 일정 있는 학생';""", 
                              driver.find_element(By.XPATH, '//*[@id="ctitle"]')
                              )

        driver.switch_to.window(driver.window_handles[Chrome.DAILYTEST_RESULT_TAB])

        # 반 인덱스 dict
        soup = BeautifulSoup(driver.page_source, "html.parser")
        names = [el.get_text(strip=True) for el in soup.select(".style1")]
        table_index_dict = {name: i for i, name in enumerate(names) if name}
        # table_names = driver.find_elements(By.CLASS_NAME, "style1")
        # table_index_dict = {table_name.text.strip() : i for i, table_name in enumerate(table_names)}

        # 탭별 캐시: class_index -> (student_name -> inputs)
        daily_cache: dict[int, dict[str, tuple]] = {}
        nosched_cache: dict[int, dict[str, tuple]] = {}
        sched_cache: dict[int, dict[str, tuple]] = {}

        # 탭별 작업 큐
        daily_ops = []
        nosched_ops = []
        sched_ops = []

        # 루프에서 매 행마다 DOM 조작하지 말고 "작업만 수집"
        class_index = None
        class_name = None
        daily_test_name = mock_test_name = None
        daily_test_average = mock_test_average = None

        for row in range(2, form_ws.max_row + 1):
            if form_ws.cell(row, DataForm.CLASS_NAME_COLUMN).value is not None:
                class_name = str(form_ws.cell(row, DataForm.CLASS_NAME_COLUMN).value)
                daily_test_name    = str(form_ws.cell(row, DataForm.DAILYTEST_NAME_COLUMN).value)
                mock_test_name     = str(form_ws.cell(row, DataForm.MOCKTEST_NAME_COLUMN).value)
                daily_test_average = str(form_ws.cell(row, DataForm.DAILYTEST_AVERAGE_COLUMN).value)
                mock_test_average  = str(form_ws.cell(row, DataForm.MOCKTEST_AVERAGE_COLUMN).value)

                if daily_test_name is None and mock_test_name is None:
                    continue

                class_index = table_index_dict.get(class_name)
                if class_index is None:
                    prog.warning(f"아이소식에 {class_name} 반이 존재하지 않습니다.")
                    continue

            student_name     = form_ws.cell(row, DataForm.STUDENT_NAME_COLUMN).value
            daily_test_score = form_ws.cell(row, DataForm.DAILYTEST_SCORE_COLUMN).value
            mock_test_score  = form_ws.cell(row, DataForm.MOCKTEST_SCORE_COLUMN).value

            if daily_test_score is not None:
                test_name, test_score, test_average = daily_test_name, daily_test_score, daily_test_average
            elif mock_test_score is not None:
                test_name, test_score, test_average = mock_test_name, mock_test_score, mock_test_average
            else:
                continue

            if type(test_score) not in (int, float):
                continue

            daily_ops.append((class_index, student_name, test_name, test_score, test_average))

            # 재시험 분기(여기서는 DOM 안 건드리고 “어느 탭에 쓸지”만 결정)
            if test_score >= 80:
                continue
            if form_ws.cell(row, DataForm.MAKEUP_TEST_CHECK_COLUMN).value in ("x", "X"):
                continue

            info_exists, makeup_test_weekday, makeup_test_time, _ = omikron.studentinfo.get_student_info(student_ws, student_name)
            if info_exists and makeup_test_weekday:
                complete, calculated_schedule, time_index = calculate_makeup_test_schedule(makeup_test_weekday, makeup_test_date)
                if complete:
                    s = date_to_kor_date(calculated_schedule)
                    if makeup_test_time is not None:
                        mt = str(makeup_test_time)
                        if "/" in mt and len(makeup_test_weekday.split("/")) == len(mt.split("/")):
                            s = f"{s} {mt.split('/')[time_index]}시"
                        elif "/" not in mt:
                            s = f"{s} {mt}시"
                    sched_ops.append((class_index, student_name, test_name, s))
                    continue
            elif not info_exists:
                prog.warning(f"{student_name}의 학생 정보가 존재하지 않습니다.")

            nosched_ops.append((class_index, student_name, test_name))

        prog.step("시험 결과 요약 완료")

        # DAILY

        driver.switch_to.window(driver.window_handles[Chrome.DAILYTEST_RESULT_TAB])
        for class_index, student_name, test_name, test_score, test_average in daily_ops:
            if class_index not in daily_cache:
                daily_cache[class_index] = _cache_table_inputs(driver, class_index)

            inputs = daily_cache[class_index].get(student_name)
            if not inputs:
                prog.warning(f"아이소식에 {student_name} 학생이 존재하지 않습니다.")
                continue

            in0, in1, in2 = inputs
            _set_input(driver, in0, test_name)
            _set_input(driver, in1, test_score)
            _set_value_with_events(driver, in2, test_average)
        driver.execute_script("window.scrollTo(0, 0);")

        prog.step("시험 결과 메시지 작성 완료")

        # NO_SCHEDULE

        driver.switch_to.window(driver.window_handles[Chrome.MAKEUPTEST_NO_SCHEDULE_TAB])
        for class_index, student_name, test_name in nosched_ops:
            if class_index not in nosched_cache:
                nosched_cache[class_index] = _cache_table_inputs(driver, class_index)

            inputs = nosched_cache[class_index].get(student_name)
            if not inputs:
                continue

            in0, in1, in2 = inputs
            _set_value_with_events(driver, in0, test_name)
        driver.execute_script("window.scrollTo(0, 0);")

        prog.step("재시험 메시지 작성 완료")

        # SCHEDULE

        driver.switch_to.window(driver.window_handles[Chrome.MAKEUPTEST_SCHEDULE_TAB])
        for class_index, student_name, test_name, schedule_str in sched_ops:
            if class_index not in sched_cache:
                sched_cache[class_index] = _cache_table_inputs(driver, class_index)

            inputs = sched_cache[class_index].get(student_name)
            if not inputs:
                continue

            in0, in1, in2 = inputs
            _set_input(driver, in0, test_name)
            _set_value_with_events(driver, in1, schedule_str)
        driver.execute_script("window.scrollTo(0, 0);")

        prog.step("재시험 일정 메시지 작성 완료")

        driver.switch_to.window(driver.window_handles[Chrome.DAILYTEST_RESULT_TAB])

    finally:
        try:
            if form_wb is not None:
                form_wb.close()
        except Exception:
            pass
        try:
            if student_wb is not None:
                student_wb.close()
        except Exception:
            pass

        return True

def send_individual_test_message(student_name:str, class_name:int, test_name:int, test_score:int, test_average:int, makeup_test_check:bool, makeup_test_date:dict, prog:Progress) -> bool:
    """
    개별 시험에 대한 결과 메시지 전송
    """

    service = Service()
    service.creation_flags = CREATE_NO_WINDOW
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--blink-settings=imagesEnabled=false")
    options.page_load_strategy = "eager"

    if " (모의고사)" in class_name: class_name = class_name[:-7]

    student_wb = omikron.studentinfo.open()
    student_ws = omikron.studentinfo.open_worksheet(student_wb)

    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=service, options=options)
    
    # 아이소식 접속
    driver.get(URL)
    driver.execute_script("document.title = '시험 결과 전송'")
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
        prog.warning(f"아이소식의 {class_name} 내 {student_name} 학생이 존재하지 않습니다.")
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

    # 재시험 안내
    driver.execute_script(f"window.open('{URL}')")
    driver.switch_to.window(driver.window_handles[Chrome.INDIVIDUAL_MAKEUPTEST_TAB])
    driver.execute_script("document.title = '재시험 안내'")

    # 학생 정보 검색
    info_exists, makeup_test_weekday, makeup_test_time, _ = omikron.studentinfo.get_student_info(student_ws, student_name)
    if not info_exists:
        prog.warning(f"{student_name}의 학생 정보가 존재하지 않습니다.")

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
                        prog.warning(f"{student_name}의 재시험 시간이 올바른 양식이 아닙니다.")

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
            prog.warning(f"{student_name}의 재시험 요일이 올바른 양식이 아닙니다.")

    # 재시험 일정 없음
    driver.execute_script(f"arguments[0].value = '{MAKEUP_TEST_NO_SCHEDULE_MESSAGE}'", driver.find_element(By.XPATH, '//*[@id="ctitle"]'))
    driver.find_element(By.XPATH, '//*[@id="ctitle"]').send_keys(' \b')
    trs = driver.find_element(By.ID, f"table_{str(class_index)}").find_elements(By.CLASS_NAME, "style12")
    tds = trs[student_index].find_elements(By.TAG_NAME, "td")
    driver.execute_script(f"arguments[0].value = '{test_name}'", tds[0].find_element(By.TAG_NAME, "input"))

    tds[1].find_element(By.TAG_NAME, "input").send_keys(' \b')

    driver.switch_to.window(driver.window_handles[Chrome.DAILYTEST_RESULT_TAB])

    return True
