"""Selenium-based Aisosik message composition."""

from typing import Any, TypeAlias

from selenium.common.exceptions import SessionNotCreatedException, WebDriverException
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver as ChromeWebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from win32process import CREATE_NO_WINDOW
from bs4 import BeautifulSoup

import tdm.config
import tdm.excel.data_form
import tdm.excel.student_info
from tdm.domain.errors import ChromeDriverVersionMismatchException
from tdm.domain.models import Chrome, DataForm
from tdm.domain.progress import Progress
from tdm.excel.utils import calculate_makeup_test_schedule, date_to_kor_date

InputTriple: TypeAlias = tuple[WebElement, WebElement, WebElement]

def _set_input(driver: ChromeWebDriver, input_el: WebElement, value: Any) -> None:
    driver.execute_script("arguments[0].value = arguments[1]", input_el, str(value))

def _set_value_with_events(driver: ChromeWebDriver, el: WebElement, value: Any) -> None:
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
    """, el, str(value))

def _cache_table_inputs(driver: ChromeWebDriver, class_index: int) -> dict[str, InputTriple]:
    """
    table_{class_index}에서
    학생이름 -> (시험명 input, 점수 input, 평균 input) 캐싱
    """
    table = driver.find_element(By.ID, f"table_{class_index}")
    rows = table.find_elements(By.CLASS_NAME, "style12")

    name_to_inputs: dict[str, InputTriple] = {}
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

def _create_chrome_driver(service: Service, options: ChromeOptions) -> ChromeWebDriver:
    try:
        return ChromeWebDriver(service=service, options=options)
    except (SessionNotCreatedException, WebDriverException) as e:
        msg = str(e).lower()
        version_mismatch_patterns = (
            "this version of chromedriver only supports chrome version",
            "current browser version is",
            "only supports chrome version",
        )
        if any(p in msg for p in version_mismatch_patterns):
            raise ChromeDriverVersionMismatchException(
                "셀레니움 기능을 실행할 수 없습니다. 설치된 Chrome 버전과 ChromeDriver(셀레니움)가 호환되지 않습니다. 프로그램을 최신 버전으로 업데이트하거나 Chrome 버전을 확인해 주세요."
            ) from e
        raise

def send_test_result_message(filepath: str, makeup_test_date: dict[str, Any], prog: Progress) -> bool:
    """
    기록 양식의 데이터를 추출하여 아이소식 스크립트 작성
    """
    form_wb = None
    student_wb = None
    try:
        service = Service()
        service.creation_flags = CREATE_NO_WINDOW
        options = ChromeOptions()
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--blink-settings=imagesEnabled=false")
        options.add_argument("--ignore-certificate-errors")
        options.add_argument("--allow-running-insecure-content")
        options.accept_insecure_certs = True
        options.page_load_strategy = "eager"
        options.add_experimental_option("detach", True)

        form_wb = tdm.excel.data_form.open(filepath)
        form_ws = tdm.excel.data_form.open_worksheet(form_wb)

        student_wb = tdm.excel.student_info.open()
        student_ws = tdm.excel.student_info.open_worksheet(student_wb)

        driver = _create_chrome_driver(service=service, options=options)
        
        # 아이소식 접속
        driver.get(tdm.config.URL)
        _set_value_with_events(driver, driver.find_element(By.XPATH, '//*[@id="ctitle"]'), tdm.config.TEST_RESULT_MESSAGE)
        driver.execute_script("document.title = '시험 결과 전송'")

        driver.execute_script("window.open(arguments[0])", tdm.config.URL)
        driver.switch_to.window(driver.window_handles[Chrome.MAKEUPTEST_NO_SCHEDULE_TAB])
        _set_value_with_events(
            driver,
            driver.find_element(By.XPATH, '//*[@id="ctitle"]'),
            tdm.config.MAKEUP_TEST_NO_SCHEDULE_MESSAGE,
        )
        driver.execute_script("document.title = '재시험 일정 없는 학생'")

        driver.execute_script("window.open(arguments[0])", tdm.config.URL)
        driver.switch_to.window(driver.window_handles[Chrome.MAKEUPTEST_SCHEDULE_TAB])
        _set_value_with_events(
            driver,
            driver.find_element(By.XPATH, '//*[@id="ctitle"]'),
            tdm.config.MAKEUP_TEST_SCHEDULE_MESSAGE,
        )
        driver.execute_script("document.title = '재시험 일정 있는 학생'")

        driver.switch_to.window(driver.window_handles[Chrome.DAILYTEST_RESULT_TAB])

        # 반 인덱스 dict
        soup = BeautifulSoup(driver.page_source, "html.parser")
        names = [el.get_text(strip=True) for el in soup.select(".style1")]
        table_index_dict = {name: i for i, name in enumerate(names) if name}
        # table_names = driver.find_elements(By.CLASS_NAME, "style1")
        # table_index_dict = {table_name.text.strip() : i for i, table_name in enumerate(table_names)}

        # 탭별 캐시: class_index -> (student_name -> inputs)
        daily_cache: dict[int, dict[str, InputTriple]] = {}
        nosched_cache: dict[int, dict[str, InputTriple]] = {}
        sched_cache: dict[int, dict[str, InputTriple]] = {}

        # 탭별 작업 큐
        daily_ops: list[tuple[int, str, str | None, int | float, str | None]] = []
        nosched_ops: list[tuple[int, str, str | None]] = []
        sched_ops: list[tuple[int, str, str | None, str]] = []

        # 루프에서 매 행마다 DOM 조작하지 말고 "작업만 수집"
        class_index = None
        class_name = None
        daily_test_name = mock_test_name = None
        daily_test_average = mock_test_average = None

        for row in range(2, form_ws.max_row + 1):
            if form_ws.cell(row, DataForm.CLASS_NAME_COLUMN).value is not None:
                class_name = str(form_ws.cell(row, DataForm.CLASS_NAME_COLUMN).value)
                daily_name_value = form_ws.cell(row, DataForm.DAILYTEST_NAME_COLUMN).value
                mock_name_value = form_ws.cell(row, DataForm.MOCKTEST_NAME_COLUMN).value
                daily_avg_value = form_ws.cell(row, DataForm.DAILYTEST_AVERAGE_COLUMN).value
                mock_avg_value = form_ws.cell(row, DataForm.MOCKTEST_AVERAGE_COLUMN).value

                daily_test_name = str(daily_name_value) if daily_name_value is not None else None
                mock_test_name = str(mock_name_value) if mock_name_value is not None else None
                daily_test_average = str(daily_avg_value) if daily_avg_value is not None else None
                mock_test_average = str(mock_avg_value) if mock_avg_value is not None else None

                if daily_test_name is None and mock_test_name is None:
                    continue

                class_index = table_index_dict.get(class_name)
                if class_index is None:
                    prog.warning(f"아이소식에 {class_name} 반이 존재하지 않습니다.")
                    continue

            student_name_raw = form_ws.cell(row, DataForm.STUDENT_NAME_COLUMN).value
            if student_name_raw is None:
                continue
            student_name = str(student_name_raw).strip()
            if not student_name:
                continue
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

            if class_index is None:
                # 반 매핑 실패 상태에서는 쓰기 작업을 생성하지 않는다.
                continue

            daily_ops.append((class_index, student_name, test_name, test_score, test_average))

            # 재시험 분기(여기서는 DOM 안 건드리고 “어느 탭에 쓸지”만 결정)
            if test_score >= 80:
                continue
            if form_ws.cell(row, DataForm.MAKEUP_TEST_CHECK_COLUMN).value in ("x", "X"):
                continue

            info_exists, makeup_test_weekday, makeup_test_time, _ = tdm.excel.student_info.get_student_info(student_ws, student_name)
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
        return True
    except ChromeDriverVersionMismatchException:
        raise
    except Exception as e:
        raise Exception(f"메시지 작성 중 오류가 발생했습니다: {e}")

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

def send_individual_test_message(
    student_name: str,
    class_name: str,
    test_name: str,
    test_score: int | float,
    test_average: int | float | str,
    makeup_test_check: bool,
    makeup_test_date: dict[str, Any],
    prog: Progress,
) -> bool:
    """
    개별 시험에 대한 결과 메시지 전송
    """

    service = Service()
    service.creation_flags = CREATE_NO_WINDOW
    options = ChromeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--blink-settings=imagesEnabled=false")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--allow-running-insecure-content")
    options.accept_insecure_certs = True
    options.page_load_strategy = "eager"
    options.add_experimental_option("detach", True)

    if " (모의고사)" in class_name:
        class_name = class_name[:-7]

    student_wb = None
    student_ws = None
    try:
        driver = _create_chrome_driver(service=service, options=options)
        # 아이소식 접속
        driver.get(tdm.config.URL)
        driver.execute_script("document.title = '시험 결과 전송'")
        _set_value_with_events(driver, driver.find_element(By.XPATH, '//*[@id="ctitle"]'), tdm.config.TEST_RESULT_MESSAGE)

        # 반 인덱스 dict (BeautifulSoup 사용으로 DOM 접근 최소화)
        soup = BeautifulSoup(driver.page_source, "html.parser")
        table_names = [el.get_text(strip=True) for el in soup.select(".style1")]
        table_index_dict = {name: i for i, name in enumerate(table_names) if name}

        class_index = table_index_dict.get(class_name)
        if class_index is None:
            prog.warning(f"아이소식에 {class_name} 반이 존재하지 않습니다.")
            return False

        # DAILY 탭에서 학생 입력칸 캐시
        daily_inputs = _cache_table_inputs(driver, class_index)
        target_inputs = daily_inputs.get(student_name)
        if not target_inputs:
            prog.warning(f"아이소식의 {class_name} 내 {student_name} 학생이 존재하지 않습니다.")
            return False

        in0, in1, in2 = target_inputs
        _set_input(driver, in0, test_name)
        _set_input(driver, in1, test_score)
        _set_value_with_events(driver, in2, test_average)

        if test_score >= 80 or makeup_test_check:
            return True

        # 재시험 안내가 필요한 경우에만 학생정보 파일 오픈
        student_wb = tdm.excel.student_info.open()
        student_ws = tdm.excel.student_info.open_worksheet(student_wb)

        # 재시험 탭 오픈
        driver.execute_script("window.open(arguments[0])", tdm.config.URL)
        driver.switch_to.window(driver.window_handles[Chrome.INDIVIDUAL_MAKEUPTEST_TAB])
        driver.execute_script("document.title = '재시험 안내'")

        makeup_inputs = _cache_table_inputs(driver, class_index)
        makeup_target_inputs = makeup_inputs.get(student_name)
        if not makeup_target_inputs:
            prog.warning(f"아이소식의 {class_name} 내 {student_name} 학생이 존재하지 않습니다.")
            driver.switch_to.window(driver.window_handles[Chrome.DAILYTEST_RESULT_TAB])
            return False

        m0, m1, m2 = makeup_target_inputs

        # 학생 정보 검색
        info_exists, makeup_test_weekday, makeup_test_time, _ = tdm.excel.student_info.get_student_info(student_ws, student_name)
        if not info_exists:
            prog.warning(f"{student_name}의 학생 정보가 존재하지 않습니다.")

        if info_exists and makeup_test_weekday is not None:
            complete, calculated_schedule, time_index = calculate_makeup_test_schedule(makeup_test_weekday, makeup_test_date)
            if complete:
                _set_value_with_events(
                    driver,
                    driver.find_element(By.XPATH, '//*[@id="ctitle"]'),
                    tdm.config.MAKEUP_TEST_SCHEDULE_MESSAGE,
                )
                _set_input(driver, m0, test_name)

                calculated_schedule_str = date_to_kor_date(calculated_schedule)
                schedule_text = calculated_schedule_str

                if makeup_test_time is not None:
                    mt = str(makeup_test_time)
                    if "/" in mt:
                        if len(makeup_test_weekday.split("/")) == len(mt.split("/")):
                            schedule_text = f"{calculated_schedule_str} {mt.split('/')[time_index]}시"
                        else:
                            prog.warning(f"{student_name}의 재시험 시간이 올바른 양식이 아닙니다.")
                    else:
                        schedule_text = f"{calculated_schedule_str} {mt}시"

                _set_value_with_events(driver, m1, schedule_text)
                _set_value_with_events(driver, m2, "")
                driver.switch_to.window(driver.window_handles[Chrome.DAILYTEST_RESULT_TAB])
                return True
            else:
                prog.warning(f"{student_name}의 재시험 요일이 올바른 양식이 아닙니다.")

        # 재시험 일정 없음
        _set_value_with_events(
            driver,
            driver.find_element(By.XPATH, '//*[@id="ctitle"]'),
            tdm.config.MAKEUP_TEST_NO_SCHEDULE_MESSAGE,
        )
        _set_input(driver, m0, test_name)
        _set_value_with_events(driver, m1, "")

        driver.switch_to.window(driver.window_handles[Chrome.DAILYTEST_RESULT_TAB])
        return True
    finally:
        if student_wb is not None:
            try:
                student_wb.close()
            except Exception:
                pass
