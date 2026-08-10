"""Read-only HTTP access to Aisosik class and student data."""

from urllib.parse import urlparse

import requests
import urllib3
from bs4 import BeautifulSoup

import tdm.config
from tdm.domain.models import Chrome

def _fetch_aisosik_soup() -> BeautifulSoup:
    """
    아이소식 페이지 HTML을 가져와 BeautifulSoup로 반환.
    - 로그인/쿠키가 필요한 페이지면, 여기에서 세션/쿠키 처리하도록 확장하면 됨.
    """
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
    }

    target_host = urlparse(tdm.config.URL).hostname

    with requests.Session() as s:
        try:
            r = s.get(tdm.config.URL, headers=headers, timeout=10)
        except requests.exceptions.SSLError:
            # Some deployed iday-b2 endpoints currently serve expired certs.
            # Fallback keeps the app usable until server-side certs are fixed.
            if target_host != "dbserver2.iday-b2.com":
                raise
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            r = s.get(tdm.config.URL, headers=headers, timeout=10, verify=False)
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
