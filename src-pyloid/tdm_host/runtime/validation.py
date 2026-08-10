"""Validation helpers that do not depend on the RPC transport."""

from typing import Optional
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from bs4 import BeautifulSoup


def validate_script_page_url(url: str) -> Optional[str]:
    """Return a user-facing warning when the URL is not a Script page."""
    try:
        request = Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
            },
        )
        with urlopen(request, timeout=10) as response:
            raw = response.read()

        html = raw.decode("utf-8", errors="replace")
        soup = BeautifulSoup(html, "html.parser")
        marker_ok = any(
            center.get_text(strip=True) == "Script" for center in soup.select("center")
        )
        if not marker_ok:
            return "URL 페이지가 스크립트 페이지가 아닙니다. /html/body/div[1]/h1/center 값이 'Script'여야 합니다."
        return None
    except (HTTPError, URLError, TimeoutError):
        return "URL에 접속할 수 없습니다. 주소와 네트워크 상태를 확인해 주세요."
    except Exception:
        return "URL 검증 중 오류가 발생했습니다."
