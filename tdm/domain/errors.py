
class NoMatchingSheetException(Exception):
    """
    이름이 일치하는 데이터 시트를 찾을 수 없음
    """
    pass

class FileOpenException(Exception):
    """
    파일이 열려 있어 작업을 수행할 수 없음
    """
    pass

class ReopenFileException(Exception):
    """
    zipfile.BadZipFile
    """
    pass


class ExcelRequiredException(Exception):
    """
    Microsoft Excel 이 설치되어 있지 않거나 실행할 수 없음
    """
    pass


class ChromeDriverVersionMismatchException(Exception):
    """
    설치된 Chrome 버전과 Selenium/ChromeDriver가 호환되지 않음
    """
    pass
