
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