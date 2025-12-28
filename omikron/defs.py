class Chrome:
    ACTUAL_CLASS_START_INDEX   =  3 # 아이소식 내 실제 반이 시작되는 테이블 인덱스

    DAILYTEST_RESULT_TAB       =  0 # 시험 결과 탭
    MAKEUPTEST_NO_SCHEDULE_TAB =  1 # 재시험 고지 탭(날짜 미지정)
    MAKEUPTEST_SCHEDULE_TAB    =  2 # 재시험 고지 탭(날짜 지정)
    INDIVIDUAL_MAKEUPTEST_TAB  =  1 # 개별 시험 결과 메시지 탭

class DataFile:
    PRE_DATA_FILE_NAME         = "지난 데이터"
    TEMP_FILE_NAME             = "9IwTEoG59MS6h2UoqveD"
    DEFAULT_SHEET_NAME         = "테스트 데이터"
    CLASS_NAME_COLUMN          =  1
    TEACHER_NAME_COLUMN        =  2
    STUDENT_NAME_COLUMN        =  3
    AVERAGE_SCORE_COLUMN       =  4
    MAX                        = AVERAGE_SCORE_COLUMN
    DATA_COLUMN                = MAX + 1

class DataForm: 
    DEFAULT_NAME               = "데일리테스트 기록 양식"
    CLASS_WEEKDAY_COLUMN       =  1
    TEST_TIME_COLUMN           =  2
    CLASS_NAME_COLUMN          =  3
    STUDENT_NAME_COLUMN        =  4
    TEACHER_NAME_COLUMN        =  5
    DAILYTEST_NAME_COLUMN      =  6
    DAILYTEST_SCORE_COLUMN     =  7
    DAILYTEST_AVERAGE_COLUMN   =  8
    MOCKTEST_NAME_COLUMN       =  9
    MOCKTEST_SCORE_COLUMN      = 10
    MOCKTEST_AVERAGE_COLUMN    = 11
    MAKEUP_TEST_CHECK_COLUMN   = 12
    MAX                        = MAKEUP_TEST_CHECK_COLUMN

class MakeupTestList: 
    DEFAULT_NAME               = "재시험 명단"
    TEST_DATE_COLUMN           = 1
    CLASS_NAME_COLUMN          = 2
    TEACHER_NAME_COLUMN        = 3
    STUDENT_NAME_COLUMN        = 4
    TEST_NAME_COLUMN           = 5
    MAKEUPTEST_DATE_COLUMN     = 6
    MAKEUPTEST_SCORE_COLUMN    = 7
    ETC_COLUMN                 = 8
    MAX                        = ETC_COLUMN

class ClassInfo: 
    DEFAULT_NAME               = "반 정보"
    TEMP_FILE_NAME             = "반 정보(임시)"
    CLASS_NAME_COLUMN          = 1
    TEACHER_NAME_COLUMN        = 2
    CLASS_WEEKDAY_COLUMN       = 3
    TEST_TIME_COLUMN           = 4
    MOCKTEST_CHECK_COLUMN      = 5
    MAX                        = MOCKTEST_CHECK_COLUMN

class StudentInfo: 
    DEFAULT_NAME               = "학생 정보"
    STUDENT_NAME_COLUMN        = 1
    MAKEUPTEST_WEEKDAY_COLUMN  = 2
    MAKEUPTEST_TIME_COLUMN     = 3
    NEW_STUDENT_CHECK_COLUMN   = 4
    MAX                        = NEW_STUDENT_CHECK_COLUMN
