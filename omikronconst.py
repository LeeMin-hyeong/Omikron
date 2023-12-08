VERSION = "Omikron v1.4.2"

class DataFile:
    TEST_TIME_COLUMN          =  1
    CLASS_WEEKDAY_COLUMN      =  2
    CLASS_NAME_COLUMN         =  3
    TEACHER_NAME_COLUMN       =  4
    STUDENT_NAME_COLUMN       =  5
    AVERAGE_SCORE_COLUMN      =  6
    MAX                       =  AVERAGE_SCORE_COLUMN
    DATA_COLUMN               =  MAX + 1

class DataForm:
    CLASS_WEEKDAY_COLUMN      =  1
    TEST_TIME_COLUMN          =  2
    CLASS_NAME_COLUMN         =  3
    STUDENT_NAME_COLUMN       =  4
    TEACHER_NAME_COLUMN       =  5
    DAILYTEST_NAME_COLUMN     =  6
    DAILYTEST_SCORE_COLUMN    =  7
    DAILYTEST_AVERAGE_COLUMN  =  8
    MOCKTEST_NAME_COLUMN      =  9
    MOCKTEST_SCORE_COLUMN     = 10
    MOCKTEST_AVERAGE_COLUMN   = 11
    MAKEUP_TEST_CHECK_COLUMN  = 12
    MAX                       = MAKEUP_TEST_CHECK_COLUMN

class MakeupTestList:
    TEST_DATE_COLUMN          =  1
    CLASS_NAME_COLUMN         =  2
    TEACHER_NAME_COLUMN       =  3
    STUDENT_NAME_COLUMN       =  4
    TEST_NAME_COLUMN          =  5
    TEST_SCORE_COLUMN         =  6
    MAKEUPTEST_WEEKDAY_COLUMN =  7
    MAKEUPTEST_TIME_COLUMN    =  8
    MAKEUPTEST_DATE_COLUMN    =  9
    MAKEUPTEST_SCORE_COLUMN   = 10
    ETC_COLUMN                = 11
    MAX                       = ETC_COLUMN

class ClassInfo:
    CLASS_NAME_COLUMN         =  1
    TEACHER_NAME_COLUMN       =  2
    CLASS_WEEKDAY_COLUMN      =  3
    TEST_TIME_COLUMN          =  4
    MAX                       =  TEST_TIME_COLUMN

class StudentInfo:
    STUDENT_NAME_COLUMN       =  1
    CLASS_NAME_COLUMN         =  2
    TEACHER_NAME_COLUMN       =  3
    MAKEUPTEST_WEEKDAY_COLUMN =  4
    MAKEUPTEST_TIME_COLUMN    =  5
    NEW_STUDENT_CHECK_COLUMN  =  6
    MAX                       =  NEW_STUDENT_CHECK_COLUMN
