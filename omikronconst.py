# Omikron v1.2.0-beta6
class DataFile:
    TEST_TIME_COLUMN = 1
    DATE_COLUMN = 2
    CLASS_NAME_COLUMN = 3
    TEACHER_COLUMN = 4
    STUDENT_NAME_COLUMN = 5
    AVERAGE_SCORE_COLUMN = 6
    DATA_COLUMN = AVERAGE_SCORE_COLUMN + 1
    MAX = 6

class DataForm:
    DATE_COLUMN = 1
    TEST_TIME_COLUMN = 2
    CLASS_NAME_COLUMN = 3
    STUDENT_NAME_COLUMN = 4
    TEACHER_COLUMN = 5
    DAILYTEST_TEST_NAME_COLUMN = 6
    DAILYTEST_SCORE_COLUMN = 7
    DAILYTEST_AVERAGE_COLUMN = 8
    MOCKTEST_TEST_NAME_COLUMN = 9
    MOCKTEST_SCORE_COLUMN = 10
    MOCKTEST_AVERAGE_COLUMN = 11
    MAKEUP_TEST_CHECK_COLUMN = 12
    MAX = 12

class MakeupTestList:
    TEST_DATE_COLUMN = 1
    CLASS_NAME_COLUMN = 2
    TEACHER_COLUMN = 3
    STUDENT_NAME_COLUMN = 4
    TEST_NAME_COLUMN = 5
    TEST_SCORE_COLUMN = 6
    MAKEUP_TEST_WEEK_DATE_COLUMN = 7
    MAKEUP_TEST_TIME_COLUMN = 8
    MAKEUP_TEST_DATE_COLUMN = 9
    MAKEUP_TEST_SCORE_COLUMN = 10
    ETC_COLUMN = 11
    MAX = 11

class ClassInfo:
    CLASS_NAME_COLUMN = 1
    TEACHER_COLUMN = 2
    DATE_COLUMN = 3
    TEST_TIME_COLUMN = 4
    MAX = 4

class StudentInfo:
    STUDENT_NAME_COLUMN = 1
    CLASS_NAME_COLUMN = 2
    TEACHER_COLUMN = 3
    MAKEUP_TEST_WEEK_DATE_COLUMN = 4
    MAKEUP_TEST_TIME_COLUMN = 5
    NEW_STUDENT_CHECK_COLUMN = 6
    MAX = 6
