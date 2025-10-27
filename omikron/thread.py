import omikron.chrome
import omikron.classinfo
import omikron.datafile
import omikron.dataform
import omikron.makeuptest
import omikron.studentinfo

from omikron.log import OmikronLog

thread_end_flag = False
"""thread 종료 플래그"""

def make_class_info_file_thread():
    try:
        OmikronLog.log("반 정보 파일 생성 중...")

        omikron.classinfo.make_file()

        OmikronLog.log("반 정보 파일을 생성했습니다.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True

def make_student_info_file_thread():
    try:
        OmikronLog.log("학생 정보 파일 생성 중...")

        omikron.studentinfo.make_file()

        OmikronLog.log("학생 정보 파일을 생성했습니다.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True

def update_student_info_file_thread():
    try:
        OmikronLog.log("학생 정보 파일 업데이트 중...")

        omikron.studentinfo.update_student()

        OmikronLog.log("학생 정보 파일을 업데이트했습니다.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True

def make_data_file_thread():
    try:
        OmikronLog.log("데이터 파일 생성 중...")

        omikron.datafile.make_file()

        OmikronLog.log("데이터 파일을 생성했습니다.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True

def update_class_thread():
    try:
        OmikronLog.log("반 업데이트 진행 중...")

        omikron.datafile.update_class()

        omikron.classinfo.update_class()

        omikron.classinfo.delete_temp()

        OmikronLog.log("반 업데이트를 완료하였습니다.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True

def change_class_info_thread(target_class_name:str, target_teacher_name:str):
    try:
        OmikronLog.log("선생님 변경 진행 중...")

        omikron.datafile.change_class_info(target_class_name, target_teacher_name)

        omikron.classinfo.change_class_info(target_class_name, target_teacher_name)

        OmikronLog.log("선생님 변경을 완료하였습니다.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True

def make_data_form_thread():
    try:
        OmikronLog.log("데일리테스트 기록 양식 생성 중...")

        omikron.dataform.make_file()

        OmikronLog.log("데일리테스트 기록 양식을 생성했습니다.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True

def save_test_result_thread(filepath:str, makeup_test_date:dict):
    try:
        OmikronLog.log("데이터 저장 및 재시험 명단 작성중...")

        data_file_wb = omikron.datafile.save_test_data(filepath)

        makeup_test_wb = omikron.makeuptest.save_makeup_test_list(filepath, makeup_test_date)

        omikron.datafile.save(data_file_wb)
        omikron.datafile.delete_temp()

        omikron.makeuptest.save(makeup_test_wb)

        OmikronLog.log("데이터 저장 및 재시험 명단 작성을 완료했습니다.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True

def send_message_thread(filepath:str, makeup_test_date:dict):
    try:
        OmikronLog.log("테스트 결과 메시지 작성 중...")

        omikron.chrome.send_test_result_message(filepath, makeup_test_date)

        OmikronLog.log("테스트 결과 메시지 작성을 완료했습니다.")
        OmikronLog.log("메시지 확인 후 전송해주세요.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True

def save_individual_test_thread(student_name:str, class_name:str, test_name:str, target_row:int, target_col:int, test_score:int|float, makeup_test_check:bool, makeup_test_date:dict):
    try:
        OmikronLog.log(f"{student_name} 개별 시험 결과 저장 중...")
        test_average = omikron.datafile.save_individual_test_data(target_row, target_col, test_score)

        if test_score < 80 and not makeup_test_check:
            omikron.makeuptest.save_individual_makeup_test(student_name, class_name, test_name, test_score, makeup_test_date)

        omikron.chrome.send_individual_test_message(student_name, class_name, test_name, test_score, test_average, makeup_test_check, makeup_test_date)

        OmikronLog.log("테스트 결과 메시지 작성을 완료했습니다.")
        OmikronLog.log("메시지 확인 후 전송해주세요.")

        OmikronLog.log("데이터 저장을 완료했습니다.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True

def save_makeup_test_result_thread(target_row:int, makeup_test_score:str):
    try:
        omikron.makeuptest.save_makeup_test_result(target_row, makeup_test_score)

        OmikronLog.log(f"{target_row} 행에 재시험 점수를 기록하였습니다.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True

def conditional_formatting_thread():
    try:
        OmikronLog.log("조건부 서식 재지정 중...")

        omikron.datafile.conditional_formatting()

        OmikronLog.log("조건부 서식 재지정을 완료했습니다.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True

def add_student_thread(target_student_name:str, target_class_name:str):
    try:
        OmikronLog.log(f"{target_student_name} 학생 추가 중...")

        if not omikron.chrome.check_student_exists(target_student_name, target_class_name):
            OmikronLog.error(f"아이소식에 {target_student_name} 학생이 {target_class_name} 반에 업데이트 되지 않아 중단되었습니다.")
            return

        omikron.datafile.add_student(target_student_name, target_class_name)

        omikron.studentinfo.add_student(target_student_name)

        OmikronLog.log(f"{target_student_name} 학생을 {target_class_name} 반에 추가하였습니다.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True

def delete_student_thread(target_student_name:str):
    try:
        OmikronLog.log(f"{target_student_name} 학생 퇴원 처리 중...")

        omikron.datafile.delete_student(target_student_name)

        omikron.studentinfo.delete_student(target_student_name)

        OmikronLog.log(f"{target_student_name} 학생을 퇴원처리 하였습니다.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True

def move_student_thread(target_student_name:str, target_class_name:str, current_class_name:str):
    try:
        OmikronLog.log(f"{target_student_name} 학생 반 이동 중...")

        if not omikron.chrome.check_student_exists(target_student_name, target_class_name):
            OmikronLog.error(f"아이소식에 {target_student_name} 학생이 {target_class_name} 반에 업데이트 되지 않아 중단되었습니다.")
            return

        omikron.datafile.move_student(target_student_name, target_class_name, current_class_name)

        OmikronLog.log(f"{target_student_name} 학생을 {current_class_name}에서 {target_class_name}으로 이동하였습니다.")
    except Exception as e:
        OmikronLog.error(repr(e))
    finally:
        global thread_end_flag
        thread_end_flag = True
