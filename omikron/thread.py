import omikron.chrome
import omikron.classinfo
import omikron.datafile
import omikron.dataform
import omikron.makeuptest
import omikron.studentinfo
import omikron.thread

from omikron.log import OmikronLog

global thread_end_flag
thread_end_flag = False
"""thread 종료 플래그"""

def make_class_info_file_thread():
    OmikronLog.log("반 정보 파일 생성 중...")
    if omikron.classinfo.make_file():
        OmikronLog.log("반 정보 파일을 생성했습니다.")

    global thread_end_flag
    thread_end_flag = True

def make_student_info_file_thread():
    OmikronLog.log("학생 정보 파일 생성 중...")
    if omikron.studentinfo.make_file():
        OmikronLog.log("학생 정보 파일을 생성했습니다.")

    global thread_end_flag
    thread_end_flag = True

def make_data_file_thread():
    OmikronLog.log("데이터 파일 생성 중...")
    if omikron.datafile.make_file():
        OmikronLog.log("데이터 파일을 생성했습니다.")

    global thread_end_flag
    thread_end_flag = True

def update_class_thread():
    OmikronLog.log("반 업데이트 진행중")

    complete, data_file_wb = omikron.datafile.update_class()
    if not complete: return

    class_wb = omikron.classinfo.open_temp()

    try:
        omikron.datafile.save(data_file_wb)
    except:
        OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
        return
    try:
        omikron.classinfo.save(class_wb)
    except:
        OmikronLog.error(r"반 정보 파일을 닫은 뒤 다시 시도해 주세요.")
        return

    omikron.classinfo.delete_temp()

    OmikronLog.log("반 업데이트를 완료하였습니다.")

    global thread_end_flag
    thread_end_flag = True

def make_data_form_thread():
    OmikronLog.log("데일리테스트 기록 양식 생성 중...")
    if omikron.dataform.make_file():
        OmikronLog.log("데일리테스트 기록 양식을 생성했습니다.")

    global thread_end_flag
    thread_end_flag = True

def save_test_data_thread(filepath:str, makeup_test_date:dict):
    OmikronLog.log("데이터 저장 및 재시험 명단 작성중...")

    complete, data_file_wb = omikron.datafile.save_test_data(filepath)
    if not complete: return

    complete, makeup_test_wb = omikron.makeuptest.save_makeup_test_list(filepath, makeup_test_date)
    if not complete: return

    try:
        omikron.datafile.save(data_file_wb)
        omikron.datafile.delete_temp()
    except:
        OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
        return
    try:
        omikron.makeuptest.save(makeup_test_wb)
    except:
        OmikronLog.error(r"재시험 명단 파일을 닫은 뒤 다시 시도해 주세요.")
        return

    OmikronLog.log("데이터 저장 및 재시험 명단 작성을 완료했습니다.")

    global thread_end_flag
    thread_end_flag = True

def send_message_thread(filepath:str, makeup_test_date:dict):
    OmikronLog.log("테스트 결과 메시지 작성 중...")
    if omikron.chrome.send_test_result_message(filepath, makeup_test_date):
        OmikronLog.log("테스트 결과 메시지 작성을 완료했습니다.")
        OmikronLog.log("메시지 확인 후 전송해주세요.")
    else: return

    global thread_end_flag
    thread_end_flag = True

def save_individual_test_thread(student_name, class_name, test_name, target_row, target_col, test_score, makeup_test_check, makeup_test_date):
    # TODO : 오류 발생 시 데이터 보존
    OmikronLog.log(f"{student_name} 개별 시험 결과 저장 중...")
    complete, test_average, data_file_wb = omikron.datafile.save_individual_test_data(target_row, target_col, test_score)
    if complete:
        OmikronLog.log("데이터 저장을 완료했습니다.")
    else: return

    if test_score < 80 and not makeup_test_check:
        complete, makeup_test_wb = omikron.makeuptest.save_individual_makeup_test(student_name, class_name, test_name, test_score, makeup_test_date)
        if not complete: return

    OmikronLog.log("데이터 저장을 완료했습니다.")

    try:
        omikron.datafile.save(data_file_wb)
        omikron.datafile.delete_temp()
    except:
        OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
        return
    try:
        omikron.makeuptest.save(makeup_test_wb)
    except:
        OmikronLog.error(r"재시험 명단 파일을 닫은 뒤 다시 시도해 주세요.")
        return

    if omikron.chrome.send_individual_test_message(student_name, class_name, test_name, test_score, test_average, makeup_test_check, makeup_test_date):
        OmikronLog.log("테스트 결과 메시지 작성을 완료했습니다.")
        OmikronLog.log("메시지 확인 후 전송해주세요.")
    else: return

    global thread_end_flag
    thread_end_flag = True

def save_makeup_test_result_thread(target_row:int, makeup_test_score:str):
    if omikron.makeuptest.save_makeup_test_result(target_row, makeup_test_score):
        OmikronLog.log(f"{target_row} 행에 재시험 점수를 기록하였습니다.")

    global thread_end_flag
    thread_end_flag = True

def conditional_formatting_thread():
    OmikronLog.log("조건부 서식 재지정 중...")
    if omikron.datafile.conditional_formatting():
        OmikronLog.log("조건부 서식 재지정을 완료했습니다.")

    global thread_end_flag
    thread_end_flag = True

def add_student_thread(target_student_name:str, target_class_name:str):
    OmikronLog.log("학생 추가 중...")

    if not omikron.chrome.check_student_exists(target_student_name, target_class_name):
        OmikronLog.error(f"아이소식에 {target_student_name} 학생이 {target_class_name} 반에 업데이트 되지 않아 중단되었습니다.")
        return

    complete, data_file_wb = omikron.datafile.add_student(target_student_name, target_class_name)
    if not complete: return

    complete, student_wb = omikron.studentinfo.add_student(target_student_name)
    if not complete: return

    try:
        omikron.datafile.save(data_file_wb)
    except:
        OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
        return
    try:
        omikron.studentinfo.save(student_wb)
    except:
        OmikronLog.error(r"학생 정보 파일을 닫은 뒤 다시 시도해 주세요.")
        return

    OmikronLog.log(f"{target_student_name} 학생을 {target_class_name} 반에 추가하였습니다.")

    global thread_end_flag
    thread_end_flag = True

def delete_student_thread(target_student_name:str):
    OmikronLog.log("학생 퇴원 처리 중...")

    complete, data_file_wb = omikron.datafile.delete_student(target_student_name)
    if not complete: return

    complete, student_wb = omikron.studentinfo.delete_student(target_student_name)
    if not complete: return

    try:
        omikron.datafile.save(data_file_wb)
    except:
        OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
        return
    try:
        omikron.studentinfo.save(student_wb)
    except:
        OmikronLog.error(r"학생 정보 파일을 닫은 뒤 다시 시도해 주세요.")
        return

    OmikronLog.log(f"{target_student_name} 학생을 퇴원처리 하였습니다.")

    global thread_end_flag
    thread_end_flag = True

def move_student_thread(target_student_name:str, target_class_name:str, current_class_name:str):
    OmikronLog.log("학생 반 이동 중...")

    if not omikron.chrome.check_student_exists(target_student_name, target_class_name):
        OmikronLog.error(f"아이소식에 {target_student_name} 학생이 {target_class_name} 반에 업데이트 되지 않아 중단되었습니다.")
        return

    complete, data_file_wb = omikron.datafile.move_student(target_student_name, target_class_name, current_class_name)
    if not complete: return

    try:
        omikron.datafile.save(data_file_wb)
    except:
        OmikronLog.error(r"데이터 파일을 닫은 뒤 다시 시도해 주세요.")
        return

    OmikronLog.log(f"{target_student_name} 학생을 {current_class_name}에서 {target_class_name}으로 이동하였습니다.")

    global thread_end_flag
    thread_end_flag = True
