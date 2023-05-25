import json
import os.path
import openpyxl as xl

config = json.load(open("./config.json", encoding="UTF8"))

# 데이터파일 시트명 변경
data_file_wb = xl.load_workbook(f"./data/{config['dataFileName']}.xlsx")
data_file_ws = data_file_wb["DailyTest"]
data_file_ws.title = "데일리테스트"
data_file_wb.save(f"./data/{config['dataFileName']}.xlsx")

# 재시험 정보 파일 -> 학생 정보 파일
student_wb = xl.load_workbook("./재시험 정보.xlsx")
student_ws = student_wb["재시험 정보"]
student_ws.title = "학생 정보"
student_wb.save("./학생 정보.xlsx")
os.remove("./재시험 정보.xlsx")
