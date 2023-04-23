# Omikron v1.2.0-alpha
import calendar
import tkinter as tk

from omikronconst import *
from datetime import date as DATE, datetime, timedelta
from dateutil.relativedelta import relativedelta

def holiday_dialog() -> dict:
    def quitEvent():
        window.destroy()
    window=tk.Tk()
    width = 200
    height = 300
    x = int((window.winfo_screenwidth()/4) - (width/2))
    y = int((window.winfo_screenheight()/2) - (height/2))
    window.geometry(f'{width}x{height}+{x}+{y}')
    window.title('휴일 선택')
    window.resizable(False, False)
    window.protocol("WM_DELETE_WINDOW", quitEvent)

    today = DATE.today()
    weekday = ('월', '화', '수', '목', '금', '토', '일')
    makeup_test_date = {weekday[i] : today + relativedelta(weekday=i) for i in range(7)}
    for key, value in makeup_test_date.items():
        if value == today: makeup_test_date[key] += timedelta(days=7)

    mon = tk.IntVar()
    tue = tk.IntVar()
    wed = tk.IntVar()
    thu = tk.IntVar()
    fri = tk.IntVar()
    sat = tk.IntVar()
    sun = tk.IntVar()
    var_tuple = (mon, tue, wed, thu, fri, sat, sun)
    tk.Label(window, text='\n다음 중 휴일을 선택해주세요\n').pack()
    sort = today.weekday()+1
    for i in range(7):
        tk.Checkbutton(window, text=str(makeup_test_date[weekday[(sort+i)%7]]) + ' ' + weekday[(sort+i)%7], variable=var_tuple[(sort+i)%7]).pack()
    tk.Label(window, text='\n').pack()
    tk.Button(window, text="확인", width=10 , command=window.destroy).pack()
    
    window.mainloop()
    for i in range(7):
        if var_tuple[i].get() == 1:
            makeup_test_date[weekday[i]] += timedelta(days=7)

    return makeup_test_date

print(holiday_dialog())