import tkinter as tk
import calendar

from datetime import date, timedelta
from dateutil.relativedelta import relativedelta

def selector(weekday):
    makeupDate[weekday] += timedelta(days=7)

window=tk.Tk()
window.title('메시지 전송 : 휴일 선택')
window.geometry('200x300')
window.resizable(False, False)

today = date.today()

makeupDate={}
if today == today + relativedelta(weekday=calendar.MONDAY):
    makeupDate['월'] = today + timedelta(days=7)
else:
    makeupDate['월'] = today + relativedelta(weekday=calendar.MONDAY)

if today == today + relativedelta(weekday=calendar.TUESDAY):
    makeupDate['화'] = today + timedelta(days=7)
else:
    makeupDate['화'] = today + relativedelta(weekday=calendar.TUESDAY)

if today == today + relativedelta(weekday=calendar.WEDNESDAY):
    makeupDate['수'] = today + timedelta(days=7)
else:
    makeupDate['수'] = today + relativedelta(weekday=calendar.WEDNESDAY)

if today == today + relativedelta(weekday=calendar.THURSDAY):
    makeupDate['목'] = today + timedelta(days=7)
else:
    makeupDate['목'] = today + relativedelta(weekday=calendar.THURSDAY)

if today == today + relativedelta(weekday=calendar.FRIDAY):
    makeupDate['금'] = today + timedelta(days=7)
else:
    makeupDate['금'] = today + relativedelta(weekday=calendar.FRIDAY)

if today == today + relativedelta(weekday=calendar.SATURDAY):
    makeupDate['토'] = today + timedelta(days=7)
else:
    makeupDate['토'] = today + relativedelta(weekday=calendar.SATURDAY)

if today == today + relativedelta(weekday=calendar.SUNDAY):
    makeupDate['일'] = today + timedelta(days=7)
else:
    makeupDate['일'] = today + relativedelta(weekday=calendar.SUNDAY)


temp = sorted(makeupDate.items(), key=lambda x:x[1])

ch1 = tk.BooleanVar()
tk.Label()
for i in range(0, 7):
    tk.Checkbutton(window, text=temp[i][1], command=lambda: selector(temp[i][0])).pack()

window.mainloop()
print(makeupDate['월'])
print(makeupDate['화'])
print(makeupDate['수'])
print(makeupDate['목'])
print(makeupDate['금'])
print(makeupDate['토'])
print(makeupDate['일'])