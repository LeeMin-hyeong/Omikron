import calendar
from datetime import date
from dateutil.relativedelta import relativedelta

today = date.today()

makeupDate={}
makeupDate['월'] = today + relativedelta(weekday=calendar.MONDAY)
makeupDate['화'] = today + relativedelta(weekday=calendar.TUESDAY)
makeupDate['수'] = today + relativedelta(weekday=calendar.WEDNESDAY)
makeupDate['목'] = today + relativedelta(weekday=calendar.THURSDAY)
makeupDate['금'] = today + relativedelta(weekday=calendar.FRIDAY)
makeupDate['토'] = today + relativedelta(weekday=calendar.SATURDAY)
makeupDate['일'] = today + relativedelta(weekday=calendar.SUNDAY)
