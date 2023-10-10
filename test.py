import datetime

# date1 = '2024-01-01'
# date1 = datetime.datetime.strptime(date1, '%Y-%m-%d')
# print(datetime.datetime(date1.year - 1, 12, 1))
# print(date1.year==2024)

from rpaAbout import rpaPyScriptFunc
from rpaAbout import rpaPyScriptFunc

a = rpaPyScriptFunc.dateTimeAbout(dateTime=datetime.datetime.now()).nowDateTime()
b = rpaPyScriptFunc.dateTimeAbout(dateTime=datetime.datetime.now()).nowWeek()
print(a)
print(b)
