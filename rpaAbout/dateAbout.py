import datetime

class dateTimeAbout:

    def __init__(self, dateTime):
        if type(dateTime) is str:
            self.dateTime = dateTime.replace('.', '').replace('-', '').replace('/', '').replace(' ','')
            if len(self.dateTime) == 8:
                self.dateTime = datetime.datetime.strptime(self.dateTime, '%Y%m%d')
            elif len(self.dateTime) == 12:
                self.dateTime = datetime.datetime.strptime(self.dateTime, '%Y%m%d%H%M')
            elif len(self.dateTime) == 14:
                self.dateTime = datetime.datetime.strptime(self.dateTime, '%Y%m%d%H%M%S')
        else:
            self.dateTime = dateTime

    # 当前时间
    def nowDateTime(self):
        return self.dateTime

    # 前一天
    def lastDay(self):
        lastDay = self.dateTime + datetime.timedelta(days=-1)
        return lastDay

    # 当前日期上月的第一天
    def lastMonthFirstDay(self):
        if self.dateTime.month == 1:
            lastMonthFirstDay = datetime.datetime(self.dateTime.year - 1, 12 ,1)
        else:
            lastMonthFirstDay = datetime.datetime(self.dateTime.year, self.dateTime.month - 1, 1)
        return lastMonthFirstDay

    # 当前日期对应的星期
    def nowWeek(self):
        print(self.dateTime)
        # 格式化输出当前日期时间
        t = self.dateTime.strftime('%Y%m%d%H%M%S')
        Weekday_Beijing = (datetime.datetime.strptime(t,'%Y%m%d%H%M%S') + datetime.timedelta(days=1)).weekday()
        Weekday_Greenwich = datetime.datetime.utcnow().weekday()
        return Weekday_Beijing


