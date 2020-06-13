import time,datetime

DATE_FORMAT_YMD = '%Y-%m-%d'


def date2TimeStamp(date,date_format = DATE_FORMAT_YMD):
    timeArray = time.strptime(date, date_format)
    timeStamp = int(time.mktime(timeArray)) * 1000
    return timeStamp

def timeStamp2Date(time_stamp,date_format = DATE_FORMAT_YMD):
    timeArray = time.localtime(time_stamp)
    return time.strftime(date_format, timeArray)

def oneDaySecond():
    return (24*60*60)
#print(timeStamp2Date(1557502800))

