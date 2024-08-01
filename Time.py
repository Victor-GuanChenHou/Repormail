#!/usr/bin/env python3
# coding: utf-8
import datetime
def lasttime():
    data=[]
    # 獲取當前日期和時間
    current_date = datetime.datetime.now()

    # 計算前一天的日期
    previous_date = current_date - datetime.timedelta(days=1)
    last_day = previous_date.day
    last_month = previous_date.month
    last_year = previous_date.year

    # 計算前一天的MTD和YTD範圍
    if current_date.day == 1:
        previous_month = current_date - datetime.timedelta(days=current_date.day)
        previous_month_first_day = datetime.date(previous_month.year, previous_month.month, 1)
    else:
        previous_month = current_date
        previous_month_first_day = datetime.date(previous_month.year, previous_month.month, 1)
    lastdate=datetime.date(previous_date.year, previous_month_first_day.month, last_day)
    data.append(lastdate)
    Mtd=datetime.date(previous_date.year, previous_month_first_day.month, 1)
    data.append(Mtd)
    Ytd=datetime.date(previous_date.year, 1, 1)
    data.append(Ytd)
    Plastdate=datetime.date(previous_date.year-1, previous_month_first_day.month, last_day)
    data.append(Plastdate)
    PMtd=datetime.date(previous_date.year-1, previous_month_first_day.month, 1)
    data.append(PMtd)
    PYtd=datetime.date(previous_date.year-1, 1, 1)
    data.append(PYtd)
    MTD = '(' + str(previous_month_first_day.month) + '/1~' + str(previous_month_first_day.month) + '/' + str(last_day) + ')'
    data.append(MTD)
    YTD = '(1/1~' + str(last_month) + '/' + str(last_day) + ')'
    data.append(YTD)
    # 格式化日期為 yyyymmdd
    datadate = str(previous_date.year) + str(previous_month_first_day.month).zfill(2) + str(last_day).zfill(2)
    data.append(datadate)
    # 打印結果
    date = previous_date.strftime('%Y-%m-%d')
    data.append(date)
    # print(data)
    # sample
    # [datetime.date(2024, 7, 31),    0lastdate
    #  datetime.date(2024, 7, 1),     1Mtd
    #  datetime.date(2024, 1, 1),     2Ytd
    #  datetime.date(2023, 7, 31),    3Plastdate
    #  datetime.date(2023, 7, 1),     4PMtd
    #  datetime.date(2023, 1, 1),     5PYtd
    #  '(7/1~7/31)',                  6MTD  str
    #  '(1/1~7/31)',                  7YTD  str
    #  '20240731',                    8lastdate     str
    #  '2024-07-31']                  9lastdate     str
    return data
    
