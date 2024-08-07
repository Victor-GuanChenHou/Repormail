#!/usr/bin/env python3
# coding: utf-8
import os
import datetime
os.system("python3 ./SFTPgetdata.py")
os.system("python3 ./Datamerge.py")
os.system('python3 ./ExcelCreater.py')
wait=True
while wait:
    current_time = datetime.datetime.now()
    am9 = current_time.replace(hour=9, minute=0, second=0, microsecond=0)
    am930 = current_time.replace(hour=9, minute=30, second=0, microsecond=0)
    if current_time>=am9 and current_time<=am930:
        wait=False
os.system('python3 ./reportmail.py')