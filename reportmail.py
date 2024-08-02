#!/usr/bin/env python3
# coding: utf-8

import smtplib
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import logging
from dotenv import load_dotenv
import os
import datetime
import Time as Time
load_dotenv()
logging.basicConfig(
    filename='report_mail.log',
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)



# 設置 SMTP 伺服器和端口
smtp_server = 'mail.kingza.com.tw'
port = 465

# 設置發送者和接收者的電子郵件地址
from_addr = 'kingzareport@kingza.com.tw'
to_addr = ['victor.hou@kingza.com.tw']#寄送對象




# 登入憑證
username = os.getenv('mailusername')
password = os.getenv('mailpassword')


#傳送檔案
try:
    TIME=Time.lasttime()
    date = TIME[9]
    filepath='./Senddata/'+ date + '_Report.xlsx'
    filename= date + '_Report.xlsx'
except Exception as e:
    print(f"郵件傳送失敗: {e}")
    logging.error(f'Send Mail Error: {e}')
# 發送郵件
try:
    for i in range(len(to_addr)):
        
        msg = MIMEMultipart()
        msg['From'] = from_addr
        msg['To'] = to_addr[i]
        msg['Subject'] = 'Daily Report'
        msg.attach(MIMEText('Daily Report', 'plain', 'utf-8'))#設置郵件文檔
        with open(filepath, 'rb') as attachment:
            part = MIMEApplication(attachment.read(), Name=filename)
            part['Content-Disposition'] = f'attachment; filename="{filename}"'
            msg.attach(part)
        with smtplib.SMTP_SSL(smtp_server, port) as smtp:
            smtp.login(username, password)
            smtp.sendmail(from_addr, to_addr[i], msg.as_string())
    print("郵件傳送成功!")
    logging.info('Send Mail Success!')

except Exception as e:
    print(f"郵件傳送失敗: {e}")
    logging.error(f'Send Mail Error: {e}')

