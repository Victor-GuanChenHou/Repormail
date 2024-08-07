#!/usr/bin/env python3
# coding: utf-8

import smtplib
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import logging
from dotenv import load_dotenv
import os
import pandas as pd
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
file_path = './寄送郵件通訊錄.xlsx'
df = pd.read_excel(file_path)
sheet_names = ['杏子豬排', '大阪王將', '京都勝牛', '段純貞', '橋村','杏美小食堂']



# 登入憑證
username = os.getenv('mailusername')
password = os.getenv('mailpassword')



# 發送郵件
try:
    TIME=Time.lasttime()
    date = TIME[9]
    for i in range(len(df['寄送對象'])):
        msg = MIMEMultipart()
        msg['From'] = from_addr
        msg['To'] = df['寄送對象'][i]
        msg['Subject'] = date+'門市報表'
        msg.attach(MIMEText(date+'門市報表', 'plain', 'utf-8'))#設置郵件文檔
        for j in range(len(sheet_names)):
            if df[sheet_names[j]][i]=='yes':
                filepath ='./Senddata/'+sheet_names[j]+'_'+ date + '_Report.xlsx'
                filename =sheet_names[j]+'_'+ date + '_Report.xlsx'
                with open(filepath, 'rb') as attachment:
                    part = MIMEApplication(attachment.read(), Name=filename)
                    part['Content-Disposition'] = f'attachment; filename="{filename}"'
                    msg.attach(part)
        with smtplib.SMTP_SSL(smtp_server, port) as smtp:
            smtp.login(username, password)
            smtp.sendmail(from_addr, df['寄送對象'][i], msg.as_string())
    print("郵件傳送成功!")
    logging.info('Send Mail Success!')

except Exception as e:
    print(f"郵件傳送失敗: {e}")
    logging.error(f'Send Mail Error: {e}')

