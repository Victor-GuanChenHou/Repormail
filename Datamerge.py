#!/usr/bin/env python3
# coding: utf-8
import pandas as pd
import csv
import math
import datetime 
from openpyxl import load_workbook# 檢測編譯格式
import os
import numpy as np
import SQL_tra as MYSQL
import Time as Time
####
import smtplib
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
# import chardet
# with open("./Origianldata/20240702235509.csv", 'rb') as f:
#     result = chardet.detect(f.read())
#     encoding = result['encoding']
#     print(f"檢測到的編碼: {encoding}")

###品牌與店名切割
def data_segmentation(name):
    names=name.split("-")
    rname=[]
    if names[1]=='杏':
        rname.append(names[1])
        rname.append(names[0])
    else:
        rname=names
    return rname
def date_segmentation(date):
    data=date.split(' ')
    dates=data[0].split('/')
    rdates=str(dates[0]+'-'+dates[1]+'-'+dates[2])
    return rdates
###選取原始資料檔
TIME=Time.lasttime()
datadate = TIME[8]
folder_path = "./Origianldata"
csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')and f[:8]==datadate]
df_csv = pd.read_csv("./Origianldata/"+str(csv_files[len(csv_files)-1]), encoding='utf-16')
df_csv = pd.DataFrame(df_csv)
df_csv = df_csv.drop_duplicates(subset='so_id', keep='last')
store_csv = pd.read_excel("./王座國際門市通訊錄.xlsx")
store_csv = pd.DataFrame(store_csv)
sd=[]
for i in range(len(store_csv)):
    if store_csv['品牌'][i]=='杏':
        d=store_csv['POS店名'][i]+'-'+store_csv['品牌'][i]
        sd.append(d)
    elif store_csv['品牌'][i]=='杏子小食堂' and store_csv['POS店名'][i]=='台北101':
        d='杏美小食堂-'+store_csv['POS店名'][i]
        sd.append(d)
    else:
        d=d=store_csv['品牌'][i]+'-'+store_csv['POS店名'][i]
        sd.append(d)
install_data_column=[]
install_store_data=[]
install_sales_data=[]

for i in df_csv.index:
    try:
        if df_csv['so_type'][i]=='A'  and not np.isnan(df_csv['invoice_amt'][i]) :
            storename=data_segmentation(df_csv['store_name'][i])
            if df_csv['store_name'][i] in sd:
                sd.remove(df_csv['store_name'][i])
            if {'store_id': df_csv['store_id'][i], 'date': date_segmentation(df_csv['so_date'][i])} not in install_data_column and storename[0]!=None:
                try:
                    
                    sales_data={
                    'store_id':df_csv['store_id'][i],
                    'invoice_amt':(float(df_csv['invoice_amt'][i]*1)),
                    'total_customer':1,
                    'DATE':date_segmentation(df_csv['so_date'][i])
                    
                    }
                    store_data={
                        'store_id':df_csv['store_id'][i],
                        'brand':storename[0],
                        'store_name':storename[1]

                    }
                    install_store_data.append(store_data)
                    install_sales_data.append(sales_data)
                    columndata={'store_id': df_csv['store_id'][i], 'date': date_segmentation(df_csv['so_date'][i])}
                    install_data_column.append(columndata)
                except:
                    try:
                        sales_data={
                        'store_id':df_csv['store_id'][i],
                        'invoice_amt':(float(df_csv['invoice_amt'][i]*1)),
                        'total_customer':1,
                        'date':date_segmentation(df_csv['so_date'][i])
                        
                        }
                        store_data={
                            'store_id':df_csv['store_id'][i],
                            'brand':storename[0],
                            'store_name':storename[1]

                        }
                        install_store_data.append(store_data)
                        install_sales_data.append(sales_data)
                        columndata={'store_id': df_csv['store_id'][i], 'date': date_segmentation(df_csv['so_date'][i])}
                        install_data_column.append(columndata)
                    except:
                        pass
            else:
                try:
                    index = install_data_column.index({'store_id': df_csv['store_id'][i], 'date': date_segmentation(df_csv['so_date'][i])})
                    install_sales_data[index]['invoice_amt']= install_sales_data[index]['invoice_amt']+(float(df_csv['invoice_amt'][i]*1))
                    install_sales_data[index]['total_customer']= install_sales_data[index]['total_customer']+1

                except:
                    try:
                        index = install_data_column.index({'store_id': df_csv['store_id'][i], 'date': date_segmentation(df_csv['so_date'][i])})
                        install_sales_data[index]['invoice_amt']= install_sales_data[index]['invoice_amt']+(float(df_csv['invoice_amt'][i]*1))
                        install_sales_data[index]['total_customer']= install_sales_data[index]['total_customer']+1

                    except:
                        pass
    except:
        pass

MYSQL.insert_sales_data(install_sales_data)
MYSQL.insert_store_data(install_store_data)
if len(sd)!=0:
    smtp_server = 'mail.kingza.com.tw'
    port = 465
    from_addr = 'kingzareport@kingza.com.tw'
    username = os.getenv('mailusername')
    password = os.getenv('mailpassword')
    sendmember=['victor.hou@kingza.com.tw']
    date=TIME[9]
    for j in range(len(sendmember)):
        msg = MIMEMultipart()
        msg['From'] = from_addr
        msg['To'] = sendmember[j]
        msg['Subject'] = date+'門市報表異常'
        # 生成門市報表的內容
        store_content = "\n".join([f"{i+1}.{sd[i]}" for i in range(len(sd))])
        email_body = f"{date} 門市報表異常名單\n\n{store_content}"
        # 附加到郵件主體
        msg.attach(MIMEText(email_body, 'plain', 'utf-8'))
        with smtplib.SMTP_SSL(smtp_server, port) as smtp:
                smtp.login(username, password)
                smtp.sendmail(from_addr, sendmember[j], msg.as_string())


###存入多天資料

# folder_path = "./Origianldata"
# csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]
# for z in range(len(csv_files)):
#     df_csv = pd.read_csv("./Origianldata/"+str(csv_files[z]), encoding='utf-16')
#     df_csv = pd.DataFrame(df_csv)
#     df_csv = df_csv.drop_duplicates(subset='so_id', keep='last')

#     install_data_column=[]
#     install_store_data=[]
#     install_sales_data=[]

#     for i in df_csv.index:
#         try:    
#             if df_csv['so_type'][i]=='A'  and not np.isnan(df_csv['invoice_amt'][i]):
#                 storename=data_segmentation(df_csv['store_name'][i])
#                 if {'store_id': df_csv['store_id'][i], 'date': date_segmentation(df_csv['so_date'][i])} not in install_data_column and storename[0]!=None:
#                     try:
                        
#                         sales_data={
#                         'store_id':df_csv['store_id'][i],
#                         'invoice_amt':(float(df_csv['invoice_amt'][i]*1)),
#                         'total_customer':1,
#                         'DATE':date_segmentation(df_csv['so_date'][i])
                        
#                         }
#                         store_data={
#                             'store_id':df_csv['store_id'][i],
#                             'brand':storename[0],
#                             'store_name':storename[1]

#                         }
#                         install_store_data.append(store_data)
#                         install_sales_data.append(sales_data)
#                         columndata={'store_id': df_csv['store_id'][i], 'date': date_segmentation(df_csv['so_date'][i])}
#                         install_data_column.append(columndata)
#                     except:
#                         try:
#                             sales_data={
#                             'store_id':df_csv['store_id'][i],
#                             'invoice_amt':(float(df_csv['invoice_amt'][i]*1)),
#                             'total_customer':1,
#                             'date':date_segmentation(df_csv['so_date'][i])
                            
#                             }
#                             store_data={
#                                 'store_id':df_csv['store_id'][i],
#                                 'brand':storename[0],
#                                 'store_name':storename[1]

#                             }
#                             install_store_data.append(store_data)
#                             install_sales_data.append(sales_data)
#                             columndata={'store_id': df_csv['store_id'][i], 'date': date_segmentation(df_csv['so_date'][i])}
#                             install_data_column.append(columndata)
#                         except:
#                             pass
#                 else:
#                     try:
#                         index = install_data_column.index({'store_id': df_csv['store_id'][i], 'date': date_segmentation(df_csv['so_date'][i])})
#                         install_sales_data[index]['invoice_amt']= install_sales_data[index]['invoice_amt']+(float(df_csv['invoice_amt'][i]*1))
#                         install_sales_data[index]['total_customer']= install_sales_data[index]['total_customer']+1

#                     except:
#                         try:
#                             index = install_data_column.index({'store_id': df_csv['store_id'][i], 'date': date_segmentation(df_csv['so_date'][i])})
#                             install_sales_data[index]['invoice_amt']= install_sales_data[index]['invoice_amt']+(float(df_csv['invoice_amt'][i]*1))
#                             install_sales_data[index]['total_customer']= install_sales_data[index]['total_customer']+1

#                         except:
#                             pass
#         except:
#             pass
#     MYSQL.insert_sales_data(install_sales_data)
#     MYSQL.insert_store_data(install_store_data)


