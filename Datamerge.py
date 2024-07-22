#!/usr/bin/env python
# coding: utf-8
import pandas as pd
import csv
import math
import datetime 
from openpyxl import load_workbook# 檢測編譯格式
import os
import numpy as np
import SQL_tra as MYSQL
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
folder_path = "./Origianldata"
csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]
df_csv = pd.read_csv("./Origianldata/"+str(csv_files[len(csv_files)-1]), encoding='utf-16')
df_csv = pd.DataFrame(df_csv)
df_csv = df_csv.drop_duplicates(subset='so_id', keep='last')

install_data_column=[]
install_store_data=[]
install_sales_data=[]

for i in df_csv.index:
    if df_csv['so_type'][i]=='A'  and not np.isnan(df_csv['invoice_amt'][i]):
        storename=data_segmentation(df_csv['store_name'][i])
        if df_csv['store_id'][i] not in install_data_column  :
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
                install_data_column.append(df_csv['store_id'][i])
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
                    install_data_column.append(df_csv['store_id'][i])
                except:
                    pass
        else:
            try:
                index = install_data_column.index(df_csv['store_id'][i])
                install_sales_data[index]['invoice_amt']= install_sales_data[index]['invoice_amt']+(float(df_csv['invoice_amt'][i]*1))
            except:
                try:
                    index = install_data_column.index(df_csv['store_id'][i])
                    install_sales_data[index]['invoice_amt']= install_sales_data[index]['invoice_amt']+(float(df_csv['invoice_amt'][i]*1))
                except:
                    pass

MYSQL.insert_sales_data(install_sales_data)
MYSQL.insert_store_data(install_store_data)

#存入MySQL
# print(install_store_data)
# print("LINE PAY")
# print(l)
# print("現金")
# print(c)
# print("信用卡")
# print(d)
# print("UE支付")
# print(u)
# print("一卡通")
# print(e)
# print("總和")

###創建新資料檔案
# df = pd.DataFrame(install_data)
# df.to_excel("./output.xlsx", sheet_name='Sheet1', index=False)
# workbook = load_workbook("./output.xlsx")
# worksheet = workbook["Sheet1"]
# worksheet.merge_cells(start_row=2, start_column=2, end_row=3, end_column=2)
# worksheet.cell(row=2, column=2).value = '杏'
# workbook.save("./output.xlsx")