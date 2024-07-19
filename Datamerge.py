#!/usr/bin/env python
# coding: utf-8
import pandas as pd
import csv
import math
import datetime 
from openpyxl import load_workbook# 檢測編譯格式
print(datetime.date(2020,1,2))
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
    print(date)
    dates=date.split('/')
    print(dates)
    rdates=datetime.date(dates[0],dates[1],dates[2])
    return rdates
###選取原始資料檔
df_csv = pd.read_csv("./Origianldata/20240702235509.csv", encoding='utf-16')
df_csv = pd.DataFrame(df_csv)
df_csv = df_csv.drop_duplicates(subset='so_id', keep='last')

install_data_column=[]
install_store_data=[]
install_sales_data=[]
#print(df_csv.columns)
# l=0
# d=0
# c=0
# u=0
# e=0
for i in df_csv.index:
    # if df_csv['store_id'][i]=='dc03022' and df_csv['payment_type'][i]=='LINE PAY':
    #         l=l+df_csv['invoice_amt'][i]
    # if df_csv['store_id'][i]=='dc03022' and df_csv['payment_type'][i]=='現金': 
    #         c=c+df_csv['invoice_amt'][i]
    # if df_csv['store_id'][i]=='dc03022' and df_csv['payment_type'][i]=='信用卡':       
    #         d=d+df_csv['invoice_amt'][i]
    # if df_csv['store_id'][i]=='dc03022' and df_csv['payment_type'][i]=='UE支付':       
    #         u=u+df_csv['invoice_amt'][i]
    # if df_csv['store_id'][i]=='dc03022' and df_csv['payment_type'][i]=='一卡通':       
    #         e=e+df_csv['invoice_amt'][i]

    if df_csv['so_type'][i]=='A' or df_csv['so_type'][i]=='B' :
        storename=data_segmentation(df_csv['store_name'][i])
        if df_csv['store_name'][i] not in install_data_column :
            try:
                
                data={
                'store_id':df_csv['store_id'][i],
                'brand':storename[0],
                'store_name':storename[1],
                'invoice_amt':(float(df_csv['invoice_amt'][i]*1)),
                'date':date_segmentation(df_csv['so_date'][i])
                }
                install_sales_data.append(data)
                install_data_column.append(df_csv['store_name'][i])
            except:
                try:
                    data={
                    'store_id':df_csv['store_id'][i],
                    'brand':storename[0],
                    'store_name':storename[1],
                    'invoice_amt':(float(df_csv['invoice_amt'][i]*1)),
                    'date':date_segmentation(df_csv['so_date'][i])

                    }
                    install_sales_data.append(data)
                    install_data_column.append(df_csv['store_name'][i])
                except:
                    pass
        else:
            try:
                index = install_data_column.index(df_csv['store_name'][i])
                install_sales_data[index]['invoice_amt']= install_sales_data[index]['invoice_amt']+(float(df_csv['invoice_amt'][i]*1))
            except:
                try:
                    index = install_data_column.index(df_csv['store_name'][i])
                    install_sales_data[index]['invoice_amt']= install_sales_data[index]['invoice_amt']+(float(df_csv['invoice_amt'][i]*1))
                except:
                    pass
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
print(install_sales_data)
###創建新資料檔案
# df = pd.DataFrame(install_data)
# df.to_excel("./output.xlsx", sheet_name='Sheet1', index=False)
# workbook = load_workbook("./output.xlsx")
# worksheet = workbook["Sheet1"]
# worksheet.merge_cells(start_row=2, start_column=2, end_row=3, end_column=2)
# worksheet.cell(row=2, column=2).value = '杏'
# workbook.save("./output.xlsx")