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
datadate = str(datetime.datetime.now().year)  + str(datetime.datetime.now().month).zfill(2)  + str(int(datetime.datetime.now().day-1)).zfill(2)
folder_path = "./Origianldata"
csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')and f[:8]==datadate]
df_csv = pd.read_csv("./Origianldata/"+str(csv_files[len(csv_files)-1]), encoding='utf-16')
df_csv = pd.DataFrame(df_csv)
df_csv = df_csv.drop_duplicates(subset='so_id', keep='last')

install_data_column=[]
install_store_data=[]
install_sales_data=[]

for i in df_csv.index:
    try:
        if df_csv['so_type'][i]=='A'  and not np.isnan(df_csv['invoice_amt'][i]) :
            storename=data_segmentation(df_csv['store_name'][i])
            if df_csv['store_id'][i] not in install_data_column and storename[0]!=None :
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
                    install_sales_data[index]['total_customer']= install_sales_data[index]['total_customer']+1

                except:
                    try:
                        index = install_data_column.index(df_csv['store_id'][i])
                        install_sales_data[index]['invoice_amt']= install_sales_data[index]['invoice_amt']+(float(df_csv['invoice_amt'][i]*1))
                        install_sales_data[index]['total_customer']= install_sales_data[index]['total_customer']+1

                    except:
                        pass
    except:
        pass
MYSQL.insert_sales_data(install_sales_data)
MYSQL.insert_store_data(install_store_data)


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
#                 if df_csv['store_id'][i] not in install_data_column and storename[0]!=None :
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
#                         install_data_column.append(df_csv['store_id'][i])
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
#                             install_data_column.append(df_csv['store_id'][i])
#                         except:
#                             pass
#                 else:
#                     try:
#                         index = install_data_column.index(df_csv['store_id'][i])
#                         install_sales_data[index]['invoice_amt']= install_sales_data[index]['invoice_amt']+(float(df_csv['invoice_amt'][i]*1))
#                         install_sales_data[index]['total_customer']= install_sales_data[index]['total_customer']+1

#                     except:
#                         try:
#                             index = install_data_column.index(df_csv['store_id'][i])
#                             install_sales_data[index]['invoice_amt']= install_sales_data[index]['invoice_amt']+(float(df_csv['invoice_amt'][i]*1))
#                             install_sales_data[index]['total_customer']= install_sales_data[index]['total_customer']+1

#                         except:
#                             pass
#         except:
#             pass
#     MYSQL.insert_sales_data(install_sales_data)
#     MYSQL.insert_store_data(install_store_data)

