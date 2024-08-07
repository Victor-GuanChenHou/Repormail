#!/usr/bin/env python3
# coding: utf-8
import pandas as pd
import datetime
import SQL_tra as MYSQL
import Time as Time
import numpy as np

TIME=Time.lasttime()

date = TIME[9]
lastdate=TIME[0]
Mtd=TIME[1]
Ytd=TIME[2]

Plastdate=TIME[3]
PMtd=TIME[4]
PYtd=TIME[5]

# print(lastdate)
# print(Mtd)
# print(Ytd)
# print(Plastdate)
# print(PMtd)
# print(PYtd)
def datatitle(brand):
    # 讀取 Excel 文件
    file_path = './王座國際門市通訊錄.xlsx'
    df = pd.read_excel(file_path)

    # 根據品牌和營運主管進行分組
    data=[]
    name=[]
    for i in range(len(df['品牌'])):
        if df['品牌'][i]==brand:
            if df['營運主管'][i] not in name:
                name.append(df['營運主管'][i])
    # thedata={
    #     '品牌':'全品牌總和',
    #     'POS店名':'全品牌總和',
    #     '店名':'全品牌總和',
    #     '營運主管':'全品牌總和'
    # }
    # data.append(thedata)
    thedata={
        '品牌':'Total',
        'POS店名':'Total',
        '店名':'Total',
        '營運主管':' '
    }
    data.append(thedata)
    for j in range(len(name)):
        for i in range(len(df['品牌'])):
            if df['品牌'][i]==brand and name[j]==df['營運主管'][i]:
                    thedata={
                        '品牌':df['品牌'][i],
                        'POS店名':df['POS店名'][i],
                        '店名':df['店名'][i],
                        '營運主管':df['營運主管'][i]
                    }
                    data.append(thedata)
    data=pd.DataFrame(data)  
    return data
def datamerge(title,CYdata,PYdata):
    data=[]
    CYid=[i[1] for i in CYdata]
    PYid=[i[1] for i in PYdata]
    CYtotal=0
    PYtotal=0
    CYTC=0
    PYTC=0
    for i in range(len(title)):
         CYpositions = [index for index, value in enumerate(CYid) if value == title[i]]
         PYpositions = [index for index, value in enumerate(PYid) if value == title[i]]
         if CYpositions!=[] and PYpositions!=[]:
                CYtotal=CYtotal+CYdata[CYpositions[0]][3]
                CYTC=CYTC+CYdata[CYpositions[0]][4]
                PYtotal=PYtotal+PYdata[PYpositions[0]][3]
                PYTC=PYTC+PYdata[PYpositions[0]][4]
             # id name cysales pysales index TC TC_P TC_index TA TA_P TA_index
                thedata=(
                    CYdata[CYpositions[0]][0],CYdata[CYpositions[0]][1],
                    float(CYdata[CYpositions[0]][3]),
                    float(PYdata[PYpositions[0]][3]),
                    (float(CYdata[CYpositions[0]][3])/float(PYdata[PYpositions[0]][3]))*100,
                    int(int(CYdata[CYpositions[0]][4])),
                    int(int(PYdata[PYpositions[0]][4])),
                    (float(CYdata[CYpositions[0]][4])/float(PYdata[PYpositions[0]][4]))*100,
                    float(CYdata[CYpositions[0]][3])/float(CYdata[CYpositions[0]][4]),
                    float(PYdata[PYpositions[0]][3])/float(PYdata[PYpositions[0]][4]),
                    ((float(CYdata[CYpositions[0]][3])/float(CYdata[CYpositions[0]][4]))/(float(PYdata[PYpositions[0]][3])/float(PYdata[PYpositions[0]][4])))*100
                )
                data.append(thedata)
         elif CYpositions!=[] and PYpositions==[]:
             CYtotal=CYtotal+CYdata[CYpositions[0]][3]
             CYTC=CYTC+CYdata[CYpositions[0]][4]
             thedata=(
                CYdata[CYpositions[0]][0],
                CYdata[CYpositions[0]][1],
                float(CYdata[CYpositions[0]][3]),
                None,
                None,
                int(CYdata[CYpositions[0]][4]),
                None,
                None,
                float(CYdata[CYpositions[0]][3])/float(CYdata[CYpositions[0]][4]),
                None,
                None)
             data.append(thedata)
         else:
             thedata=(
                title[i],
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None)
             data.append(thedata)
             
             
    if PYtotal !=0:
        thedata=(
                "Total","Total",
                float(CYtotal),
                float(PYtotal),
                (float(CYtotal)/float(PYtotal))*100,
                int(int(CYTC)),
                int(int(PYTC)),
                (float(CYTC)/float(PYTC))*100,
                float(CYtotal)/float(CYTC),
                float(PYtotal)/float(PYTC),
                ((float(CYtotal)/float(CYTC))/(float(PYtotal)/float(PYTC)))*100
        )
    else:
        thedata=(
                "Total","Total",
                float(CYtotal),
                None,
                None,
                int(int(CYTC)),
                None,
                None,
                float(CYtotal)/float(CYTC),
                None,
                None
        )
    data.insert(0,thedata)
    data=pd.DataFrame(data)
    data[5] = data[5].astype('Int64')
    data[6] = data[6].astype('Int64')
    return data

brand=['杏','大阪王將','勝牛','段純貞', '橋村','杏子小食堂']
Nuberofstore=[]
df={}

for i in range(len(brand)):
    total_daily_sales_cy=0
    total_daily_sales_py=0
    total_daily_sales_index=0
    total_daily_tc_cy=0
    total_daily_tc_py=0
    total_daily_tc_index=0
    total_daily_ta_cy=0
    total_daily_ta_py=0
    total_daily_ta_index=0
    total_mtd_sales_cy=0
    total_mtd_sales_py=0
    total_mtd_sales_index=0
    total_mtd_tc_cy=0
    total_mtd_tc_py=0
    total_mtd_tc_index=0
    total_mtd_ta_cy=0
    total_mtd_ta_py=0
    total_mtd_ta_index=0
    total_ytd_sales_cy=0
    total_ytd_sales_py=0
    total_ytd_sales_index=0
    total_ytd_tc_cy=0
    total_ytd_tc_py=0
    total_ytd_tc_index=0
    total_ytd_ta_cy=0
    total_ytd_ta_py=0
    total_ytd_ta_index=0
    CYdata=MYSQL.searchdata(brand[i],lastdate,lastdate)
    PYdata=MYSQL.searchdata(brand[i],Plastdate,Plastdate)
    CYMTDdata=MYSQL.searchdata(brand[i],Mtd,lastdate)
    PYMTDdata=MYSQL.searchdata(brand[i],PMtd,Plastdate)
    CYYTDdata=MYSQL.searchdata(brand[i],Ytd,lastdate)
    PYYTDdata=MYSQL.searchdata(brand[i],PYtd,Plastdate)

    data_tile=datatitle(brand[i])
    # 數據資料
    daily_data=datamerge(data_tile['店名'],CYdata,PYdata)
    MTD_data=datamerge(data_tile['店名'],CYMTDdata,PYMTDdata)
    YTD_data=datamerge(data_tile['店名'],CYYTDdata,PYYTDdata)

    data = {
        " ": data_tile['店名'],
        "營運主管":data_tile['營運主管'],
        "Daily Sales CY": daily_data[2],
        "Daily Sales PY": daily_data[3],
        "Daily Sales Index":daily_data[4],
        "Daily TC CY":daily_data[5],
        "Daily TC PY":daily_data[6],
        "Daily TC Index":daily_data[7],
        "Daily TA CY":daily_data[8],
        "Daily TA PY":daily_data[9],
        "Daily TA Index":daily_data[10],
        "MTD Sales CY": MTD_data[2],
        "MTD Sales PY":MTD_data[3],
        "MTD Sales Index": MTD_data[4],
        "MTD TC CY":MTD_data[5],
        "MTD TC PY":MTD_data[6],
        "MTD TC Index":MTD_data[7],
        "MTD TA CY":MTD_data[8],
        "MTD TA PY":MTD_data[9],
        "MTD TA Index":MTD_data[10],
        "YTD Sales CY": YTD_data[2],
        "YTD Sales PY": YTD_data[3],
        "YTD Sales Index": YTD_data[4],
        "YTD TC CY":YTD_data[5],
        "YTD TC PY":YTD_data[6],
        "YTD TC Index":YTD_data[7],
        "YTD TA CY":YTD_data[8],
        "YTD TA PY":YTD_data[9],
        "YTD TA Index":YTD_data[10],
    }
    num=len(daily_data[1])-1
    Nuberofstore.append(num)
    # 創建dataframe
    df[brand[i]]=(pd.DataFrame(data))
    #所有品牌業績加總
    total_daily_sales_cy+= sum(0 if x is np.nan or x is None or x is pd.NA else x for x in data["Daily Sales CY"])
    total_daily_sales_py+= sum(0 if x is np.nan or x is None or x is pd.NA else x for x in data["Daily Sales PY"])
    if total_daily_sales_py == 0:
        total_daily_sales_py = None
        total_daily_sales_index = None
    else:
        total_daily_sales_index = (total_daily_sales_cy / total_daily_sales_py) * 100
    total_daily_tc_cy+= sum(0 if x is np.nan or x is None or x is pd.NA else x for x in data["Daily TC CY"])
    total_daily_tc_py+= sum(0 if x is np.nan or x is None or x is pd.NA else x for x in data["Daily TC PY"])
    if total_daily_tc_py == 0:
        total_daily_tc_py = None
        total_daily_tc_index = None
    else:
        total_daily_tc_index = (total_daily_tc_cy / total_daily_tc_py) * 100
    total_daily_ta_cy=total_daily_sales_cy/total_daily_tc_cy
    if total_daily_tc_py==None or total_daily_sales_py==None:
        total_daily_ta_py=None
        total_daily_ta_index=None
    else:
        total_daily_ta_py=total_daily_sales_py/total_daily_tc_py
        total_daily_ta_index=(total_daily_ta_cy/total_daily_ta_py)*100
    total_mtd_sales_cy+= sum(0 if x is np.nan or x is None or x is pd.NA else x for x in data["MTD Sales CY"])
    total_mtd_sales_py+= sum(0 if x is np.nan or x is None or x is pd.NA else x for x in data["MTD Sales PY"])
    if total_mtd_sales_py==0:
         total_mtd_sales_py=None
         total_mtd_sales_index=None
    else:
        total_mtd_sales_index=(total_mtd_sales_cy/total_mtd_sales_py)*100
    total_mtd_tc_cy+= sum(0 if x is np.nan or x is None or x is pd.NA else x for x in data["MTD TC CY"])
    total_mtd_tc_py+= sum(0 if x is np.nan or x is None or x is pd.NA else x for x in data["MTD TC PY"])
    if total_mtd_tc_py==0:
         total_mtd_tc_py=None
         total_mtd_tc_index=None
    else:
        total_mtd_tc_index=(total_mtd_tc_cy/total_mtd_tc_py)*100
    total_mtd_ta_cy=total_mtd_sales_cy/total_mtd_tc_cy
    if total_mtd_tc_py==None or total_mtd_sales_py==None:
         total_mtd_ta_py=None
         total_mtd_ta_index=None
    else:
        total_mtd_ta_py=total_mtd_sales_py/total_mtd_tc_py
        total_mtd_ta_index=(total_mtd_ta_cy/total_mtd_ta_py)*100
    total_ytd_sales_cy+= sum(0 if x is np.nan or x is None or x is pd.NA else x for x in data["YTD Sales CY"])
    total_ytd_sales_py+= sum(0 if x is np.nan or x is None or x is pd.NA else x for x in data["YTD Sales PY"])
    if total_ytd_sales_py==0:
         total_ytd_sales_py=None
         total_ytd_sales_index=None
    else:
        total_ytd_sales_index=(total_ytd_sales_cy/total_ytd_sales_py)*100
    total_ytd_tc_cy+= sum(0 if x is np.nan or x is None or x is pd.NA else x for x in data["YTD TC CY"])
    total_ytd_tc_py+= sum(0 if x is np.nan or x is None or x is pd.NA else x for x in data["YTD TC PY"])
    if total_ytd_tc_py==0:
         total_ytd_tc_py=None
         total_ytd_tc_index=None
    else:
        total_ytd_tc_index=(total_ytd_tc_cy/total_ytd_tc_py)*100
    total_ytd_ta_cy=total_ytd_sales_cy/total_ytd_tc_cy
    if total_ytd_tc_py==None or total_ytd_sales_py==None:
        total_ytd_ta_py=None
        total_ytd_ta_index=None
    else:
        total_ytd_ta_py=total_ytd_sales_py/total_ytd_tc_py
        total_ytd_ta_index=(total_ytd_ta_cy/total_ytd_ta_py)*100

for i in range(len(brand)):
    df[brand[i]].loc[-1] = ["全品牌",total_daily_sales_cy,total_daily_sales_py,total_daily_sales_index,total_daily_tc_cy,total_daily_tc_py, total_daily_tc_index,total_daily_ta_cy,total_daily_ta_py,total_daily_ta_index,total_mtd_sales_cy,total_mtd_sales_py, total_mtd_sales_index,total_mtd_tc_cy, total_mtd_tc_py, total_mtd_tc_index,total_mtd_ta_cy,total_mtd_ta_py,total_mtd_ta_index,total_ytd_sales_cy,total_ytd_sales_py,total_ytd_sales_index,total_ytd_tc_cy,total_ytd_tc_py,total_ytd_tc_index,total_ytd_ta_cy,total_ytd_ta_py,total_ytd_ta_index] 
    df[brand[i]].index = df[brand[i]].index + 1  
    df[brand[i]].sort_index(inplace=True)  


sheet_names = ['杏子豬排', '大阪王將', '京都勝牛', '段純貞', '橋村','杏美小食堂']
for i in range(len(sheet_names)):
    fileName ='./Senddata/'+sheet_names[i]+'_'+ date + '_Report.xlsx'
    with pd.ExcelWriter(str(fileName), engine='xlsxwriter') as writer:
            MTD=TIME[6]
            YTD=TIME[7]
            print(MTD)
            print(YTD)
            df[brand[i]].to_excel(writer, sheet_name=sheet_names[i], index=False, startrow=6)
            workbook  = writer.book
            worksheet = writer.sheets[sheet_names[i]]
            worksheet.write('A1', TIME[9])
            worksheet.write('A2', '店數')
            worksheet.write('B2', Nuberofstore[i])
            # 合併儲存格
            worksheet.merge_range('A1:B1', TIME[9], workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True}))
            
            worksheet.merge_range('C6:E6', 'Daily Sales', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
            worksheet.merge_range('F6:H6', 'Daily TC', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
            worksheet.merge_range('I6:K6', 'Daily TA', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
            worksheet.merge_range('L6:N6', 'MTD Sales'+str(MTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
            worksheet.merge_range('O6:Q6', 'MTD TC Sales'+str(MTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
            worksheet.merge_range('R6:T6', 'MTD TA Sales'+str(MTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
            worksheet.merge_range('U6:W6', 'YTD Sales'+str(YTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
            worksheet.merge_range('X6:Z6', 'YTD TC Sales'+str(YTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
            worksheet.merge_range('AA6:AC6', 'YTD TA Sales'+str(YTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))



            worksheet.write('C7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
            worksheet.write('D7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
            worksheet.write('E7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
            worksheet.write('F7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
            worksheet.write('G7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
            worksheet.write('H7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
            worksheet.write('I7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
            worksheet.write('J7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
            worksheet.write('K7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))

            worksheet.write('L7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
            worksheet.write('M7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
            worksheet.write('N7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
            worksheet.write('O7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
            worksheet.write('P7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
            worksheet.write('Q7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
            worksheet.write('R7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
            worksheet.write('S7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
            worksheet.write('T7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))

            worksheet.write('U7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
            worksheet.write('V7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
            worksheet.write('W7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
            worksheet.write('X7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
            worksheet.write('Y7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
            worksheet.write('Z7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
            worksheet.write('AA7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
            worksheet.write('AB7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
            worksheet.write('AC7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))

            # 設置B8到U8以下100格的格式
            for col in range(1,26):
                col_letter = chr(ord('B') + col - 1)
                worksheet.set_column(f'{col_letter}8:{col_letter}100', 11,  workbook.add_format({'num_format': '#,##0', 'align': 'right'}))
            worksheet.set_column('AA8:AA100', 11,  workbook.add_format({'num_format': '#,##0', 'align': 'right'}))
            worksheet.set_column('AB8:AB100', 11,  workbook.add_format({'num_format': '#,##0', 'align': 'right'}))
            worksheet.set_column('A1:A100',16)

print("報表完成")
