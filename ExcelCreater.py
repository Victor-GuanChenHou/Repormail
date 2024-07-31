#!/usr/bin/env python3
# coding: utf-8
import pandas as pd
import datetime
import SQL_tra as MYSQL
# 找前一天日期
date = str(datetime.datetime.now().year) + '-' + str(datetime.datetime.now().month) + '-' + str(int(datetime.datetime.now().day-1))
fileName ='./Senddata/'+ date + '_Report.xlsx'
lastdate=datetime.date(int(datetime.datetime.now().year),int(datetime.datetime.now().month),int(datetime.datetime.now().day)-1)
Mtd=datetime.date(int(datetime.datetime.now().year),int(datetime.datetime.now().month),1)
Ytd=datetime.date(int(datetime.datetime.now().year),1,1)

Plastdate=datetime.date(int(datetime.datetime.now().year)-1,int(datetime.datetime.now().month),int(datetime.datetime.now().day)-1)
PMtd=datetime.date(int(datetime.datetime.now().year)-1,int(datetime.datetime.now().month),1)
PYtd=datetime.date(int(datetime.datetime.now().year)-1,1,1)

# print(lastdate)
# print(Mtd)
# print(Ytd)
# print(Plastdate)
# print(PMtd)
# print(PYtd)
def datatitle(CYdata,PYdata,CYMTDdata,PYMTDdata,CYYTDdata,PYYTDdata):
    data=[]
    for i in range(len(CYdata)):
        if CYdata[i][0] not in data and CYdata[i][0] is not None:
            data.append(CYdata[i][0])
    for i in range(len(PYdata)):
        if PYdata[i][0] not in data and PYdata[i][0] is not None:
            data.append(PYdata[i][0])  
    for i in range(len(CYMTDdata)):
        if CYMTDdata[i][0] not in data and CYMTDdata[i][0] is not None:
            data.append(CYMTDdata[i][0])     
    for i in range(len(PYMTDdata)):
        if PYMTDdata[i][0] not in data and PYMTDdata[i][0] is not None:
            data.append(PYMTDdata[i][0])      
    for i in range(len(CYYTDdata)):
        if CYYTDdata[i][0] not in data and CYYTDdata[i][0] is not None:
            data.append(CYYTDdata[i][0])   
    for i in range(len(PYYTDdata)):
        if PYYTDdata[i][0] not in data and PYYTDdata[i][0] is not None:
            data.append(PYYTDdata[i][0])     
    return data
def datamerge(title,CYdata,PYdata):
    data=[]
    CYid=[i[0] for i in CYdata]
    PYid=[i[0] for i in PYdata]
    for i in range(len(title)):
         CYpositions = [index for index, value in enumerate(CYid) if value == title[i]]
         PYpositions = [index for index, value in enumerate(PYid) if value == title[i]]
         if CYpositions!=[] and PYpositions!=[]:
             # id name cysales pysales index TC TC_P TC_index TA TA_P TA_index
                thedata=(
                    CYdata[CYpositions[0]][0],CYdata[CYpositions[0]][1],
                    float(CYdata[CYpositions[0]][3]),
                    float(PYdata[PYpositions[0]][3]),
                    float(CYdata[CYpositions[0]][3])/float(PYdata[PYpositions[0]][3]),
                    int(int(CYdata[CYpositions[0]][4])),
                    int(int(PYdata[PYpositions[0]][4])),
                    float(CYdata[CYpositions[0]][4])/float(PYdata[PYpositions[0]][4]),
                    float(CYdata[CYpositions[0]][3])/float(CYdata[CYpositions[0]][4]),
                    float(PYdata[PYpositions[0]][3])/float(PYdata[PYpositions[0]][4]),
                    (float(CYdata[CYpositions[0]][3])/float(CYdata[CYpositions[0]][4]))/(float(PYdata[PYpositions[0]][3])/float(PYdata[PYpositions[0]][4])))
                data.append(thedata)
         elif CYpositions!=[] and PYpositions==[]:
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
    data=pd.DataFrame(data)
    data[5] = data[5].astype('Int64')
    data[6] = data[6].astype('Int64')
    return data

brand=['杏','大阪王將','勝牛','段純貞', '橋村','杏子小食堂']
df={}
for i in range(len(brand)):
    CYdata=MYSQL.searchdata(brand[i],lastdate,lastdate)
    PYdata=MYSQL.searchdata(brand[i],Plastdate,Plastdate)
    CYMTDdata=MYSQL.searchdata(brand[i],Mtd,lastdate)
    PYMTDdata=MYSQL.searchdata(brand[i],PMtd,Plastdate)
    CYYTDdata=MYSQL.searchdata(brand[i],Ytd,lastdate)
    PYYTDdata=MYSQL.searchdata(brand[i],PYtd,Plastdate)

    data_tile=datatitle(CYdata,PYdata,CYMTDdata,PYMTDdata,CYYTDdata,PYYTDdata)
    # 數據資料
    daily_data=datamerge(data_tile,CYdata,PYdata)
    MTD_data=datamerge(data_tile,CYMTDdata,PYMTDdata)
    YTD_data=datamerge(data_tile,CYYTDdata,PYYTDdata)

    data = {
        " ": daily_data[1],
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

    # 創建dataframe
    df[brand[i]]=(pd.DataFrame(data))


with pd.ExcelWriter(fileName, engine='xlsxwriter') as writer:
    sheet_names = ['杏子豬排', '大阪王將', '京都勝牛', '段純貞', '橋村','杏美小食堂']
    MTD= '('+str(datetime.datetime.now().month) + '/1~'+ str(datetime.datetime.now().month)+'/'+ str(int(datetime.datetime.now().day)-1)+')'
    YTD='(1/1~'+ str(datetime.datetime.now().month)+'/'+ str(int(datetime.datetime.now().day)-1)+')'
    print(MTD)
    print(YTD)
    for i in range(len(sheet_names)):
        df[brand[i]].to_excel(writer, sheet_name=sheet_names[i], index=False, startrow=1)
        workbook  = writer.book
        worksheet = writer.sheets[sheet_names[i]]
        
        # 合併儲存格
        worksheet.merge_range('B1:D1', 'Daily Sales', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.merge_range('E1:G1', 'Daily TC', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.merge_range('H1:J1', 'Daily TA', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.merge_range('K1:M1', 'MTD Sales'+str(MTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.merge_range('N1:P1', 'MTD TC Sales'+str(MTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.merge_range('Q1:S1', 'MTD TA Sales'+str(MTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.merge_range('T1:V1', 'YTD Sales'+str(YTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.merge_range('W1:Y1', 'YTD TC Sales'+str(YTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.merge_range('Z1:AB1', 'YTD TA Sales'+str(YTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))



        worksheet.write('B2', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('C2', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('D2', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('E2', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('F2', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('G2', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('H2', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.write('I2', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.write('J2', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))

        worksheet.write('K2', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('L2', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('M2', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('N2', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('O2', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('P2', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('Q2', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.write('R2', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.write('S2', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))

        worksheet.write('T2', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('U2', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('V2', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('W2', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('X2', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('Y2', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('Z2', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.write('AA2', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.write('AB2', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))

print("報表完成")
