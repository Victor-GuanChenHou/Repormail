#!/usr/bin/env python3
# coding: utf-8
import pandas as pd
import datetime
import SQL_tra as MYSQL
import Time as Time

TIME=Time.lasttime()

date = TIME[9]
fileName ='./Senddata/'+ date + '_Report.xlsx'
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
def datatitle(CYdata,PYdata,CYMTDdata,PYMTDdata,CYYTDdata,PYYTDdata):
    data=['全品牌總和','total']
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
                    (float(float(CYdata[CYpositions[0]][3])/float(CYdata[CYpositions[0]][4]))/(float(PYdata[PYpositions[0]][3])/float(PYdata[PYpositions[0]][4]))))*100
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
    if PYtotal !=0:
        thedata=(
                "total","total",
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
                "total","total",
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
    num=len(daily_data[1])-1
    Nuberofstore.append(num)
    # 創建dataframe
    df[brand[i]]=(pd.DataFrame(data))


with pd.ExcelWriter(fileName, engine='xlsxwriter') as writer:
    sheet_names = ['杏子豬排', '大阪王將', '京都勝牛', '段純貞', '橋村','杏美小食堂']
    MTD=TIME[6]
    YTD=TIME[7]
    print(MTD)
    print(YTD)
    for i in range(len(sheet_names)):
        df[brand[i]].to_excel(writer, sheet_name=sheet_names[i], index=False, startrow=6)
        workbook  = writer.book
        worksheet = writer.sheets[sheet_names[i]]
        worksheet.write('A1', TIME[9])
        worksheet.write('A2', 'CY店數')
        worksheet.write('B2', Nuberofstore[i])
        # 合併儲存格
        worksheet.merge_range('A1:B1', TIME[9], workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True}))
        
        worksheet.merge_range('B6:D6', 'Daily Sales', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.merge_range('E6:G6', 'Daily TC', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.merge_range('H6:J6', 'Daily TA', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.merge_range('K6:M6', 'MTD Sales'+str(MTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.merge_range('N6:P6', 'MTD TC Sales'+str(MTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.merge_range('Q6:S6', 'MTD TA Sales'+str(MTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.merge_range('T6:V6', 'YTD Sales'+str(YTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.merge_range('W6:Y6', 'YTD TC Sales'+str(YTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.merge_range('Z6:AB6', 'YTD TA Sales'+str(YTD), workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))



        worksheet.write('B7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('C7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('D7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('E7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('F7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('G7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('H7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.write('I7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.write('J7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))

        worksheet.write('K7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('L7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('M7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('N7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('O7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('P7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('Q7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.write('R7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.write('S7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))

        worksheet.write('T7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('U7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('V7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800000','font_color': '#FFFFFF'}))
        worksheet.write('W7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('X7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('Y7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#008080','font_color': '#FFFFFF'}))
        worksheet.write('Z7', 'CY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.write('AA7', 'PY', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))
        worksheet.write('AB7', 'Index', workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,'bg_color': '#800080','font_color': '#FFFFFF'}))

        # 設置B8到U8以下100格的格式
        for col in range(1,26):
            col_letter = chr(ord('B') + col - 1)
            worksheet.set_column(f'{col_letter}8:{col_letter}100', 15,  workbook.add_format({'num_format': '#,##0', 'align': 'right'}))
        worksheet.set_column('AA8:AA100', 11,  workbook.add_format({'num_format': '#,##0', 'align': 'right'}))
        worksheet.set_column('AB8:AB100', 11,  workbook.add_format({'num_format': '#,##0', 'align': 'right'}))
        worksheet.set_column('A1:A100',16)

print("報表完成")
