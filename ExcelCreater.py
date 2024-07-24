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
def datamerge(CYdata,PYdata):
    data=[]
    CY_unique=[]
    for i in range(len(CYdata)):
        Status=False
        for j in range(len(PYdata)):
            if CYdata[i][0]==PYdata[j][0]:
                # id name cysales pysales index TC TC_P TC_index TA TA_P TA_index
                thedata=(CYdata[i][0],CYdata[i][1],float(CYdata[i][3]),float(PYdata[j][3]),float(CYdata[i][3])/float(PYdata[j][3]),float(CYdata[i][4]),float(PYdata[j][4]),float(CYdata[i][4])/float(PYdata[j][4]),float(CYdata[i][3])/float(CYdata[i][4]),float(PYdata[j][3])/float(PYdata[j][4]),(float(CYdata[i][3])/float(CYdata[i][4]))/(float(PYdata[j][3])/float(PYdata[j][4])))
                data.append(thedata)
                Status=True
                if PYdata[j] in PYdata:
                    PYdata.remove(PYdata[j]) 
                break
        if Status==False:
            CY_unique.append(CYdata[i])
    try:
        for i in range(len(CY_unique)):
            thedata=(CY_unique[i][0],CY_unique[i][1],float(CY_unique[i][3]),None,None,float(CY_unique[i][4]),None,None,float(CY_unique[i][3])/float(CY_unique[i][4]),None,None)
            data.append(thedata)
    except:
        pass
    try:
        for i in range(len(PYdata)):
            thedata=(PYdata[i][0],PYdata[i][1],None,float(PYdata[j][3]),None,None,float(PYdata[j][4]),None,None,float(PYdata[j][3])/float(PYdata[j][4]),None)
            data.append(thedata)
    except:
        pass
    data=pd.DataFrame(data)
    return data
brand=['杏']
CYdata=MYSQL.searchdata(brand[0],lastdate,lastdate)
PYdata=MYSQL.searchdata(brand[0],Plastdate,Plastdate)
CYMTDdata=MYSQL.searchdata(brand[0],Mtd,lastdate)
PYMTDdata=MYSQL.searchdata(brand[0],PMtd,Plastdate)
CYYTDdata=MYSQL.searchdata(brand[0],Ytd,lastdate)
PYYTDdata=MYSQL.searchdata(brand[0],PYtd,Plastdate)
# 數據資料


data = {
    "Category": CYdata[0],
    "Daily Sales CY": CYdata[3],
    "Daily Sales PY": PYdata[3],
    "Daily Sales Index":
    "Daily TC CY":
    "Daily TC PY":
    "Daily TC Index":
    "Daily TA CY":
    "Daily TA PY":
    "Daily TA Index":
    "MTD Sales CY": CYMTDdata[3],
    "MTD Sales PY":PYMTDdata[3],
    "MTD Sales Index": 
    "MTD TC CY":
    "MTD TC PY":
    "MTD TC Index":
    "MTD TA CY":
    "MTD TA PY":
    "MTD TA Index":
    "YTD Sales CY": CYYTDdata[3],
    "YTD Sales PY": PYYTDdata[3],
    "YTD Sales Index": 
    "YTD TC CY":
    "YTD TC PY":
    "YTD TC Index":
    "YTD TA CY":
    "YTD TA PY":
    "YTD TA Index":
}

# 創建dataframe
df = pd.DataFrame(data)


with pd.ExcelWriter(fileName, engine='xlsxwriter') as writer:
    sheet_names = ['杏子豬排', '大阪王將', '京都勝牛', '段純貞', '橋村']
    MTD= '('+str(datetime.datetime.now().month) + '/1~'+ str(datetime.datetime.now().month)+'/'+ str(int(datetime.datetime.now().day)-1)+')'
    YTD='(1/1~'+ str(datetime.datetime.now().month)+'/'+ str(int(datetime.datetime.now().day)-1)+')'
    print(MTD)
    print(YTD)
    for sheet_name in sheet_names:
        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]
        
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
