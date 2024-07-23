import pandas as pd
import datetime

# 找前一天日期
date = str(datetime.datetime.now().year) + '-' + str(datetime.datetime.now().month) + '-' + str(int(datetime.datetime.now().day-1))
fileName ='./Senddata/'+ date + '_Report.xlsx'
lastyear = str(int(datetime.datetime.now().year)-1) + '-' + str(datetime.datetime.now().month) + '-' + str(int(datetime.datetime.now().day)-1)
print(lastyear)
# 數據資料
data = {
    "Category": ["全市场 - Total", "SS WPSA", "北一区 - Total", "SS WPSA", "台北仁爱", "台北八德", "台北内湖", "台北信义三", "大直爱买", "台北敦南二"],
    "Daily Sales CY": [6602526, 369192, 2033880, 398822, 50433, 55168, 78179, 44828, 27471, 81921],
    "Daily Sales CY (不含早餐)": [6070351, 339605, 1861736, 364574, 44251, 50665, 71951, 38082, 24771, 76361],
    "Daily Sales PY": ["########", 935570, 4528552, 897720, 100261, 85292, 191659, 81936, 103847, 174783],
    "Daily TC CY": [38191, 2143, 12031, 2376, 271, 332, 414, 288, 297, 433],
    "Daily TC CY (不含早餐)": [78273, 4274, 20999, 4172, 444, 420, 389, 444, 608, 783],
    "Daily TC PY": [75180, 4107, 20244, 4028, 444, 398, 389, 444, 608, 744],
    "Daily TA CY": [147.6, 147.2, 142.0, 140.7, 143.7, 137.6, 131.9, 111.2, 111.8, 157.8],
    "Daily TA CY (不含早餐)": [158.9, 158.5, 154.7, 153.5, 163.3, 152.6, 147.8, 128.9, 133.3, 168.7],
    "Daily TA PY": [223.2, 223.3, 215.7, 215.2, 225.8, 208.2, 287.8, 132.9, 164.1, 223.2],
    "Daily TA PY (不含早餐)": [227.8, 227.8, 219.8, 219.1, 227.8, 214.3, 294.5, 133.9, 163.1, 229.0]
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
