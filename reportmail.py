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
import io

load_dotenv()
logging.basicConfig(
    filename='report_mail.log',
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

def otherpeople(filepath,name):
    df = pd.read_excel(filepath)
    df.columns = df.iloc[5]
    df = df.iloc[6:].reset_index(drop=True)
    Store=[]
    Man=[]

    DScy=[]
    DSpy=[]
    DSin=[]
    DCcy=[]
    DCpy=[]
    DCin=[]
    DAcy=[]
    DApy=[]
    DAin=[]

    MScy=[]
    MSpy=[]
    MSin=[]
    MCcy=[]
    MCpy=[]
    MCin=[]
    MAcy=[]
    MApy=[]
    MAin=[]

    YScy=[]
    YSpy=[]
    YSin=[]
    YCcy=[]
    YCpy=[]
    YCin=[]
    YAcy=[]
    YApy=[]
    YAin=[]
    column=df.columns.values
    column[2]='Daily Sales CY'
    column[3]='Daily Sales PY'
    column[4]='Daily Sales Index'
    column[5]='Daily TC CY'
    column[6]='Daily TC PY'
    column[7]='Daily TC Index'
    column[8]='Daily TA CY'
    column[9]='Daily TA PY'
    column[10]='Daily TA Index'

    column[11]='MTD Sales CY'
    column[12]='MTD Sales PY'
    column[13]='MTD Sales Index'
    column[14]='MTD TC CY'
    column[15]='MTD TC PY'
    column[16]='MTD TC Index'
    column[17]='MTD TA CY'
    column[18]='MTD TA PY'
    column[19]='MTD TA Index'

    column[20]='YTD Sales CY'
    column[21]='YTD Sales PY'
    column[22]='YTD Sales Index'
    column[23]='YTD TC CY'
    column[24]='YTD TC PY'
    column[25]='YTD TC Index'
    column[26]='YTD TA CY'
    column[27]='YTD TA PY'
    column[28]='YTD TA Index'
    df.columns=column

    for i in range(len(df)):
        if i==0 or i==1 or df['營運主管'][i]==name:
            Store.append(df[' '][i])
            Man.append(df['營運主管'][i])
            DScy.append(df['Daily Sales CY'][i])
            DSpy.append(df['Daily Sales PY'][i])
            DSin.append(df['Daily Sales Index'][i])
            DCcy.append(df['Daily TC CY'][i])
            DCpy.append(df['Daily TC PY'][i])
            DCin.append(df['Daily TC Index'][i])
            DAcy.append(df['Daily TA CY'][i])
            DApy.append(df['Daily TA PY'][i])
            DAin.append(df['Daily TA Index'][i])
            MScy.append(df['MTD Sales CY'][i])
            MSpy.append(df['MTD Sales PY'][i])
            MSin.append(df['MTD Sales Index'][i])
            MCcy.append(df['MTD TC CY'][i])
            MCpy.append(df['MTD TC PY'][i])
            MCin.append(df['MTD TC Index'][i])
            MAcy.append(df['MTD TA CY'][i])
            MApy.append(df['MTD TA PY'][i])
            MAin.append(df['MTD TA Index'][i])
            YScy.append(df['YTD Sales CY'][i])
            YSpy.append(df['YTD Sales PY'][i])
            YSin.append(df['YTD Sales Index'][i])
            YCcy.append(df['YTD TC CY'][i])
            YCpy.append(df['YTD TC PY'][i])
            YCin.append(df['YTD TC Index'][i])
            YAcy.append(df['YTD TA CY'][i])
            YApy.append(df['YTD TA PY'][i])
            YAin.append(df['YTD TA Index'][i])
        
    data = {
            " ": Store,
            "營運主管":Man,
            "Daily Sales CY": DScy,
            "Daily Sales PY": DSpy,
            "Daily Sales Index":DSin,
            "Daily TC CY":DCcy,
            "Daily TC PY":DCpy,
            "Daily TC Index":DCin,
            "Daily TA CY":DAcy,
            "Daily TA PY":DApy,
            "Daily TA Index":DAin,
            "MTD Sales CY": MScy,
            "MTD Sales PY":MSpy,
            "MTD Sales Index": MSin,
            "MTD TC CY":MCcy,
            "MTD TC PY":MCpy,
            "MTD TC Index":MCin,
            "MTD TA CY":MAcy,
            "MTD TA PY":MApy,
            "MTD TA Index":MAin,
            "YTD Sales CY": YScy,
            "YTD Sales PY": YSpy,
            "YTD Sales Index": YSin,
            "YTD TC CY":YCcy,
            "YTD TC PY":YCpy,
            "YTD TC Index":YCin,
            "YTD TA CY":YAcy,
            "YTD TA PY":YApy,
            "YTD TA Index":YAin
        }
    df=pd.DataFrame(data)
    excel_stream = io.BytesIO()
    with pd.ExcelWriter(excel_stream, engine='xlsxwriter') as writer:
        MTD=TIME[6]
        YTD=TIME[7]
        df.to_excel(writer, sheet_name=sheet_names[0], index=False, startrow=6)
        workbook  = writer.book
        worksheet = writer.sheets[sheet_names[0]]
        worksheet.write('A1', TIME[9])
       
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
    return excel_stream
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
            if df[sheet_names[j]][i]=='yes' and df['備註'][i]=='ALL':
                filepath ='./Senddata/'+sheet_names[j]+'_'+ date + '_Report.xlsx'
                filename =sheet_names[j]+'_'+ date + '_Report.xlsx'
                with open(filepath, 'rb') as attachment:
                    part = MIMEApplication(attachment.read(), Name=filename)
                    part['Content-Disposition'] = f'attachment; filename="{filename}"'
                    msg.attach(part)
            elif df[sheet_names[j]][i]=='yes':
                filepath ='./Senddata/'+sheet_names[j]+'_'+ date + '_Report.xlsx'
                excel_stream=otherpeople(filepath,df['備註'][i])
                excel_stream.seek(0)
                filename =sheet_names[j]+'_'+ date + '_Report.xlsx'
                part = MIMEApplication(excel_stream.read(), Name=filename)
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

