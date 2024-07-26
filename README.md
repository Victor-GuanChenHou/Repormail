## Reportmail
[TOC]

![upload_9304c5832db2853a9c168e97f079226d](https://github.com/user-attachments/assets/d914c659-1ec3-4e8a-aed7-e8dadf9b4158)

### Introduction:
Use the POS server to generate sales data for the day every day, and then use SFTP to send it to Nas。Report Server uses Python SFTP to capture data on Nas every day and store it in folders. Then uses Python to process the data and store it in MySQL. Then uses Python to retrieve the data needed for the report and make it into an .xlsx. Finally, it uses Python SMTP to send mail.

### Installation
#### MySQL(8.0.37):
- https://downloads.mysql.com/archives/community/
#### Python(3.12.3):
- https://www.python.org/downloads/release/python-3123/
- Pandas:2.1.4
- PyMySQL:1.0.2
- python-dotenv:1.0.1
#### Creat .env File
``` 
Nashost='example'
Nasusername='example'
Naspassword='example'
mailusername='example'
mailpassword='example'
SQLhost='example'
SQLuser='example'
SQLpassword='example'
database='example'
port='example'
```
### Program Introduction & Notice
#### SFTPgetdata
- Nas connect:
![image](https://hackmd.io/_uploads/S1yM9igFC.png)
#### Datamerge
- File reading path
![image](https://hackmd.io/_uploads/rJp73sxFA.png)
#### ExcelCreater
- File storage path
![image](https://hackmd.io/_uploads/rynh2sgYR.png)
#### Reportmail
- Mail Sender credential
![image](https://hackmd.io/_uploads/H1AXpogt0.png)
- SMTP server
![image](https://hackmd.io/_uploads/ByhT6ilKC.png)
- Receiver
![image](https://hackmd.io/_uploads/r1_bRieKA.png)
#### SQL_tra
- SQL logging 
![image](https://hackmd.io/_uploads/HyF_CieKC.png)



