## Reportmail
* [Introduction](#introduction)
* [Installation](#installation)
  * [MySQL(8.0.37)](#mysql8037)
  * [Python(3.12.3)](#python3123)
  * [Creat .env File](#creat-env-file)
* [Program Introduction & Notice](#program-introduction--notice)
  * [SFTPgetdata](#sftpgetdata)
  * [Datamerge](#datamerge)
  * [ExcelCreater](#excelcreater)
  * [Reportmail](#reportmail)
  * [SQL_tra](#sql_tra)
  
  
    
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

![upload_5f575d7cf36c24e72051b26ea0d062b6](https://github.com/user-attachments/assets/5f6d17d1-a7fb-4b8e-973c-0bc331c347d6)

#### Datamerge
- File reading path

![image](https://github.com/user-attachments/assets/e62d0f4a-f629-463f-9343-5548e7dc0568)


#### ExcelCreater
- File storage path

![upload_6fa45b518707a75aab1e1264d64990b2](https://github.com/user-attachments/assets/b67ad1c6-cbe2-4dbd-8338-4003e8abf818)

#### Reportmail
- Mail Sender credential

![螢幕擷取畫面 2024-07-26 125840](https://github.com/user-attachments/assets/dc3ffc81-4c1a-4758-8aad-d9c3def79095)

- SMTP server

![螢幕擷取畫面 2024-07-26 130148](https://github.com/user-attachments/assets/bc578691-7585-4909-b132-90bc5d6b1508)

- Receive
  
![螢幕擷取畫面 2024-07-26 130343](https://github.com/user-attachments/assets/a970041c-178d-4bdc-ae96-3f45a9fcd614)

#### SQL_tra
- SQL logging

![螢幕擷取畫面 2024-07-26 130408](https://github.com/user-attachments/assets/e1fc54fb-d140-4c67-8487-b1ff9686e7b9)




