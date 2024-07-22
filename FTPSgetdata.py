import paramiko
from dotenv import load_dotenv
import os
import sys
import logging
load_dotenv()
logging.basicConfig(
    filename='sftp.log',
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# SFTP 連線資料
try:
    hostname = os.getenv('Nashost')
    username = os.getenv('Nasusername')
    password = os.getenv('Naspassword')
    port = 22 

    # 創建SSH客戶端
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(hostname, port, username, password)
    logging.info("Connect Success")
except:
    logging.info("Connect Error")
    sys.exit()
# 創建SFTP客戶端
sftp = ssh.open_sftp()
sftp.chdir("All")#在Nas All的資料夾拿資料
folder_file_list = sftp.listdir()

#拿取前一天資料
remote_filename = folder_file_list[len(folder_file_list)-1]
local_filename = str(folder_file_list[len(folder_file_list)-1][0:8])+'.csv'#將儲存的資料轉成前一天日期.csv
sftp.get(remote_filename, './Origianldata/'+str(local_filename))
#拿Nas取所有資料
# for i in range(len(folder_file_list)):
#     remote_filename = folder_file_list[i]
#     local_filename = str(folder_file_list[i][0:8])+'.csv'#將儲存的資料轉成前一天日期.csv
#     sftp.get(remote_filename, './Origianldata/'+str(local_filename))
sftp.close()
ssh.close()
