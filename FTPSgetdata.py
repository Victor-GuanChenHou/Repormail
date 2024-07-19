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

# file_list = sftp.listdir("./ALL")
# print("文件列表:", file_list)
sftp.chdir("All")
folder_file_list = sftp.listdir()
print(folder_file_list[0][0:8])

remote_filename = folder_file_list[0]
local_filename = str(folder_file_list[0][0:8])+'.csv'
sftp.get(remote_filename, './Origianldata/'+str(local_filename))

sftp.close()
ssh.close()