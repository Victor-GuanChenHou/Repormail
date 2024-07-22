import os
from dotenv import load_dotenv
load_dotenv()
import pymysql
host=os.getenv('SQLhost')
user=os.getenv('SQLuser')
password=os.getenv('SQLpassword')
database=os.getenv('database')
port=os.getenv('port')
def insert_sales_data(data):
    global host
    global user
    global password
    pymysql.install_as_MySQLdb()
    db=pymysql.connect(host=host,user=user,password=password,database=database,port=port)
    cursor = db.cursor()
    val=[]
    for i in range(len(data)):
         val.append((str(data[i]['store_id']), str(data[i]['invoice_amt']), str(data[i]['DATE'])))
    sql = "INSERT INTO salesdata (store_id, invoice_amt, DATE) VALUES (%s, %s, %s)"
    cursor.executemany(sql, val)
def insert_store_data(data):
    global host
    global user
    global password
    pymysql.install_as_MySQLdb()
    db=pymysql.connect(host=host,user=user,password=password,database=database,port=port)
    cursor = db.cursor()
    val=[]
    for i in range(len(data)):
         val.append((str(data[i]['store_id']), str(data[i]['brand']), str(data[i]['store_name'])))
    sql = "INSERT INTO storedata (store_id, brand, store_name) VALUES (%s, %s, %s)"
    cursor.executemany(sql, val)
    
    
