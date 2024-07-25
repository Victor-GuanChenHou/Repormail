import os
from dotenv import load_dotenv
load_dotenv()
import pymysql
import datetime
host=os.getenv('SQLhost')
user=os.getenv('SQLuser')
password=os.getenv('SQLpassword')
database=os.getenv('database')
port=int(os.getenv('port'))
def insert_sales_data(data):
    global host
    global user
    global password
    pymysql.install_as_MySQLdb()
    db=pymysql.connect(host=host,user=user,password=password,database=database,port=port)
    cursor = db.cursor()
    val=[]
    for i in range(len(data)):
         val.append((str(data[i]['store_id']),int(data[i]['total_customer']), str(data[i]['invoice_amt']), str(data[i]['DATE'])))
    sql = "INSERT INTO salesdata (store_id, total_customer,invoice_amt, DATE) VALUES (%s, %s,%s, %s)"
    cursor.executemany(sql, val)
    db.commit()
    db.close()
def insert_store_data(data):
    global host
    global user
    global password
    pymysql.install_as_MySQLdb()
    db=pymysql.connect(host=host,user=user,password=password,database=database,port=port)
    cursor = db.cursor()
    for i in range(len(data)):
        try:
            val=(str(data[i]['store_id']), str(data[i]['brand']), str(data[i]['store_name']))
            sql = "INSERT INTO storedata (store_id, brand, store_name) VALUES (%s, %s, %s)"
            cursor.execute(sql, val)
            db.commit()
        except:
             pass
    db.close()
    
def searchdata(brand,start_time,end_time):
    global host
    global user
    global password
    pymysql.install_as_MySQLdb()
    try:
        db=pymysql.connect(host=host,user=user,password=password,database=database,port=port)
        cursor = db.cursor()
        cursor.execute("SELECT store_id FROM storedata WHERE brand = %s",brand)
        storeids = cursor.fetchall()
        data=[]
        for storeid_tuple in storeids:
                storeid = storeid_tuple[0]

                # 查詢每個storeid在指定日期範圍內的銷售數據
                query = """
                SELECT 
                    sd.store_id, 
                    sd.store_name, 
                    sa.DATE, 
                    SUM(sa.sale_amount) AS total_sales,
                    SUM(sa.total_customer) AS total_customers
                FROM 
                    storedata sd
                JOIN 
                    salesdata sa
                ON 
                    sd.store_id = sa.store_id
                WHERE 
                    sd.store_id = %s
                    AND sa.DATE BETWEEN %s AND %s;
                """
                cursor.execute(query, (storeid, start_time, end_time))
                result = cursor.fetchall()

                # 處理結果，例如打印或者存儲到文件
                for row in result:
                    data.append(row)
        return data
    finally:
        db.close()

