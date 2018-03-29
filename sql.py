import os
import re
import sys
import time
from datetime import date
import fnmatch
import pandas as pd
import itertools as it
from openpyxl import load_workbook
import xlrd
import myfun as dd
import pyodbc
import datetime

#----------- get SMA daily transactions and AL information ------------
driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
db_file = r"F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA.accdb;"
user = "admin"
password = ""
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)
conn = pyodbc.connect(odbc_conn_str)
#--------------------------------------------------------------------

sql = '''
	SELECT tbl_SMA.[SMAKey], tbl_SMA.[Event Effective Date], tbl_SMA.[Event Process Date]
	FROM tbl_SMA
	WHERE tbl_SMA.[CycDate] = #12/15/2016#;
	'''
	
df = pd.read_sql_query(sql,conn)

#df['Event Process Date'] = 
df['New'] = df['Event Effective Date'].astype(str).str[0:4] + '/' + df['Event Effective Date'].astype(str).str[4:6] + '/' + df['Event Effective Date'].astype(str).str[6:8]  

print df
for index, row in df.iterrows():
	sql = '''
	UPDATE tbl_SMA
	SET [Event Process Date] = '%s'
	WHERE [SMAKey] = %s '''
	sql = sql % (row['New'], row['SMAKey'])
	conn = pyodbc.connect(odbc_conn_str)
	cursor = conn.cursor()
	cursor.execute(sql)
	cursor.commit()
	cursor.close()
conn.close()