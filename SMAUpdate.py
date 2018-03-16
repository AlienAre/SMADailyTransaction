import os, re, sys, time, xlrd
from datetime import date
import fnmatch
import pandas as pd
import itertools as it
from openpyxl import load_workbook
import myfun as dd
import pyodbc
import datetime

## dd/mm/yyyy format
print 'Process date is ' + str(time.strftime("%d/%m/%Y"))

startday = dd.getCycleStartDate(date.today())
endday = dd.getCycleEndDate(date.today())
#startday = dd.getCycleStartDate(datetime.datetime.strptime(str('2017/11/8'), '%Y/%m/%d'))
#endday = dd.getCycleEndDate(datetime.datetime.strptime(str('2017/11/8'), '%Y/%m/%d'))
print 'Cycle start date is ' + str(startday)
print 'Cycle end date is ' + str(endday)


SMAdata = pd.read_excel('F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMADaily' + endday.strftime("%Y%m%d") + '.xlsx')
#print SMAdata.head()

#---------------------------------------
driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
db_file = r"F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA.accdb;"
user = "admin"
password = ""
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)

table = "tbl_SMA"

for index, row in SMAdata.iterrows():
	sql = '''
	UPDATE %s
	SET [TransType] = %s, [Notes] = '%s'
	WHERE [SMAKey] = %s '''
	sql = sql % (table, row['TransType'], row['Notes'], row['SMAKey'])
#-------------------------------------	
#	print sql
#	with open("Output.txt", "w") as text_file:
#		text_file.write(sql)
#	sys.exit("done")	
#--------------------------------------	
	#sys.exit("done")
	conn = pyodbc.connect(odbc_conn_str)
	cursor = conn.cursor()
	cursor.execute(sql)
	cursor.commit()
	cursor.close()
conn.close()
