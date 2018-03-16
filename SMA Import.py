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

## dd/mm/yyyy format
print 'Process date is ' + str(time.strftime("%d/%m/%Y"))

startday = dd.getCycleStartDate(date.today())
endday = dd.getCycleEndDate(date.today())
#startday = dd.getCycleStartDate(datetime.datetime.strptime(str('2017/11/8'), '%Y/%m/%d'))
#endday = dd.getCycleEndDate(datetime.datetime.strptime(str('2017/11/8'), '%Y/%m/%d'))
print 'Cycle start date is ' + str(startday)
print 'Cycle end date is ' + str(endday)

filesdir = 'F:\\3-Compensation Programs\\IIROC Compensation\\' + endday.strftime("%Y%m%d")

labels = ['Event Record Type', 'Event Effective Date', 'Event Process Date', 'Event Activity Type', 'Event Activity Description', 'Event Gross Amount', 'Plan Product Code', 'Account Market Value', 'Client Number', 'Client Last Name', 'Client Given Name', 'Client Servicing Consultant Number', 'Client Deceased Indicator', 'Client Company Name', 'Client Province Code', 'Account Number', 'Account Dealer Code', 'Account IGSI Net Share Quantity', 'Product Code', 'Product Share Price Amount', 'Product IGSI Symbol', 'Product Description', 'Product Security Type', 'Product Security Class', 'Product Security Category']
#cashlabels = ['Event Record Type', 'Event Effective Date', 'Event Process Date', 'Event Activity Type', 'Event Activity Description', 'Event Gross Amount', 'Plan Product Code', 'Account Market Value', 'Client Number', 'Client Last Name', 'Client Given Name', 'Client Servicing Consultant Number', 'Client Deceased Indicator', 'Client Company Name', 'Client Province Code', 'Account Number', 'Account Dealer Code', 'Account IGSI Net Share Quantity', 'Product Code', 'Product Share Price Amount', 'Product IGSI Symbol', 'Product Description', 'Product Security Type', 'Product Security Class', 'Product Security Category']
SMAdata = pd.DataFrame()	#set blank data frame for SMA daily use
#SMAcashdata = pd.DataFrame()	#set blank data frame for SMA cash daily use

SMAlist = []
#SMAcashlist = []

pattern = '*SMA.EVENTS*.xls'	#use to find SMA daily files
#cashpattern = '*SMA.CASH.EVENTS*.xls'	#use to find SMA cash daily files

### go to dir and get all SMA.EVENTS excel list
files = os.listdir(filesdir)
for file in fnmatch.filter(files, pattern):
		SMAlist.append(os.path.join(filesdir, file))

### iterate all SMA.EVENTS excel files and extract data to df
for sma in SMAlist:
	df = pd.read_excel(sma, header=None)
	df1 = (df.loc[df[0] == 'D'])

	if not df1.empty:
		SMAdata = SMAdata.append(df1, ignore_index=True)
	
SMAdata.columns = labels
SMAdata['Client Company Name'] = SMAdata['Client Company Name'].astype(str).str.replace("'", "")

#---------------------------------------
driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
db_file = r"F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA.accdb;"
user = "admin"
password = ""
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)

table = "tbl_SMA"
columns = '''
[CycDate],
[CycleDate],
[TransType],
[Event Record Type],
[Event Effective Date],
[Event Process Date],
[Event Activity Type],
[Event Activity Description],
[Event Gross Amount],
[Plan Product Code],
[Account Market Value],
[Client Number],
[Client Last Name],
[Client Given Name],
[Client Servicing Consultant Number],
[Client Deceased Indicator],
[Client Company Name],
[Client Province Code],
[Account Number],
[Account Dealer Code],
[Account IGSI Net Share Quantity],
[Product Code],
[Product Share Price Amount],
[Product IGSI Symbol],
[Product Description],
[Product Security Type],
[Product Security Class],
[Product Security Category]
'''

for row in SMAdata.to_records(index=False):
	values = ", ".join(['\'%s\'' % x for x in row])
	values = values.replace("'nan'", "NULL")
	#print values
	sql = '''INSERT INTO %s (%s) VALUES ( # ''' + endday.strftime("%Y/%m/%d") + ''' #, ' ''' + endday.strftime("%Y%m%d") + ''' ', 1, %s);'''
	sql = sql % (table, columns, values)
	#print sql 
	#sys.exit("done")
	conn = pyodbc.connect(odbc_conn_str)
	cursor = conn.cursor()
	cursor.execute(sql)
	cursor.commit()
	cursor.close()
conn.close()
