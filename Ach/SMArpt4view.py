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
print 'Cycle start date is ' + str(startday)
print 'Cycle end date is ' + str(endday)

#---------------------------------------
driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
db_file = r"F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA.accdb;"
user = "admin"
password = ""
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)

sql = '''
SELECT DISTINCT
	tbl_SMA.[SMAKey]
	,tbl_SMA.[CycleDate]
	,tbl_SMA.[TransType]
	,tbl_SMA.[Event Effective Date]
	,tbl_SMA.[Client Servicing Consultant Number] AS [Cslt]
	,tbl_SMA.[Client Number]
	,tbl_SMA.[Client Last Name]
	,tbl_SMA.[Client Given Name]
	,tbl_SMA.[Account Number]
	,tbl_SMA.[Event Gross Amount]
	,tbl_SMA.[Account Market Value]
	,tbl_SMA.[Product IGSI Symbol]
	,tbl_SMA.[Event Activity Description]
	,tbl_SMA.[Product Description]
	,tbl_SMA.[Notes]
FROM tbl_SMA
WHERE ((tbl_SMA.CycDate) = # ''' + endday.strftime("%Y/%m/%d") + ''' #)
ORDER BY 
	tbl_SMA.[CycleDate]
	,tbl_SMA.[Client Servicing Consultant Number]
	,tbl_SMA.[Client Number]
	,tbl_SMA.[Account Number];
'''

conn = pyodbc.connect(odbc_conn_str)
cur = conn.cursor()
SMAdf = pd.read_sql_query(sql,conn)
conn.close()
#print SMAdf.head()
#print SMAdf.dtypes
#sys.exit("done")

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('SMADaily' + endday.strftime("%Y%m%d") + '.xlsx', engine='xlsxwriter')
#writer = pd.ExcelWriter('F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMADaily' + endday.strftime("%Y%m%d") + '.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
SMAdf.to_excel(writer, sheet_name='Sheet1', index=False)

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

# Add some cell formats.
format1 = workbook.add_format({'bold': True, 'bg_color': '#FFFF00'})
format2 = workbook.add_format({'num_format': '#,##0.00', 'bold': True, 'bg_color': '#FFC7CE'})
format3 = workbook.add_format({'num_format': '#,##0.00'})

# Set the column width and format.
worksheet.set_column('E:E', 15, format1) #[Cslt]
worksheet.set_column('J:J', 18, format2) #[Event Gross Amount]
worksheet.set_column('K:K', 18, format3) #[Account Market Value]

# Close the Pandas Excel writer and output the Excel file.
writer.save()

print 'output all daily transaction to ' + 'SMADaily' + endday.strftime("%Y%m%d") + '.xlsx'
