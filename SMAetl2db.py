import os, re, sys, time, xlrd, pyodbc, datetime
from datetime import date
import fnmatch
import numpy as np
import pandas as pd
import itertools as it
from openpyxl import load_workbook
from shutil import copyfile

sys.path.append('C:\\pycode\\libs')
import igtools as ig
import dbquery as dbq

#------ program starting point --------
if __name__=="__main__":
	## dd/mm/yyyy format
	print("Process date is " + str(time.strftime("%m/%d/%Y")))
	print("The following will get cycle data and load to SMA database.")

	#------- get cycle date ----------------------
	getcycledate = datetime.datetime.strptime(input("Please enter the cycle end date (mm/dd/yyyy) you want to process:"), '%m/%d/%Y')
	endday = getcycledate
	startday = ig.getStartDate(getcycledate)

	print('Cycle start date is ' + str(startday))
	print('Cycle end date is ' + str(endday))
	#---------------------------------------------

	#----------- backup db first ---------
	copyfile('F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA.accdb', 'M:\\bak\\SMA' + str(date.today().strftime("%m%d%Y")) + '.accdb')
	#-------------------------------------

	#----------- get SMA daily transactions and AL information ------------
	driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
	db_file = r"F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA.accdb;"
	user = "admin"
	password = ""
	#odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)
	#conn = pyodbc.connect(odbc_conn_str)
	#--------------------------------------------------------------------
	#--------- check latest cycle date in database ----------
	sql = '''SELECT Max(tbl_SMA.CycDate) AS [CycDate] FROM tbl_SMA; '''
	#Cdatedf = pd.read_sql_query(sql,conn)
	dfcdate = dbq.df_select(driver, df_file, sql)
	latestcycledate = Cdatedf.at[(0, 'CycDate')]
	#print type(latestcycledate)
	print('In database, the latest cycle end date is ' + str(latestcycledate))
	if latestcycledate > endday:
		(print 'It seems that the cycle date in database is later than your cycle date, which means the transactions may already be entered into database. Please type "1", if you want to proceed:')
		if input() != '1':
			sys.exit("The process is stopped")
	#-----------------------------------------------------------------------

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

	#---- if there is no transaction for the cycle, the process should be stopped
	if SMAdata.empty:
		sys.exit("This is no data for this cycle. Process is stopped")

	SMAdata.columns = labels
	SMAdata['Client Company Name'] = SMAdata['Client Company Name'].astype(str).str.replace("'", "")
	SMAdata['Event Process Date'] = SMAdata['Event Process Date'].astype(str).str.replace(" 00:00:00", "")
	#print SMAdata.head()
	#---------------------------------------
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

#	for row in SMAdata.to_records(index=False):
#		values = ", ".join(['\'%s\'' % x for x in row])
#		values = values.replace("'nan'", "NULL")
#		#print values
#		sql = '''INSERT INTO %s (%s) VALUES ( # ''' + endday.strftime("%Y/%m/%d") + ''' #, ' ''' + endday.strftime("%Y%m%d") + ''' ', 1, %s);'''
#		sql = sql % (table, columns, values)
#		#print sql
#		#sys.exit("done")
#		conn = pyodbc.connect(odbc_conn_str)
#		cursor = conn.cursor()
#		cursor.execute(sql)
#		cursor.commit()
#		cursor.close()
#	conn.close()
	dbq.add_to_tbl(driver, db_file, table, columns, SMAdata)

	print('The process is completed successfully')
