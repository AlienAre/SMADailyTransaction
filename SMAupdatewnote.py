#--------------------------------------------
#version:		1.0
#author:		West
#Description:	used to get SMA daily transaction for cycle requested
#				run by each cycle
#Workflow:		get SMA daily transaction for cycle requested
#				get AL, rate information and calculate new business, sales bonus, sales bonus AL adj
#				export to Excel
#
#--------------------------------------------

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
	print('Process date is ' + str(time.strftime("%m/%d/%Y")))

	getcycledate = datetime.datetime.strptime(input('Please enter the cycle end date (mm/dd/yyyy) you want to update:'), '%m/%d/%Y')
	endday = getcycledate
	startday = ig.getCStartDate(getcycledate)

	print('Cycle start date is ' + str(startday))
	print('Cycle end date is ' + str(endday))

	#---------- get the updated spreadsheet and read --------
	dfsmadata = pd.read_excel('F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMADaily' + endday.strftime("%Y%m%d") + '.xlsx')
	#print dfsmadata.head()

	#----------------- Update TransType and Notes ----------------------
	driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
	db_file = r"F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA.accdb;"
	#user = "admin"
	#password = ""
	#odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)

	table = "tbl_SMA"

	for index, row in dfsmadata.iterrows():
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
	#	#sys.exit("done")
	#	conn = pyodbc.connect(odbc_conn_str)
	#	cursor = conn.cursor()
	#	cursor.execute(sql)
	#	cursor.commit()
	#	cursor.close()
	#conn.close()
	dbq.update_tbldate(driver, db_file, sql)

	print('database has been updated successfully')
