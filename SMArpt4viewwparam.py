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


#--------------------------------------------------
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
	print("The following will get cycle data from SMA database and export to Excel")

	getcycledate = datetime.datetime.strptime(input('Please enter the cycle end date (mm/dd/yyyy) you want to get:'), '%m/%d/%Y')
	endday = getcycledate
	startday = ig.getCStartDate(getcycledate)

	print('Cycle start date is ' + str(startday))
	print('Cycle end date is ' + str(endday))

	#----------- get SMA daily transactions and AL information ------------
	driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
	db_file = r"F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA.accdb;"
	#user = "admin"
	#password = ""
	#odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)
	#conn = pyodbc.connect(odbc_conn_str)

	sql = '''
			SELECT DISTINCT
				qry_SMATranswALAll.[SMAKey]
				,qry_SMATranswALAll.[CycDate]
				,qry_SMATranswALAll.[TransType]
				,qry_SMATranswALAll.[Event Effective Date]
				,qry_SMATranswALAll.[Cslt]
				,qry_SMATranswALAll.[EarnedAL]
				,qry_SMATranswALAll.[AdvanceAL]
				,qry_SMATranswALAll.[Tenure]
				,qry_SMATranswALAll.[Client Number]
				,qry_SMATranswALAll.[Client Last Name]
				,qry_SMATranswALAll.[Client Given Name]
				,qry_SMATranswALAll.[Account Number]
				,qry_SMATranswALAll.[Event Gross Amount]
				,qry_SMATranswALAll.[Account Market Value]
				,qry_SMATranswALAll.[Product IGSI Symbol]
				,qry_SMATranswALAll.[Event Activity Description]
				,qry_SMATranswALAll.[Product Description]
				,qry_SMATranswALAll.[Notes]
			FROM qry_SMATranswALAll
			WHERE ((qry_SMATranswALAll.CycDate) = # ''' + str(endday) + ''' #)
			ORDER BY 
				qry_SMATranswALAll.[CycDate]
				,qry_SMATranswALAll.[Cslt]
				,qry_SMATranswALAll.[Client Number]
				,qry_SMATranswALAll.[Account Number];
		'''

	#dftrans = pd.read_sql_query(sql,conn)
	#conn.close()
	dftrans = dbq.df_select(driver, db_file, sql)
	#-----------------------------------------------------------------------

	#----------- get Sales Bonus rate ------------
	driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
	db_file = r"F:\Files For\West Wang\Rates.accdb;"
	#user = "admin"
	#password = ""
	#odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)
	#conn = pyodbc.connect(odbc_conn_str)

	#--------- get New Business rate based on AL for advancing AL ----------
	sql = '''SELECT DISTINCT NewBusinessRate.Rate AS NBRate FROM NewBusinessRate WHERE NewBusinessRate.NBYear = ''' + str(endday.year)
	#NBRatedf = pd.read_sql_query(sql,conn)
	#dftrans['NBRate'] = NBRatedf.at[(0, 'NBRate')]
	dfnbrate = dbq.df_select(driver, db_file, sql)
	dftrans["NBRate"] = dfnbrate.iloc[0,0]
	#-----------------------------------------------------------------------

	#--------- get sales bonus rate for cslt under year 4 ----------
	sql = '''SELECT FSalesBonusRate.Level AS Tenure, FSalesBonusRate.Rate FROM FSalesBonusRate''' 
	#Ratedf = pd.read_sql_query(sql,conn)
	#dftrans = dftrans.merge(Ratedf, left_on='Tenure', right_on='Tenure', how='left')
	dfsbrate = dbq.df_select(driver, db_file, sql)
	dftrans = dftrans.merge(dfsbrate, how="left", on="Tenure")

	#--------- get sales bonus rate based on AL ----------
	sql = '''SELECT DISTINCT FTransitionalSalesBonusRate.Level AS EarnedAL, FTransitionalSalesBonusRate.Rate AS ALRate FROM FTransitionalSalesBonusRate WHERE FTransitionalSalesBonusRate.ALYear = ''' + str(endday.year)
	#ALratedf = pd.read_sql_query(sql,conn)
	#dftrans =  dftrans.merge(ALratedf, left_on='EarnedAL', right_on='EarnedAL', how='left')
	dfalrate = dbq.df_select(driver, db_file, sql)
	dftrans = dftrans.merge(dfalrate, how="left", on="EarnedAL")
	dftrans.loc[dftrans['Tenure'] < 4, 'SBRate'] = dftrans['Rate']
	dftrans.loc[dftrans['Tenure'] > 3, 'SBRate'] = dftrans['ALRate']

	#--------- get sales bonus rate based on AL for advancing AL ----------
	sql = '''SELECT DISTINCT FTransitionalSalesBonusRate.Level AS AdvanceAL, FTransitionalSalesBonusRate.Rate AS AdvRate FROM FTransitionalSalesBonusRate WHERE FTransitionalSalesBonusRate.ALYear = ''' + str(endday.year)
	#Advratedf = pd.read_sql_query(sql,conn)
	#dftrans =  dftrans.merge(Advratedf, left_on='AdvanceAL', right_on='AdvanceAL', how='left')
	dfadvalrate = dbq.df_select(driver, db_file, sql)
	dftrans = dftrans.merge(dfadvalrate, how="left", on="AdvanceAL")
	dftrans["AdvRate"].fillna(0, inplace = True)
	#conn.close()
	#-----------------------------------------------------------------------

	#----------- calculate sales bonus/Advancing AL sales bonus/New Business  ------------
	dftrans['New Business'] = dftrans['Event Gross Amount'] * dftrans['NBRate'] * dftrans['TransType']
	#dftrans['Sales Bonus'] = np.where(dftrans['TransType'] == 1, dftrans['Event Gross Amount'] * dftrans['SBRate'], 0.00)
	dftrans['Sales Bonus'] = dftrans['Event Gross Amount'] * dftrans['SBRate'] * dftrans['TransType']
	dftrans['AL Advancing Adj'] = np.where(dftrans['AdvRate'] != 0, (dftrans['Event Gross Amount'] * dftrans['AdvRate'] - dftrans['Sales Bonus']) * dftrans['TransType'], 0.00)
	#dftrans['Mark'] = np.where(dftrans['TransType'] == 0, '*', '')

	#----------- Remove unrequired columns ------------
	dftrans.drop(['Rate', 'ALRate'], axis=1, inplace=True)
	#dftrans.rename(columns={'Event Gross Amount': 'Total Contribution'}, inplace=True)
	dftrans = dftrans[['SMAKey','CycDate','TransType','Event Effective Date','Cslt','Client Number','Client Last Name','Client Given Name','Account Number','Event Gross Amount','Account Market Value','Tenure','EarnedAL','NBRate','New Business','SBRate','Sales Bonus','AdvanceAL','AdvRate','AL Advancing Adj','Event Activity Description','Product IGSI Symbol','Product Description','Notes']]
	dftrans.sort_values(['CycDate', 'Cslt', 'Client Number', 'Account Number', 'Event Gross Amount'], inplace=True)

	#print dftrans.dtypes
	#sys.exit("done")

	# Create a Pandas Excel writer using XlsxWriter as the engine.
	writer = pd.ExcelWriter('SMADaily' + endday.strftime("%Y%m%d") + '.xlsx', engine='xlsxwriter')
	#writer = pd.ExcelWriter('F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMADaily' + endday.strftime("%Y%m%d") + '.xlsx', engine='xlsxwriter')

	# Convert the dataframe to an XlsxWriter Excel object.
	dftrans.to_excel(writer, sheet_name='Sheet1', index=False)

	# Get the xlsxwriter workbook and worksheet objects.
	workbook  = writer.book
	worksheet = writer.sheets['Sheet1']

	# Add some cell formats.
	formatcslt = workbook.add_format({'bold': True, 'bg_color': '#FFFF00'})
	formatnum = workbook.add_format({'num_format': '#,##0.00', 'bold': True, 'bg_color': '#FFC7CE'})
	formatrate = workbook.add_format({'num_format': '0.00%'})
	format3 = workbook.add_format({'num_format': '#,##0.00'})

	# Set the column width and format.
	worksheet.set_column('E:E', 10, formatcslt) #[Cslt]
	worksheet.set_column('I:I', 15) #[Account Number]
	worksheet.set_column('J:J', 18, formatnum) #[Event Gross Amount]
	worksheet.set_column('K:K', 18, format3) #[Account Market Value]
	worksheet.set_column('N:N', 10, formatrate) #[NBRate]
	worksheet.set_column('O:O', 12, formatnum) #[New Business]
	worksheet.set_column('P:P', 10, formatrate) #[SBRate]
	worksheet.set_column('Q:Q', 12, formatnum) #[Sales Bonus]

	# Close the Pandas Excel writer and output the Excel file.
	writer.save()

	print('output all daily transaction to ' + 'SMADaily' + endday.strftime("%Y%m%d") + '.xlsx')

