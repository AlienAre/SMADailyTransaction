#--------------------------------------------
#version:		1.0
#author:		West
#Description:	prepare final SMA cycle report and Sale Bonus/New Business upload file
#				suppose to run every cycle after manager reviews the transaction and db is updated
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

sys.path.append('C:\\pycode\\libs')
import igtools as ig
import dbquery as dbq

#------ program starting point --------	
if __name__=="__main__":		
	## dd/mm/yyyy format
	print('Process date is ' + str(time.strftime("%m/%d/%Y")))

	getcycledate = datetime.datetime.strptime(input('Please enter the cycle end date (mm/dd/yyyy) you want to output:'), '%m/%d/%Y')
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
				qry_SMATranswALAll.[CycDate]
				,qry_SMATranswALAll.[Cslt]
				,qry_SMATranswALAll.[Name]
				,qry_SMATranswALAll.[RO]
				,qry_SMATranswALAll.[ROName]
				,qry_SMATranswALAll.[TransType]
				,qry_SMATranswALAll.[Account Number]
				,qry_SMATranswALAll.[Event Process Date]
				,qry_SMATranswALAll.[Client Number]
				,qry_SMATranswALAll.[Client Last Name]
				,qry_SMATranswALAll.[Client Given Name]
				,qry_SMATranswALAll.[Event Gross Amount]	
				,qry_SMATranswALAll.[EarnedAL]
				,qry_SMATranswALAll.[AdvanceAL]
				,qry_SMATranswALAll.[Tenure]
			FROM qry_SMATranswALAll
			WHERE ((qry_SMATranswALAll.CycDate) = # ''' + str(endday) + ''' #)
			ORDER BY 
				qry_SMATranswALAll.[CycDate]
				,qry_SMATranswALAll.[Cslt]
				,qry_SMATranswALAll.[Client Number]
				,qry_SMATranswALAll.[Account Number];
		'''
	#-------------------------------------	
	#print sql
	#with open("Output.txt", "w") as text_file:
	#	text_file.write(sql)
	#sys.exit("done")	
	#--------------------------------------	
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
	dfnbrate = dbq.df_select(driver, db_file, sql)
	dftrans["NBRate"] = dfnbrate.iloc[0,0]
	#-----------------------------------------------------------------------

	#--------- get sales bonus rate for cslt under year 4 ----------
	sql = '''SELECT FSalesBonusRate.Level AS Tenure, FSalesBonusRate.Rate FROM FSalesBonusRate''' 
	dfsbrate = dbq.df_select(driver, db_file, sql)
	dftrans = dftrans.merge(dfsbrate, how="left", on="Tenure")

	#--------- get sales bonus rate based on AL ----------
	sql = '''SELECT DISTINCT FTransitionalSalesBonusRate.Level AS EarnedAL, FTransitionalSalesBonusRate.Rate AS ALRate FROM FTransitionalSalesBonusRate WHERE FTransitionalSalesBonusRate.ALYear = ''' + str(endday.year)
	dfalrate = dbq.df_select(driver, db_file, sql)
	dftrans = dftrans.merge(dfalrate, how="left", on="EarnedAL")
	dftrans.loc[dftrans['Tenure'] < 4, 'SBRate'] = dftrans['Rate']
	dftrans.loc[dftrans['Tenure'] > 3, 'SBRate'] = dftrans['ALRate']

	#--------- get sales bonus rate based on AL for advancing AL ----------
	sql = '''SELECT DISTINCT FTransitionalSalesBonusRate.Level AS AdvanceAL, FTransitionalSalesBonusRate.Rate AS AdvRate FROM FTransitionalSalesBonusRate WHERE FTransitionalSalesBonusRate.ALYear = ''' + str(endday.year)
	dfadvalrate = dbq.df_select(driver, db_file, sql)
	dftrans = dftrans.merge(dfadvalrate, how="left", on="AdvanceAL")
	dftrans["AdvRate"].fillna(0, inplace = True)
	#-----------------------------------------------------------------------

	#----------- calculate sales bonus/Advancing AL sales bonus/New Business  ------------
	dftrans['Cslt'] = dftrans['Cslt'].astype(np.int64)
	dftrans['New Business'] = dftrans['Event Gross Amount'] * dftrans['NBRate'] * dftrans['TransType']
	dftrans['New Business'] = dftrans['New Business'].round(2)
	#dftrans['Sales Bonus'] = np.where(dftrans['TransType'] == 1, dftrans['Event Gross Amount'] * dftrans['SBRate'], 0.00)
	dftrans['Sales Bonus'] = dftrans['Event Gross Amount'] * dftrans['SBRate'] * dftrans['TransType']
	dftrans['Sales Bonus'] = dftrans['Sales Bonus'].round(2)
	dftrans['AL Advancing Adj'] = np.where(dftrans['AdvRate'] != 0, (dftrans['Event Gross Amount'] * dftrans['AdvRate'] - dftrans['Sales Bonus']) * dftrans['TransType'], 0.00)
	dftrans['AL Advancing Adj'] = dftrans['AL Advancing Adj'].round(2)
	dftrans['Mark'] = np.where(dftrans['TransType'] == 0, '*', '')
	print(dftrans)
	sys.exit()
	
	##----------------------------------------------------------
	#---- old process, not used due to system changed ----
	##---- prepare New Business upload tab ------
	#dfnbtrans = dftrans.loc[((dftrans['TransType'] == 1) & (dftrans['New Business'] != 0)), ['Cslt', 'New Business']].copy()
	#dfnbtotal = dfnbtrans.groupby('Cslt', as_index=False)['New Business'].sum()
	#dfnbtotal = dfnbtotal.assign(Description='SMA Sales Credits ' + endday.strftime('%Y%m%d'))
	##print(dfnbtotal)
    #
	##---- prepare Sales Bonus upload file ------
	#dfsbtrans = dftrans.loc[((dftrans['TransType'] == 1) & (dftrans['Sales Bonus'] != 0)), ['Cslt', 'Sales Bonus']].copy()
	#dftotal = dfsbtrans.groupby('Cslt', as_index=False)['Sales Bonus'].sum()
	#
	#lssb = []
	#for index, row in dftotal.iterrows():
	#	sign = ''
	#	if row['Sales Bonus'] < 0 :
	#		sign = '-'
	#	else:
	#		sign = '+'
	#		
	#	lssb.append('D' + ig.addZero(str(int(row['Cslt'])), 5) + sign + ig.addZero(str(int(round((row['Sales Bonus'] * 100), 0))), 8) + '0102')
	#
	#colname = ['H29283N' + str(time.strftime("%Y%m%d")) + endday.strftime('%m%d%Y') + ' SMASALESBONUS']
	#dfsb = pd.DataFrame(lssb, columns=colname)	
	#print(dfsb)
	##----------------------------------------------------------
	
	#---- prepare Sale Bonus upload tab ------
	
	
	#---- prepare New Business upload tab ------
	
	#----------- Remove unrequired columns ------------
	dftrans.drop(['TransType', 'Tenure', 'Rate', 'ALRate'], axis=1, inplace=True)
	dftrans.rename(columns={'Event Gross Amount': 'Total Contribution'}, inplace=True)
	dftrans = dftrans[['CycDate','Cslt','EarnedAL','Name','RO','ROName','Account Number','Event Process Date','Client Number','Client Last Name', 'Client Given Name', 'Total Contribution','Mark','NBRate','New Business','SBRate','Sales Bonus', 'AdvanceAL','AdvRate','AL Advancing Adj']]
	dftrans.sort_values(['CycDate','Cslt', 'Client Number', 'Account Number', 'Total Contribution'], inplace=True)
	print(dftrans)
	
	dfsb.to_csv(endday.strftime('%m%d%Y') + ' SMASalesBonus.prn', index=False)	#output Sales Bonus upload file

	# Create a Pandas Excel writer using XlsxWriter as the engine.
	writer = pd.ExcelWriter('SMADailyAudit' + endday.strftime("%Y%m%d") + '.xlsx', engine='xlsxwriter')
	#writer = pd.ExcelWriter('F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMADaily' + endday.strftime("%Y%m%d") + '.xlsx', engine='xlsxwriter')

	# Convert the dataframe to an XlsxWriter Excel object.
	dftrans.to_excel(writer, sheet_name='report', index=False)
	dfnbtotal.to_excel(writer, sheet_name='NB', index=False)

	# Close the Pandas Excel writer and output the Excel file.
	writer.save()
