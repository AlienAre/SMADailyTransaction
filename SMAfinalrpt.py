#--------------------------------------------
#version:		1.0.0.1
#author:		West
#Description:	prepare final SMA cycle report and Sale Bonus/New Business upload file
#				suppose to run every cycle after manager reviews the transaction and db is updated
#Assumptions:	
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
	print 'Process date is ' + str(time.strftime("%m/%d/%Y"))
	print 'Please enter the cycle end date (mm/dd/yyyy) you want to output:'

	getcycledate = datetime.datetime.strptime(raw_input(), '%m/%d/%Y')
	endday = getcycledate
	startday = ig.getCStartDate(getcycledate)

	print 'Cycle start date is ' + str(startday)
	print 'Cycle end date is ' + str(endday)

	#----------- get SMA daily transactions and AL information ------------
	driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
	db_file = r"F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA.accdb;"
	user = "admin"
	password = ""
	odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)
	conn = pyodbc.connect(odbc_conn_str)

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
	Transdf = pd.read_sql_query(sql,conn)
	conn.close()
	#-----------------------------------------------------------------------

	#----------- get Sales Bonus rate ------------
	driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
	db_file = r"F:\Files For\West Wang\Rates.accdb;"
	user = "admin"
	password = ""
	odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)
	conn = pyodbc.connect(odbc_conn_str)

	#--------- get New Business rate based on AL for advancing AL ----------
	sql = '''SELECT DISTINCT NewBusinessRate.Rate AS NBRate FROM NewBusinessRate WHERE NewBusinessRate.NBYear = ''' + str(endday.year)
	NBRatedf = pd.read_sql_query(sql,conn)
	Transdf['NBRate'] = NBRatedf.at[(0, 'NBRate')]
	#-----------------------------------------------------------------------

	#--------- get sales bonus rate for cslt under year 4 ----------
	sql = '''SELECT FSalesBonusRate.Level AS Tenure, FSalesBonusRate.Rate FROM FSalesBonusRate''' 
	Ratedf = pd.read_sql_query(sql,conn)
	Transdf = Transdf.merge(Ratedf, left_on='Tenure', right_on='Tenure', how='left')

	#--------- get sales bonus rate based on AL ----------
	sql = '''SELECT DISTINCT FTransitionalSalesBonusRate.Level AS EarnedAL, FTransitionalSalesBonusRate.Rate AS ALRate FROM FTransitionalSalesBonusRate WHERE FTransitionalSalesBonusRate.ALYear = ''' + str(endday.year)
	ALratedf = pd.read_sql_query(sql,conn)
	Transdf =  Transdf.merge(ALratedf, left_on='EarnedAL', right_on='EarnedAL', how='left')
	Transdf.loc[Transdf['Tenure'] < 4, 'SBRate'] = Transdf['Rate']
	Transdf.loc[Transdf['Tenure'] > 3, 'SBRate'] = Transdf['ALRate']

	#--------- get sales bonus rate based on AL for advancing AL ----------
	sql = '''SELECT DISTINCT FTransitionalSalesBonusRate.Level AS AdvanceAL, FTransitionalSalesBonusRate.Rate AS AdvRate FROM FTransitionalSalesBonusRate WHERE FTransitionalSalesBonusRate.ALYear = ''' + str(endday.year)
	Advratedf = pd.read_sql_query(sql,conn)
	Transdf =  Transdf.merge(Advratedf, left_on='AdvanceAL', right_on='AdvanceAL', how='left')
	Transdf['AdvRate'].fillna(0, inplace = True)
	conn.close()
	#-----------------------------------------------------------------------

	#----------- calculate sales bonus/Advancing AL sales bonus/New Business  ------------
	Transdf['Cslt'] = Transdf['Cslt'].astype(np.int64)
	Transdf['New Business'] = Transdf['Event Gross Amount'] * Transdf['NBRate'] * Transdf['TransType']
	#Transdf['Sales Bonus'] = np.where(Transdf['TransType'] == 1, Transdf['Event Gross Amount'] * Transdf['SBRate'], 0.00)
	Transdf['Sales Bonus'] = Transdf['Event Gross Amount'] * Transdf['SBRate'] * Transdf['TransType']
	Transdf['AL Advancing Adj'] = np.where(Transdf['AdvRate'] != 0, (Transdf['Event Gross Amount'] * Transdf['AdvRate'] - Transdf['Sales Bonus']) * Transdf['TransType'], 0.00)
	Transdf['Mark'] = np.where(Transdf['TransType'] == 0, '*', '')

	#---- prepare New Business upload tab ------
	dfnbtrans = Transdf.loc[((Transdf['TransType'] == 1) & (Transdf['New Business'] != 0)), ['Cslt', 'New Business']].copy()
	dfnbtotal = dfnbtrans.groupby('Cslt', as_index=False)['New Business'].sum()
	dfnbtotal = dfnbtotal.assign(Description='SMA Sales Credits ' + endday.strftime('%Y%m%d'))
	print dfnbtotal

	#---- prepare Sales Bonus upload file ------
	dfsbtrans = Transdf.loc[((Transdf['TransType'] == 1) & (Transdf['Sales Bonus'] != 0)), ['Cslt', 'Sales Bonus']].copy()
	dftotal = dfsbtrans.groupby('Cslt', as_index=False)['Sales Bonus'].sum()
	
	lssb = []
	for index, row in dftotal.iterrows():
		sign = ''
		if row['Sales Bonus'] < 0 :
			sign = '-'
		else:
			sign = '+'
			
		lssb.append('D' + ig.addZero(str(int(row['Cslt'])), 5) + sign + ig.addZero(str(int(round((row['Sales Bonus'] * 100), 0))), 8) + '0102')
	
	colname = ['H29283N' + str(time.strftime("%Y%m%d")) + endday.strftime('%m%d%Y') + ' SMASALESBONUS']
	dfsb = pd.DataFrame(lssb, columns=colname)	
	print dfsb

	#----------- Remove unrequired columns ------------
	Transdf.drop(['TransType', 'Tenure', 'Rate', 'ALRate'], axis=1, inplace=True)
	Transdf.rename(columns={'Event Gross Amount': 'Total Contribution'}, inplace=True)
	Transdf = Transdf[['CycDate','Cslt','EarnedAL','Name','RO','ROName','Account Number','Event Process Date','Client Number','Client Last Name', 'Client Given Name', 'Total Contribution','Mark','NBRate','New Business','SBRate','Sales Bonus', 'AdvanceAL','AdvRate','AL Advancing Adj']]
	Transdf.sort_values(['CycDate','Cslt', 'Client Number', 'Account Number', 'Total Contribution'], inplace=True)
	print Transdf
	
	dfsb.to_csv(endday.strftime('%m%d%Y') + ' SMASalesBonus.prn', index=False)	#output Sales Bonus upload file

	# Create a Pandas Excel writer using XlsxWriter as the engine.
	writer = pd.ExcelWriter('SMADailyAudit' + endday.strftime("%Y%m%d") + '.xlsx', engine='xlsxwriter')
	#writer = pd.ExcelWriter('F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMADaily' + endday.strftime("%Y%m%d") + '.xlsx', engine='xlsxwriter')

	# Convert the dataframe to an XlsxWriter Excel object.
	Transdf.to_excel(writer, sheet_name='report', index=False)
	dfnbtotal.to_excel(writer, sheet_name='NB', index=False)

	# Close the Pandas Excel writer and output the Excel file.
	writer.save()
