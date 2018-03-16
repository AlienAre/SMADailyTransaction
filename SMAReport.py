import os, re, sys, time, xlrd
from datetime import date
import fnmatch
import numpy as np
import pandas as pd
import itertools as it
from openpyxl import load_workbook
import myfun as dd
import pyodbc
import datetime

#------------ get cycle start date ----------------
def getStartDate(pdate):
	if int(pdate.strftime("%d")) > 15:
		return pdate.replace(day=16)
	else: 
		return pdate.replace(day=1)
#--------------------------------------------------
		
## dd/mm/yyyy format
print 'Process date is ' + str(time.strftime("%d/%m/%Y"))
print 'Please enter the cycle end date (mm/dd/yyyy) you want to output:'

getcycledate = datetime.datetime.strptime(raw_input(), '%m/%d/%Y')
#startday = dd.getCycleStartDate(date.today())
#endday = dd.getCycleEndDate(date.today())
endday = getcycledate
startday = getStartDate(getcycledate)

print 'Cycle start date is ' + str(startday)
print 'Cycle end date is ' + str(endday)

#----------- get SMA daily transactions and AL information ------------
driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
db_file = r"F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA.accdb;"
user = "admin"
password = ""
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)
conn = pyodbc.connect(odbc_conn_str)

sql = '''SELECT * FROM qry_SMATranswAL WHERE CycDate = #''' + str(endday) + '''#'''
#print sql
#sys.exit("done")
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
Transdf['New Business'] = Transdf['Event Gross Amount'] * Transdf['NBRate'] * Transdf['TransType']
#Transdf['Sales Bonus'] = np.where(Transdf['TransType'] == 1, Transdf['Event Gross Amount'] * Transdf['SBRate'], 0.00)
Transdf['Sales Bonus'] = Transdf['Event Gross Amount'] * Transdf['SBRate'] * Transdf['TransType']
Transdf['AL Advancing Adj'] = np.where(Transdf['AdvRate'] != 0, (Transdf['Event Gross Amount'] * Transdf['AdvRate'] - Transdf['Sales Bonus']) * Transdf['TransType'], 0.00)
Transdf['Mark'] = np.where(Transdf['TransType'] == 0, '*', '')

#----------- Remove unrequired columns ------------
Transdf.drop(['StartDate', 'TermDate', 'TransType', 'Tenure', 'Rate', 'ALRate'], axis=1, inplace=True)
Transdf.rename(columns={'Event Gross Amount': 'Total Contribution'}, inplace=True)
Transdf = Transdf[['CycDate','Cslt','EarnedAL','Name','RO','ROName','Account Number','Event Process Date','Client Number','Client Last Name', 'Client Given Name', 'Total Contribution','Mark','NBRate','New Business','SBRate','Sales Bonus', 'AdvanceAL','AdvRate','AL Advancing Adj']]
Transdf.sort_values(['Cslt', 'Client Number', 'Account Number'], inplace=True)

#print Transdf#.head()
#sys.exit("done")

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('SMADailyAudit' + endday.strftime("%Y%m%d") + '.xlsx', engine='xlsxwriter')
#writer = pd.ExcelWriter('F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMADaily' + endday.strftime("%Y%m%d") + '.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
Transdf.to_excel(writer, sheet_name='Sheet1', index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
