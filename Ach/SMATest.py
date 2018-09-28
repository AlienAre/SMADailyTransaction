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
cashlabels = ['Event Record Type', 'Event Effective Date', 'Event Process Date', 'Event Activity Type', 'Event Activity Description', 'Event Gross Amount', 'Plan Product Code', 'Account Market Value', 'Client Number', 'Client Last Name', 'Client Given Name', 'Client Servicing Consultant Number', 'Client Deceased Indicator', 'Client Company Name', 'Client Province Code', 'Account Number', 'Account Dealer Code', 'Account IGSI Net Share Quantity', 'Product Code', 'Product Share Price Amount', 'Product IGSI Symbol', 'Product Description', 'Product Security Type', 'Product Security Class', 'Product Security Category']
SMAdata = pd.DataFrame()	#set blank data frame for SMA daily use
SMAcashdata = pd.DataFrame()	#set blank data frame for SMA cash daily use

SMAlist = []
SMAcashlist = []

pattern = '*SMA.EVENTS*.xls'	#use to find SMA daily files
cashpattern = '*SMA.CASH.EVENTS*.xls'	#use to find SMA cash daily files

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
SMAdata = SMAdata.sort_values('Event Effective Date')
#print SMAdata['Event Effective Date'].dtype

#SMAdata['CycleMonth'] = SMAdata['Event Effective Date'].astype(str).str[:6]
SMAdata['CycleMonth'] = endday.strftime("%Y%m%d")

'''
#################################################
#SMA trans do not include revesals, now use SMA cash trans to calculate
#data frame for Sales Credit upload
SMAtotal = SMAdata.groupby(['CycleMonth', 'Client Servicing Consultant Number'], as_index=False)['Event Gross Amount'].sum()
SMAtotal['SCAmount'] = SMAtotal['Event Gross Amount'] * 0.7
SMAtotal['ExchangeRate'] = 0.7
SMAtotal['Description'] = 'SMA Sales Credits ' + SMAtotal['CycleMonth']
SMAtotal = SMAtotal[['CycleMonth', 'Client Servicing Consultant Number', 'SCAmount', 'Description', 'Event Gross Amount', 'ExchangeRate']]
#SMAtotal.drop('Event Gross Amount', axis=1, inplace=True)
#print SMAdata.head(3)
#################################################
'''

### go to dir and get all SMA.CASH.EVENTS excel list
files = os.listdir(filesdir)
for file in fnmatch.filter(files, cashpattern):
		SMAcashlist.append(os.path.join(filesdir, file))

for sma in SMAcashlist:
	if 'IGSI SMA DAILY CASH TRANSACTION' in xlrd.open_workbook(sma, on_demand=True).sheet_names():
		df = pd.read_excel(sma, sheetname='IGSI SMA DAILY CASH TRANSACTION', header=None)
		df1 = (df.loc[df[0] == 'D'])

		if not df1.empty:
			SMAcashdata = SMAcashdata.append(df1, ignore_index=True)

SMAcashdata.columns = cashlabels
SMAcashdata = SMAcashdata.sort_values(['Event Effective Date', 'Client Servicing Consultant Number', 'Client Number', 'Account Number'])
SMAcashdata['CycleMonth'] = endday.strftime("%Y%m%d")

#################################################
#SMA trans do not include revesals, now use SMA cash trans to calculate
#data frame for Sales Credit upload
SMAtotal = SMAcashdata.groupby(['CycleMonth', 'Client Servicing Consultant Number'], as_index=False)['Event Gross Amount'].sum()
SMAtotal['SCAmount'] = SMAtotal['Event Gross Amount'] * 1.0
SMAtotal['ExchangeRate'] = 1.0
SMAtotal['Description'] = 'SMA Sales Credits ' + SMAtotal['CycleMonth']
SMAtotal = SMAtotal[['CycleMonth', 'Client Servicing Consultant Number', 'SCAmount', 'Description', 'Event Gross Amount', 'ExchangeRate']]
#SMAtotal.drop('Event Gross Amount', axis=1, inplace=True)
#print SMAdata.head(3)
#################################################

book = load_workbook('F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA Daily.xlsx')
sheetSMA = book.get_sheet_by_name("SMA")

data = sheetSMA.values
cols = next(data)
data = list(data)
data = (it.islice(r, 0, None) for r in data)
SMAdf = pd.DataFrame(data, columns=cols)

sheetTransaction = book.get_sheet_by_name("Transaction")

data = sheetTransaction.values
cols = next(data)
data = list(data)
data = (it.islice(r, 0, None) for r in data)
Transactiondf = pd.DataFrame(data, columns=cols)

sheetUpload = book.get_sheet_by_name("Upload")
data = sheetUpload.values
cols = next(data)
data = list(data)
data = (it.islice(r, 0, None) for r in data)
Uploaddf = pd.DataFrame(data, columns=cols)
#sys.exit("done")
###################

driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
db_file = r"F:\Files For\West Wang\Rates.accdb;"
user = "admin"
password = ""
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)

sql = '''SELECT * FROM qry_CsltALList WHERE Cslt IN {} AND CycDate = #''' + str(endday) + '''#'''
sql = sql.format(str(tuple(SMAtotal['Client Servicing Consultant Number'])).replace(',)', ',0)'))
conn = pyodbc.connect(odbc_conn_str)
#cur = conn.cursor()
dfcslt = pd.read_sql_query(sql,conn)

sql = '''SELECT FSalesBonusRate.Level AS Tenure, FSalesBonusRate.Rate FROM FSalesBonusRate''' 
dfrate = pd.read_sql_query(sql,conn)
dfcslt = dfcslt.merge(dfrate, left_on='Tenure', right_on='Tenure', how='left')

sql = '''SELECT DISTINCT FTransitionalSalesBonusRate.Level AS EarnedAL, FTransitionalSalesBonusRate.Rate AS ALRate FROM FTransitionalSalesBonusRate  WHERE FTransitionalSalesBonusRate.ALYear = ''' + str(endday.year)
dfALrate = pd.read_sql_query(sql,conn)
#print sql
print dfALrate
dfcslt =  dfcslt.merge(dfALrate, left_on='EarnedAL', right_on='EarnedAL', how='left')
dfcslt['ActualRate'] = 0.00
dfcslt.loc[dfcslt['Tenure'] < 4, 'ActualRate'] = dfcslt['Rate']
dfcslt.loc[dfcslt['Tenure'] > 3, 'ActualRate'] = dfcslt['ALRate']

SMAtotal =  SMAtotal.merge(dfcslt, left_on='Client Servicing Consultant Number', right_on='Cslt', how='left')
SMAtotal['SalesBonus'] = SMAtotal['Event Gross Amount'] * SMAtotal['ActualRate']
SMAtotal['AdjSalesBonus'] = 0

SMAtotal.loc[SMAtotal['Tenure'] > 3 & (SMAtotal['AdvanceAL'] == SMAtotal['AL']), 'AdjSalesBonus'] = (SMAtotal['AdvanceAL'] - SMAtotal['EarnedAL']) * SMAtotal['Event Gross Amount'] * 0.0005

SMAtotal['SalesBonusPremium'] = (SMAtotal['SalesBonus'] + SMAtotal['AdjSalesBonus'].fillna(0)) * 0.25
#print SMAtotal['SalesBonus'] + SMAtotal['AdjSalesBonus'].fillna(0)
#### Remove additional columns do not need
del SMAtotal['TermDate']
del SMAtotal['Cslt']
#print SMAtotal
#sys.exit("done")
####################

######### test ###########
#print Transactiondf.head(3)
#Transactiondf.to_csv('ex.csv', index=False)
#SMAcashdata.to_csv('smac.csv', index=False)
##########################
#append current cycle data
SMAdf = SMAdf.append(SMAdata, ignore_index=True)
Transactiondf = Transactiondf.append(SMAcashdata, ignore_index=True)
Uploaddf = Uploaddf.append(SMAtotal, ignore_index=True)
######### test ###########
#print SMAcashdata.head(3)
#print Transactiondf.head(3)
#sys.exit("done")
##########################

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter("F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA Daily.xlsx", engine='openpyxl')

writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

# Convert the dataframe to an XlsxWriter Excel object.
SMAdf.to_excel(writer, sheet_name='SMA', index=False)
Transactiondf.to_excel(writer, sheet_name='Transaction', index=False)
Uploaddf.to_excel(writer, sheet_name='Upload', index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
