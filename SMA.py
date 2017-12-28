###########################################
#	SMA daily positon process
#
#	Version:		1.0
#	Author:			West	
#	Create Date: 	09/29/2017
###########################################

import os
import re
import fnmatch
import pandas as pd
from openpyxl import load_workbook
import xlrd
import time

## dd/mm/yyyy format
print (time.strftime("%d/%m/%Y"))

labels = ['Event Record Type', 'Event Effective Date', 'Event Process Date', 'Event Activity Type', 'Event Gross Amount', 'Plan Product Code', 'Account Market Value', 'Client Number', 'Client Last Name', 'Client Given Name', 'Client Servicing Consultant Number', 'Client Deceased Indicator', 'Client Company Name', 'Client Province Code', 'Account Number', 'Account Dealer Code', 'Account IGSI Net Share Quantity', 'Product Code', 'Product Share Price Amount', 'Product IGSI Symbol ', 'Product Description', 'Product Security Type', 'Product Security Class', 'Product Security Category']
cashlabels = ['Event Record Type', 'Event Effective Date', 'Event Process Date', 'Event Activity Type', 'Event Activity Description', 'Event Gross Amount', 'Plan Product Code', 'Account Market Value', 'Client Number', 'Client Last Name', 'Client Given Name', 'Client Servicing Consultant Number', 'Client Deceased Indicator', 'Client Company Name', 'Client Province Code', 'Account Number', 'Account Dealer Code', 'Account IGSI Net Share Quantity', 'Product Code', 'Product Share Price Amount', 'Product IGSI Symbol ', 'Product Description', 'Product Security Type', 'Product Security Class', 'Product Security Category']
SMAdata = pd.DataFrame()	#set blank data frame for SMA daily use
SMAcashdata = pd.DataFrame()	#set blank data frame for SMA cash daily use

SMAlist = []
SMAcashlist = []

def f(var):
    if isinstance(var, pd.DataFrame):
        print "do stuff"

pattern = '*SMA.EVENTS*.xls'	#use to find SMA daily files
cashpattern = '*SMA.CASH.EVENTS*.xls'	#use to find SMA cash daily files

### go to dir and get all SMA.EVENTS excel list	
files = os.listdir('F:\\3-Compensation Programs\\IIROC Compensation\\DAILY SMA EVENT')
for file in fnmatch.filter(files, pattern):
		SMAlist.append(os.path.join('F:\\3-Compensation Programs\\IIROC Compensation\\DAILY SMA EVENT', file))
#print SMAcashlist

### iterate all SMA.EVENTS excel files and extract data to df		
for sma in SMAlist:
	df = pd.read_excel(sma, header=None)
	df1 = (df.loc[df[0] == 'D'])
	
	if not df1.empty:
		SMAdata = SMAdata.append(df1, ignore_index=True)

SMAdata.columns = labels
SMAdata = SMAdata.sort_values('Event Effective Date')
#print SMAdata['Event Effective Date'].dtype

SMAdata['CycleMonth'] = SMAdata['Event Effective Date'].astype(str).str[:6]

#data frame for Sales Credit upload
SMAtotal = SMAdata.groupby(['CycleMonth', 'Client Servicing Consultant Number'], as_index=False)['Event Gross Amount'].sum()
SMAtotal['SCAmount'] = SMAtotal['Event Gross Amount'] * 0.7
SMAtotal['ExchangeRate'] = 0.7
SMAtotal['Description'] = 'SMA Sales Credits ' + SMAtotal['CycleMonth']
SMAtotal = SMAtotal[['CycleMonth', 'Client Servicing Consultant Number', 'SCAmount', 'Description', 'Event Gross Amount', 'ExchangeRate']]
#SMAtotal.drop('Event Gross Amount', axis=1, inplace=True)

### go to dir and get all SMA.CASH.EVENTS excel list	
files = os.listdir('F:\\3-Compensation Programs\\IIROC Compensation\\DAILY SMA EVENT')
for file in fnmatch.filter(files, cashpattern):
		SMAcashlist.append(os.path.join('F:\\3-Compensation Programs\\IIROC Compensation\\DAILY SMA EVENT', file))
#print SMAcashlist

for sma in SMAcashlist:
	if 'IGSI SMA DAILY CASH TRANSACTION' in xlrd.open_workbook(sma, on_demand=True).sheet_names():
		df = pd.read_excel(sma, sheetname='IGSI SMA DAILY CASH TRANSACTION', header=None)
		df1 = (df.loc[df[0] == 'D'])
	
		if not df1.empty:
			SMAcashdata = SMAcashdata.append(df1, ignore_index=True)
#print 'This si SMA'		
#print SMAData

SMAcashdata.columns = cashlabels
SMAcashdata = SMAcashdata.sort_values('Event Effective Date')

book = load_workbook('F:\\3-Compensation Programs\\IIROC Compensation\\SMA Business Credits\\SMA Daily.xlsx')

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter("F:\\3-Compensation Programs\\IIROC Compensation\\SMA Business Credits\\SMA Daily.xlsx", engine='openpyxl')

writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

# Convert the dataframe to an XlsxWriter Excel object.
SMAdata.to_excel(writer, sheet_name='SMA', index=False)
SMAcashdata.to_excel(writer, sheet_name='Transaction', index=False)
SMAtotal.to_excel(writer, sheet_name='Upload', index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
