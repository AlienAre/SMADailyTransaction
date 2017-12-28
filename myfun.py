
import datetime 
from datetime import date
from dateutil import relativedelta

def getCycleStartDate(pdate):
	if int(pdate.strftime("%d")) > 15:
		return pdate.replace(day=1)
	else: 
		return (pdate.replace(day=1) - datetime.timedelta(days=1)).replace(day=16)

def getCycleEndDate(pdate):
	if int(pdate.strftime("%d")) > 15:
		return pdate.replace(day=15)
	else: 
		return pdate.replace(day=1) - datetime.timedelta(days=1)
		
def getLastMonthEndDate(pdate):
		return pdate.replace(day=1) - datetime.timedelta(days=1)
		
def getLast2MonthEndDate(pdate):
		return (pdate.replace(day=1) - datetime.timedelta(days=1)).replace(day=1) - datetime.timedelta(days=1) 		
		
def getQuarter(pdate):
		return (int(pdate.strftime("%m")) - 1)//3 + 1

def getLastQuarterEndDate(pdate):
    if pdate.month < 4:
        return datetime.date(pdate.year - 1, 12, 31)
    elif pdate.month < 7:
        return datetime.date(pdate.year, 3, 31)
    elif pdate.month < 10:
        return datetime.date(pdate.year, 6, 30)
    return datetime.date(pdate.year, 9, 30)
	
def getTenure(pdate1, pdate2):
	#date1 = datetime.datetime.strptime(str(pdate1), '%Y-%m-%d')
	#date2 = datetime.datetime.strptime(str(pdate2), '%Y-%m-%d')
	r = relativedelta.relativedelta(pdate1, pdate2)
	#print "{0.years} years and {0.months} months".format(r)
	return abs(r.years)
		
if __name__=="__main__":	
	#today = datetime.datetime.strptime('01/26/2017', '%m/%d/%Y') #date.today()	
	today = date.today()	
	startday = getCycleStartDate(today)
	endday = getCycleEndDate(today)
	print getQuarter(today)
	print today.year
	print getLastQuarterEndDate(today)
	#print startday
	#print endday
	#
	#date1 = datetime.datetime.strptime(str('2017-10-31'), '%Y-%m-%d')
	#date2 = datetime.datetime.strptime(str('2010-12-25'), '%Y-%m-%d')
	#r = relativedelta.relativedelta(date2, date1)
	#print "{0.years} years and {0.months} months".format(r)
	#print abs(r.years)