from datetime import datetime,date
import calendar  	

dayOfWeek = datetime.now().weekday()
print dayOfWeek

dayOfWeek = datetime(2008, 2, 16).weekday()
print dayOfWeek

cal = calendar.month(2016, 8)
print type(cal)


cur_cal = range(calendar.monthrange(2016, 8)[1]+1)[1:]
print type(cur_cal)

cur_cal = ['1','34']
cur_cal = cur_cal[1:]
print cur_cal
cur_cal = cur_cal[1:]
print cur_cal

"""
        cur_year = int(new_row[2].split('/')[0])
        cur_month = int(new_row[2].split('/')[1])
        cur_cal = range(calendar.monthrange(cur_year, cur_month)[1]+1)[1:]
        for cc in cur_cal:
        	cur_date = int(new_row[2].split('/')[2])
        	if cc != cur_date:

"""


sdf = ['df','dfd']
print sdf

dtstr = '21:32:12'
dsds = datetime.strptime(dtstr, "%H:%M:%S").time()
print dsds