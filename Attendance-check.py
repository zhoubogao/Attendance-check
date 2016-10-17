#-*-coding:utf-8-*-

import xlrd
import xlwt
import time
import datetime
import calendar

def notSame(nl):
    ne = []
    ind = 0
    for n in nl:
        if n not in ne:
            ne.append(n)
        else:
            del nl[ind]
        ind += 1
print 'start *******'
data = xlrd.open_workbook('111.xls')
table = data.sheets()[0]
allrows = []
for rownum in range(table.nrows):
    allrows.append(table.row_values(rownum))
allrows = allrows[1:]
nl = []
nameTonl = []
for ri in allrows:
    nl.append(ri[1])
    nameTonl.append([ri[0], ri[1]])

num_list = list(set(nl))
result_list = []
num_list = sorted(num_list, key=lambda num_list: int(num_list))
cur_calendar_days = 0
for num in num_list:
    one_person_list = []
    new_row = []
    time_list = []
    date_list = []
    for ri in allrows:
        if num != ri[1]:
            continue
        time_list.append(ri[2])
        name = ri[0]
    for t in time_list:
        date_list.append(t.split()[0])
    date_list = list(set(date_list))
    date_list = sorted(date_list,
                       key=lambda date_list: 
                       int(date_list.replace('/', '')))
    for d in date_list:
        same_date_list = []
        for t in time_list:
            if d != t.strip().split()[0]:
                continue
            same_date_list.append(t)
        if len(same_date_list) > 2:
            same_date_list = sorted(same_date_list, 
                                    key=lambda same_date_list:
                                    int(same_date_list.split()[1].replace(':',"").strip()))
            same_date_list = [same_date_list[0], same_date_list[-1]]
        new_row.append(name)
        new_row.append(num)
        new_row.append(d)
        if len(same_date_list) == 0:
            new_row.append(u' ')
            new_row.append(u' ')
        elif len(same_date_list) == 1:
            same_date_list = same_date_list[0]#.split()[1])
            if int(same_date_list.split()[1].replace(':', '')) >= 120000:
                new_row.append(u' ')
                new_row.append(same_date_list)
            else:
                new_row.append(same_date_list)
                new_row.append(u' ')
        else:
            new_row.append(same_date_list[0])#.split()[1])
            new_row.append(same_date_list[1])#.split()[1])

        one_person_list.append(new_row)
        new_row = []

    cur_year = int(one_person_list[0][2].split('/')[0])
    cur_month = int(one_person_list[0][2].split('/')[1])
    cur_cal = range(calendar.monthrange(cur_year, cur_month)[1]+1)[1:]
    cur_cal_temp = cur_cal
    cur_calendar_days = len(cur_cal)   #当月多少天 
    if cur_cal[-1] != int(one_person_list[-1][2].split('/')[2]):
        insert_item = one_person_list[-1][:2]
        insert_item.append(str(cur_year) + '/' + str(cur_month) + '/' +str(cur_cal[-1]))
        insert_item.append(u' ')
        insert_item.append(u' ')
        one_person_list.append(insert_item)
    for cc in cur_cal:
        for one in one_person_list:
            cur_date = int(one[2].split('/')[2])
            if cc == cur_date:
            	break
            elif cc < cur_date:
                insert_item = one[:2]
                insert_item.append(str(cur_year) + '/' + str(cur_month) + '/' +str(cc))
                insert_item.append(u' ')
                insert_item.append(u' ')
                one_person_list.insert(one_person_list.index(one), insert_item)
                break
            else:
            	pass
      
    for one in one_person_list:
        cur_date = int(one[2].split('/')[2])
        dayOfWeek = datetime.datetime(cur_year, cur_month, cur_date).isoweekday()

        if (dayOfWeek > 5 or cur_date == 15 or cur_date == 16 or cur_date == 30) and cur_date != 18 :
            one.append(u'正常')
            one.append(u' ')
        else:
            if one[3] == u' ' and one[4] == u' ':
                one.append(u'异常')
                one.append(u'无打卡记录')
            elif one[3] != u' ' and one[4] == u' ':
                one.append(u'异常')
                one.append(u'下班未打卡')
            elif one[3] == u' ' and one[4] != u' ':
                one.append(u'异常')
                one.append(u'上班未打卡')
            else:
                endTime = datetime.datetime.strptime(one[4].split()[1].strip(), "%H:%M:%S")
                startTime = datetime.datetime.strptime(one[3].split()[1].strip(), "%H:%M:%S")
                dura = (endTime - startTime)
                if dura >= datetime.timedelta(hours=9,minutes=15,seconds=00) \
                        and int(one[3].split()[1].strip().replace(':', '')) <= 95500 :
                    one.append(u'正常')
                    one.append(u' ')
                elif int(one[3].split()[1].strip().replace(':', '')) > 95500:
                    one.append(u'异常')
                    one.append(u'迟到')
                else:
                    one.append(u'异常') 
                    one.append(u'未满8小时')

        if dayOfWeek == 1:
            dayOfWeek = u'星期一'
        elif dayOfWeek == 2:
            dayOfWeek = u'星期二'
        elif dayOfWeek == 3:
            dayOfWeek = u'星期三'
        elif dayOfWeek == 4:
            dayOfWeek = u'星期四'
        elif dayOfWeek == 5:
            dayOfWeek = u'星期五'
        elif dayOfWeek == 6:
            dayOfWeek = u'星期六'
        elif dayOfWeek == 7:
            dayOfWeek = u'星期日'
        else:
        	dayOfWeek = u'未知'
        one[2] = one[2] + u' ' + dayOfWeek
    result_list.extend(one_person_list)

result_list = sorted(result_list, key=lambda result_list: int(result_list[1]))
#for r in result_list:
#    print r

#******************请假统计**************************
data = xlrd.open_workbook('222.xls')
table = data.sheets()[0]
allrows = []
for rownum in range(table.nrows):
    allrows.append(table.row_values(rownum))
allrows = allrows[1:]
ext_nl = []
leave_nl = []
for ri in allrows:
    ext_nl.append([ri[2], ri[3], ri[4], ri[5],ri[6]])
    leave_nl.append([ri[2], ri[6]])
#ext_nl = ext_nl[1:]
#leave_nl = leave_nl[1:]
print ext_nl
ext_nl = sorted(ext_nl, key=lambda ext_nl: int(ext_nl[0]))
leave_nl = sorted(leave_nl, key=lambda leave_nl: int(leave_nl[0]))
notSame(ext_nl)
notSame(leave_nl)

row = 0
for n in ext_nl:
    for ii in result_list:
        if n[0] == ii[1] and \
            n[2].strip().split()[0].replace('-', '/').replace('/0', '/')\
             == ii[2].strip().split()[0] :
            ii.extend(n[1:])
            break
for ii in result_list:
    if len(ii) < 11:
        ii.extend(['', '', '', ''])


#*******************加班****************************
data = xlrd.open_workbook('333.xls')
table = data.sheets()[0]
allrows = []
for rownum in range(table.nrows):
    allrows.append(table.row_values(rownum))
allrows = allrows[1:]
ext_nl = []
tj_nl = []
for ri in allrows:
    ext_nl.append([ri[2], ri[3], ri[4], ri[5],ri[6]])
    tj_nl.append([ri[2], ri[3], ri[6]])
#ext_nl = ext_nl[1:]
ext_nl = sorted(ext_nl, key=lambda ext_nl: int(ext_nl[0]))
#tj_nl = tj_nl[1:]
tj_nl = sorted(tj_nl, key=lambda tj_nl: int(tj_nl[0]))
notSame(ext_nl)  
notSame(tj_nl)

row = 0
for n in ext_nl:
    for ii in result_list:
        if n[0] == ii[1] and \
            n[2].strip().split()[0].replace('-', '/').replace('/0', '/')\
             == ii[2].strip().split()[0] :
            ii.extend(n[1:])
            break
for ii in result_list:
    if len(ii) < 15:
        ii.extend(['', '', '', ''])

#*******************加班与请假统计**************************** 
onwork_money = []
onwork_time = []
leave_time = []
onwork_statistics = []

cl = []
for x in tj_nl:
    cl.append(x[0])
    pass
cl = list(set(cl))
for x in cl:
    onwork_money_tm = 0.0
    onwork_time_tm = 0.0
    for c in tj_nl:
        if c[0] == x and c[1] == u"安排调休":
            onwork_time_tm += float(c[2])
        elif c[0] == x and c[1] != u"安排调休":
            onwork_money_tm += float(c[2])
    onwork_money.append([x, onwork_money_tm])
    onwork_time.append([x, onwork_time_tm])

ll = []
for x in leave_nl:
    ll.append(x[0])
    pass
ll = list(set(ll))
for x in ll:
    leave_time_tm = 0.0
    for l in leave_nl:
        if l[0] == x:
            leave_time_tm += float(l[1])
    leave_time.append([x, leave_time_tm])


#*******************总剩余加班结余****************************
data = xlrd.open_workbook('444.xls')
table = data.sheets()[0]
allrows = []
for rownum in range(table.nrows):
    allrows.append(table.row_values(rownum))
allrows = allrows[1:]
allJb_nl = []
for ri in allrows:
    if ri[1]  != '':
        allJb_nl.append([ ri[1], ri[2] ])
#allJb_nl = allJb_nl[:-1]
allJb_nl = sorted(allJb_nl, key=lambda allJb_nl: int(allJb_nl[0]))
notSame(allJb_nl)  
print allJb_nl

#*******************加班总统计**************************** 
mouth_statistics = [] # 工号 加班补偿 加班调休 请假总时长 结余时长 总剩余加班结余
for x in num_list:
    mouth_item = []
    mouth_item.append(x)
    mach = False
    for om in onwork_money:
        if x == om[0]:
            mach = True
            mouth_item.append(str(om[1]))
            break
    if not mach:
        mouth_item.append(u'0.0')
    mach = False
    for ot in onwork_time:
        if x == ot[0]:
            mach = True
            mouth_item.append(str(ot[1]))
            break
    if not mach:
        mouth_item.append(u'0.0')
    mach = False
    for lt in leave_time:
        if x == lt[0]:
            mach = True
            mouth_item.append(str(lt[1]))
            break
    if not mach:
        mouth_item.append(u'0.0')
    mouth_item.append(str(float(mouth_item[2]) - float(mouth_item[3])))
    mach = False
    for lt in allJb_nl:
        if int(x) == int(lt[0]):
            mach = True
            mouth_item.append(str(float(lt[1]) + float(mouth_item[2]) - float(mouth_item[3])))
            break
    if not mach:
        mouth_item.append(str(float(mouth_item[2]) - float(mouth_item[3])))

    mouth_statistics.append(mouth_item)

'''
#*******************总剩余加班结余添加工号****************************
data = xlrd.open_workbook('555.xls')
table = data.sheets()[0]
allrows = []
for rownum in range(table.nrows):
    allrows.append(table.row_values(rownum))
allrows = allrows[1:]
jb_nl = []
for ri in allrows:
    jb_nl.append([ri[0], u'', str(ri[2])])
for j in jb_nl:
    for nn in nameTonl:
        if type(nn[0]) == type(j[0]) and j[0].strip() == nn[0].strip():
            j[1] = str(int(nn[1]))
            break

wb = xlwt.Workbook(encoding='utf-8')
table = wb.add_sheet('AttendanceResult', cell_overwrite_ok=True)
table.write(0, 0, u'姓名')
table.write(0, 1, u'工号')
table.write(0, 2, u'剩余加班总计')
row = 1
for jb in jb_nl:
    col = 0
    for i in jb:
        table.write(row, col, i)
        col += 1
    row += 1

now = time.strftime('%Y-%m-%d-%H-%M-%S', time.localtime(time.time()))
wb.save('jiaban-' + now + '.xls')
'''
#*******************写exl****************************
wb = xlwt.Workbook(encoding='utf-8')
table = wb.add_sheet('AttendanceResult', cell_overwrite_ok=True)
borders = xlwt.Borders()
borders.left = 1
borders.right = 1
borders.top = 1
borders.bottom = 1
#borders.bottom_colour=0x3A
alignment = xlwt.Alignment() # Create Alignment
# May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT,
#HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
alignment.horz = xlwt.Alignment.HORZ_CENTER   #水平居中
# May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
alignment.vert = xlwt.Alignment.VERT_CENTER    #垂直居中

style_pulic = xlwt.easyxf('pattern: pattern solid, fore_colour white;');
style_pulic.borders = borders
style_pulic.alignment = alignment
styleYellowBkg = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;');
styleYellowBkg.borders = borders
styleYellowBkg.alignment = alignment
styleSkyBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour sky_blue; font: bold on;');
styleSkyBlueBkg.borders = borders
styleSkyBlueBkg.alignment = alignment
style_center = xlwt.easyxf('pattern: pattern solid, fore_colour white;');
style_center.borders = borders
style_center.alignment = alignment # Add Alignment to Style
#styleBlueBkg = xlwt.easyxf('font: color-index red, bold on');
#styleBlueBkg = xlwt.easyxf('font: background-color-index red, bold on');
#styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour red;');
#styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour blue;');
#styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue;');
#styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;');
#styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour dark_blue;');
#styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour dark_blue_ega;');
#styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour ice_blue;');
#styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour ocean_blue; font: bold on;'); # 80% like
#styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour sky_blue;');
table.write(0, 0, u'姓名', styleSkyBlueBkg)
table.write(0, 1, u'工号', styleSkyBlueBkg)
table.write(0, 2, u'日期', styleSkyBlueBkg)
table.write(0, 3, u'上班打卡时间', styleSkyBlueBkg)
table.write(0, 4, u'下班打卡时间', styleSkyBlueBkg)
table.write(0, 5, u'状态', styleSkyBlueBkg)
table.write(0, 6, u'异常情况', styleSkyBlueBkg)

table.write(0, 7, u'请假类型', styleSkyBlueBkg)
table.write(0, 8, u'请假开始时间', styleSkyBlueBkg)
table.write(0, 9, u'请假结束时间', styleSkyBlueBkg)
table.write(0, 10, u'请假天数', styleSkyBlueBkg)

table.write(0, 11, u'加班处理', styleSkyBlueBkg)
table.write(0, 12, u'加班开始时间', styleSkyBlueBkg)
table.write(0, 13, u'加班结束时间', styleSkyBlueBkg)
table.write(0, 14, u'加班时长(小时)', styleSkyBlueBkg)

table.write(0, 15, u'加班补偿金(小时)', styleSkyBlueBkg)
table.write(0, 16, u'可调休（小时）', styleSkyBlueBkg)
table.write(0, 17, u'请假（小时）', styleSkyBlueBkg)
table.write(0, 18, u'当月加班时长结余', styleSkyBlueBkg)
table.write(0, 19, u'总剩余加班结余', styleSkyBlueBkg)

row = 1
for ii in result_list:
    col = 0
    for i in ii:
        if i == u'异常':
            table.write(row, col, i, styleYellowBkg)
        else:
            table.write(row, col, i, style_pulic)
        col += 1
    row += 1

row = 1
col = 15
for ms in mouth_statistics:
    for m in ms[1:]:
        table.write_merge(row, row + cur_calendar_days - 1, col, col, m, style_center)
        col +=1
    col = 15
    row += cur_calendar_days

now = time.strftime('%Y-%m-%d-%H-%M-%S', time.localtime(time.time()))
wb.save('AtteChk-' + now + '.xls')
print 'Finished . please check the exl result.'