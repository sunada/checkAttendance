#coding:utf - 8

import xlrd
import xlwt
#from openpyxl import load_workbook

import time
from datetime import date, datetime

import sys

#the colx of dep
DEPCOLX = 0
#the colx of name
NAMECOLX = 1
#the colx of record workday
WORKDAYCOLX = 2
#the colx to mark whether there are unusual records
MARKCOLX = 9
#import time point: 8:00, 10:00, 10:30, 16:00, 16:30, 19:00, 20:00
tpnts = [0.333333, 0.416667, 0.4375, 0.666666, 0.6875, 0.791667, 0.833333]

#test how to use xlrd
def test_xlrd(filename):
    start = time.time()
    book = xlrd.open_workbook(filename)
    sh = book.sheet_by_index(0)
    print "Worksheet name(s): ", book.sheet_names()[0]
    print 'book.nsheets', book.nsheets
    print 'sh.name:', sh.name, 'sh.nrows:', sh.nrows, 'sh.ncols:', sh.ncols
    print 'A1:', sh.cell_value(rowx=0, colx=1)
    #print 'xlrd: ', time.time() - start
    #sh = book.sheet_by_index(0)
    #print sh.name, sh.nrows, sh.ncols
    #print "Cell D30 is", sh.cell_value(rowx=29, colx=3)
    #for rx in range(sh.nrows):
    #    print sh.row(rx)

#test how to use xlwt
def test_xlwt(filename):
    book = xlwt.Workbook()
    sheet1 = book.add_sheet('hello')
    book.add_sheet('world')
    sheet1.write(0, 0, 'hello')
    sheet1.write(0, 1, 'world')
    row1 = sheet1.row(1)
    row1.write(0, 'A2')
    row1.write(1, 'B2')
    
    sheet1.col(0).width = 10000
    
    sheet2 = book.get_sheet(1)
    sheet2.row(0).write(0, 'Sheet 2 A1')
    sheet2.row(0).write(1, 'Sheet 2 B1')
    sheet2.flush_row_data()
    
    sheet2.write(1, 0, 'Sheet 2 A3')
    sheet2.col(0).width = 5000
    sheet2.col(0).hidden = True
    
    book.save(filename)
    #book.save(TemporaryFile())

#read name and his/her department (name_deps)、weekdays which is holiday (weekday)
#and weekends which is workday (weejebd)
def read_name_deps(file):
    name_deps = {}
    book = xlrd.open_workbook(file)
    sh1 = book.sheet_by_index(0)
    nrows = sh1.nrows
    for row in range(nrows):
        dep = sh1.cell_value(row, 0)
        name = sh1.cell_value(row, 1)
        name_deps[name] = dep
        #print row, dep.encode('utf - 8'), name.encode('utf - 8')
    weekday = []
    weekend = []
    sh2 = book.sheet_by_index(1)
    nrows = sh2.nrows
    #print nrows
    for row in range(nrows):
        date1 = sh2.cell_value(row, 0)
        date2 = sh2.cell_value(row, 1)
        #print 'type(date1):', type(date1), 
        #print 'type(date2):', type(date2)
        if type(date1) ==  float:
            weekday.append(date1)
            #print 'weekday:', date1
        if type(date2) ==  float:
            weekend.append(date2)
            #print 'weekend:', date2

    return name_deps, weekday, weekend

#abstract names from name_deps
def get_names(name_deps):
    names = {}
    for item in name_deps:
        #print type(item)
        #item = item.decode('gb2312')
        #print type(item)
        names[item] = False
    return names

#return workdays in a month
def get_workday(year, month, weekday, weekend):
    #calculate the start and end date of the workday
    #any time in excel is started from 1899.12.31
    base = date(1899, 12, 31).toordinal()
    #every month, the start day is last month, 21st
    start = date(year, month-1, 21).toordinal() - base + 1
    #and the end day is this month, 20th
    end = date(year, month, 20).toordinal() - base + 1
    
    workday = []
    for i in range(start, end+1):
        tmp = date.fromordinal(i+base-1).weekday()
        #print i, tmp
        #add week days which need to go to work
        if tmp in [0, 1, 2, 3, 4] and i not in weekday:
            workday.append(i)
    
    #add weekend days which need to go to work in the treated period
    period_weekend = []
    for i in weekend:
        if i <= end and i >= start:
            period_weekend.append(i)
    
    for i in period_weekend:
        workday.append(i)
    
    return sorted(workday)      

#if employees miss a workday, the function will add the record of the day
def add_workday(workday, name, base_name, workday_cnt, day, shw, cnt): 
    #print 'name:', name.encode('gb2312'), 'base_name:', base_name.encode('gb2312'), 
    #print type(day), 'workday[workday_cnt]:', workday[workday_cnt]
    if name == base_name:
        #print 'type(day)', type(day)
        if type(day) is float:
            #print 'day:', day, 'workday_cnt:', workday_cnt, 'workday[workday_cnt]:', workday[workday_cnt]
            #if employee 'name' miss some workday
            #print 'len(workday):', len(workday), ' workday_cnt:', workday_cnt, 
            #print 'day:', int(day), 'workday[workday_cnt]:', workday[workday_cnt]
            
            #if the 19th and 20th are holidays but the 'name' go to work
            if workday_cnt >= len(workday):
                pass
            #if employee 'name' go to work on holiday
            elif int(day) < workday[workday_cnt]:
                #print 'in  < '
                pass
            #if the 'name' miss some workdays
            elif int(day) > workday[workday_cnt]:
                #shw.write(cnt, DEPCOLX, name_deps[name])
                #shw.write(cnt, NAMECOLX, name)
                shw.write(cnt, WORKDAYCOLX, workday[workday_cnt])
                shw.write(cnt, WORKDAYCOLX+1, u'无打卡记录')
                shw.write(cnt, WORKDAYCOLX+2, u'无打卡记录')
                shw.write(cnt, WORKDAYCOLX+3, u'无打卡记录')
                shw.write(cnt, WORKDAYCOLX+4, u'无打卡记录')
                shw.write(cnt, WORKDAYCOLX+5, u'无打卡记录')
                shw.write(cnt, WORKDAYCOLX+6, u'无打卡记录')
                shw.write(cnt, WORKDAYCOLX+7, u'是')
                cnt += 1
                workday_cnt += 1
                shw.write(cnt, DEPCOLX, name_deps[name])
                shw.write(cnt, NAMECOLX, name)
                workday_cnt, cnt, base_name = add_workday(workday, name, base_name, workday_cnt, day, shw, cnt)
            #if employee 'name' go to work on the 'day'
            else:
                workday_cnt += 1
        #to check weather the employee 'name' havn't go to work at the end of the month
        elif type(day) is str and workday_cnt! = 0:
            dayCnt = len(workday)
            while workday_cnt! = dayCnt:
                shw.write(cnt, WORKDAYCOLX, workday[workday_cnt])
                shw.write(cnt, WORKDAYCOLX+1, u'无打卡记录')
                shw.write(cnt, WORKDAYCOLX+2, u'无打卡记录')
                shw.write(cnt, WORKDAYCOLX+3, u'无打卡记录')
                shw.write(cnt, WORKDAYCOLX+4, u'无打卡记录')
                shw.write(cnt, WORKDAYCOLX+5, u'无打卡记录')
                shw.write(cnt, WORKDAYCOLX+6, u'无打卡记录')
                shw.write(cnt, WORKDAYCOLX+7, u'是')
                cnt +=  1
                workday_cnt +=  1
                shw.write(cnt, DEPCOLX, name_deps[name])
                shw.write(cnt, NAMECOLX, name)
            print name.encode('gb2312'), workday_cnt
                
    else:
        base_name = name
        workday_cnt = 0
    return workday_cnt, cnt, base_name
    
    
#pick employees from the original tables and records their attendance to a new table
def pick_member(file_read, file_write, name_deps, workday):
    book = xlrd.open_workbook(file_read)
    shr = book.sheet_by_index(0)
    newBook = xlwt.Workbook()
    shw = newBook.add_sheet('sheet1')
    #print 'rows: ', shr.nrows
    #print 'columns: ', shr.ncols
    #print 'A15: ', shr.cell_value(rowx = 14, colx = 0).encode('gbk')
    names = get_names(name_deps)
    
    #base_name = shr.cell_value(0, NAMECOLX)
    base_name = False
    workday_cnt = 0
    
    cnt = 0
    for rx in range(shr.nrows):
        #print 'A'+str(rx)+': ', sh.cell_value(rowx = rx, colx = DEPCOLX).encode('gbk')
        name = shr.cell_value(rx, NAMECOLX)

        if name in names.keys():
            if not base_name:
                base_name = shr.cell_value(rx, NAMECOLX)
            
            names[name] = True
            shw.write(cnt, DEPCOLX, name_deps[name])
            shw.write(cnt, NAMECOLX, name)
            
            day = shr.cell_value(rx, WORKDAYCOLX)
            
            #To check whether it needs to add a row
            #if type(day) is not unicode and workday_cnt! = 22:
            #    print 'rx:', rx, 'name:', name.encode('gb2312'), ' base_name:', base_name.encode('gb2312'), 
            #    print 'day:', day, 'workday_cnt:', workday_cnt, 'workday:', workday[workday_cnt]
            workday_cnt, cnt, base_name = add_workday(workday, name, base_name, workday_cnt, day, shw, cnt)
             
            for cx in range(2, 5):
                shw.write(cnt, cx, shr.cell_value(rx, cx))
            first = shr.cell_value(rx, cx - 1)
            last = shr.cell_value(rx, cx)
            #print type(hours)
            flag = False
            if type(first) is unicode:
                shw.write(cnt, cx + 1, u'上班考勤')
                shw.write(cnt, cx + 2, u'下班考勤')
                shw.write(cnt, cx + 3, u'考勤工时')
                shw.write(cnt, cx + 4, u'是否工时不足')
                shw.write(cnt, cx + 5, u'是否有考勤异常') 
            elif type(first) is float:
                if first < tpnts[0]:
                    hours = (last - tpnts[0])*24
                else:
                    hours = (last - first)*24
                shw.write(cnt, cx + 3, hours)
                date = shr.cell_value(rx, cx - 2)
                if date in workday:
                    #10:31 - 15:59叫旷工
                    if first - tpnts[2] > 0:
                        #print 'KG'
                        #print 'first:', first*24
                        #print '10:30:', tpnts[2]
                        shw.write(cnt, cx + 1, u'***旷工***')
                        flag = True
                    #10:01至10:30之间叫迟到
                    elif first - tpnts[1] > 0:
                        #print 'CD'
                        #print 'first:', first*24, ' 10:00', tpnts[1]
                        shw.write(cnt, cx + 1, u'***迟到***')
                        flag = True
                    else:
                        #print 'first ZH', first*24
                        shw.write(cnt, cx + 1, u'正常')
                
                    #10:31 - 15:59叫旷工
                    if last - tpnts[3] < 0:
                        #print 'KG'
                        #print 'last:', last*24
                        #print '16:00:', tpnts[3]
                        shw.write(cnt, cx + 2, u'***旷工***')
                        flag = True
                    #16：00至16：29之间叫早退
                    elif last - tpnts[4] < 0:
                        #print 'ZT'
                        #print 'last:', last, ' 16:30, ', tpnts[4]
                        shw.write(cnt, cx + 2, u'***早退***')
                        flag = True
                    else:
                        #print 'last ZH', last*24
                        shw.write(cnt, cx + 2, u'正常')
                    #小时8.5小时叫工时不足
                    if hours < 8.5:
                        #print 'hours: ', hours
                        shw.write(cnt, cx + 4, u'工时不足')
                        flag = True
                    else:
                        #print 'hours: ', hours
                        shw.write(cnt, cx + 4, u'正常')
                else:
                    shw.write(cnt, cx + 1, u'非工作日')
                    shw.write(cnt, cx + 2, u'非工作日')
                    shw.write(cnt, cx + 4, u'非工作日')
                    #shw.write(cnt, cx + 5, u'加班')
            
                if flag:
                    shw.write(cnt, cx + 5, u'是')
                else:
                    shw.write(cnt, cx + 5, u'否')
            else:
                pass
            cnt += 1
            
    for name in names:
        if not names[name]:
            shw.write(cnt, DEPCOLX, name_deps[name])
            shw.write(cnt, NAMECOLX, name)
            shw.write(cnt, WORKDAYCOLX, u'原始表中无记录')
            shw.write(cnt, WORKDAYCOLX + 1, u'原始表中无记录')
            shw.write(cnt, WORKDAYCOLX + 2, u'原始表中无记录')
            shw.write(cnt, WORKDAYCOLX + 3, u'原始表中无记录')
            shw.write(cnt, WORKDAYCOLX + 4, u'原始表中无记录')
            shw.write(cnt, WORKDAYCOLX + 5, u'原始表中无记录')
            shw.write(cnt, WORKDAYCOLX + 6, u'原始表中无记录')
            shw.write(cnt, MARKCOLX, u'是')
            
            cnt +=  1
    newBook.save(file_write)
                        
    
if __name__ == '__main__':
    #test_xlrd('June.xls')
    #test_xlwt('xlwt.xls')
    
    #{name:dep}, [weekday which is holiday], [weekend which is workday]
    name_deps, weekday, weekend = read_name_deps(u'处理考勤所需信息.xls')
    '''
    for item in name_deps:
        print item.encode('gb2312'), name_deps[item].encode('gb2312')
    
    print weekday
    print weekend
    '''
    
    '''
    origin_attandence_file = sys.argv[1]
    new_attandence_file = sys.argv[2]
    year = int(sys.argv[3])
    month = int(sys.argv[4])
    '''
    #'''
    origin_attandence_file = 'april.xls'
    new_attandence_file = 'aprilNew.xls'
    year = 2013
    month = 4
    #'''
    #wordays of June
    workday = get_workday(year, month, weekday, weekend)
    #print workday
    
    
    pick_member(origin_attandence_file, new_attandence_file, name_deps, workday)
    test_xlrd('Book1.xls')

