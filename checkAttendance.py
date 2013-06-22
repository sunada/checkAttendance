#coding:utf-8

import xlrd
import xlwt
#from openpyxl import load_workbook
import time
from datetime import date,datetime

#the colx of name
nameColx=1
#import time point: 8:00,10:00,10:30,16:00,16:30,19:00,20:00
tpnts=[0.333333,0.416667,0.4375,0.666666,0.6875,0.791667,0.833333]

#xlrd is much more efficent
def testXlrd(filename):
    start=time.time()
    #, encoding_override="gbk"
    book=xlrd.open_workbook(filename)
    sh=book.sheet_by_index(0)
    print "Worksheet name(s): ",book.sheet_names()[0]
    #print 'xlrd: ', time.time()-start
    print book.biff_version
    print book.codepage
    print book.encoding
    #sh=book.sheet_by_index(0)
    #print sh.name, sh.nrows, sh.ncols
    #print "Cell D30 is", sh.cell_value(rowx=29, colx=3)
    #for rx in range(sh.nrows):
    #    print sh.row(rx)
    
def testXlwt(filename):
    book=xlwt.Workbook()
    sheet1=book.add_sheet('hello')
    book.add_sheet('world')
    sheet1.write(0,0,'hello')
    sheet1.write(0,1,'world')
    row1 = sheet1.row(1)
    row1.write(0,'A2')
    row1.write(1,'B2')
    
    sheet1.col(0).width = 10000
    
    sheet2 = book.get_sheet(1)
    sheet2.row(0).write(0,'Sheet 2 A1')
    sheet2.row(0).write(1,'Sheet 2 B1')
    sheet2.flush_row_data()
    
    sheet2.write(1,0,'Sheet 2 A3')
    sheet2.col(0).width = 5000
    sheet2.col(0).hidden = True
    
    book.save(filename)
    #book.save(TemporaryFile())

def ReadNamedeps(file):
    f=open(file)
    lines=f.readlines()
    nameDeps={}
    dep=None
    name=None
    for line in lines:
        line=line.rstrip()
        if not line:
            continue
        if line[:4]=='dep:':
            dep=line[4:].decode('gb2312')
        else:
            name=line.decode('gb2312')
            #print dep, name
            nameDeps[name]=dep
    return nameDeps

def getNames(nameDeps):
    names=[]
    for item in nameDeps:
        #print type(item)
        #item=item.decode('gb2312')
        #print type(item)
        names.append(item)
    return names

def pickMember(fileRead,fileWrite,nameDeps):
    book=xlrd.open_workbook(fileRead)
    shr=book.sheet_by_index(0)
    newBook=xlwt.Workbook()
    shw=newBook.add_sheet('sheet1')
    #print 'rows: ', shr.nrows
    #print 'columns: ', shr.ncols
    #print 'A15: ', shr.cell_value(rowx=14,colx=0).encode('gbk')
    names=getNames(nameDeps)
        
    cnt=0
    for rx in range(shr.nrows):
        #print 'A'+str(rx)+': ', sh.cell_value(rowx=rx,colx=depColx).encode('gbk')
        name=shr.cell_value(rx,nameColx)
        
        if name in names:
            shw.write(cnt,0,nameDeps[name])
            for cx in range(1,shr.ncols-3):
                shw.write(cnt,cx,shr.cell_value(rx,cx))
            first=shr.cell_value(rx,cx-1)
            last=shr.cell_value(rx,cx)
            #print type(hours)
            flag=False
            if type(first) is unicode:
                shw.write(cnt,cx+1,u'上班考勤')
                shw.write(cnt,cx+2,u'下班考勤')
                shw.write(cnt,cx+3,u'考勤工时')
                shw.write(cnt,cx+4,u'是否工时不足')
                shw.write(cnt,cx+5,u'是否有考勤异常') 
            elif type(first) is float:
                if first<tpnts[0]:
                    hours=(last-tpnts[0])*24
                else:
                    hours=(last-first)*24
                shw.write(cnt,cx+3,hours)
                #10:31-15:59叫旷工
                if first-tpnts[2]>0:
                    #print 'KG'
                    #print 'first:',first*24
                    #print '10:30:',tpnts[2]
                    shw.write(cnt,cx+1,u'***旷工***')
                    flag=True
                #10:01至10:30之间叫迟到
                elif first-tpnts[1]>0:
                    #print 'CD'
                    #print 'first:',first*24,' 10:00',tpnts[1]
                    shw.write(cnt,cx+1,u'***迟到***')
                    flag=True
                else:
                    #print 'first ZH',first*24
                    shw.write(cnt,cx+1,u'正常')
                
                #10:31-15:59叫旷工
                if last-tpnts[3]<0:
                    #print 'KG'
                    #print 'last:',last*24
                    #print '16:00:',tpnts[3]
                    shw.write(cnt,cx+2,u'***旷工***')
                    flag=True
                #16：00至16：29之间叫早退
                elif last-tpnts[4]<0:
                    #print 'ZT'
                    #print 'last:',last,' 16:30,',tpnts[4]
                    shw.write(cnt,cx+2,u'***早退***')
                    flag=True
                else:
                    #print 'last ZH',last*24
                    shw.write(cnt,cx+2,u'正常')
                #小时8.5小时叫工时不足
                if hours<8.5:
                    #print 'hours: ',hours
                    shw.write(cnt,cx+4,u'工时不足')
                    flag=True
                else:
                    #print 'hours: ',hours
                    shw.write(cnt,cx+4,u'正常')
            
                if flag:
                    shw.write(cnt,cx+5,u'是')
                else:
                    shw.write(cnt,cx+5,u'否')
            
            else:
                pass
            
            
            cnt+=1
    newBook.save(fileWrite)
    
                
                
if __name__=='__main__':
    #testXlrd('June.xls')
    #testOpenpyxl('June.xlsx')
    #testXlwt('xlwt.xls')
    
    #our employees' names
    nameDeps=ReadNamedeps(u'部门人员信息.txt')  
    #for item in nameDeps:
    #    print item
    names=getNames(nameDeps)

    pickMember('June.xls','treated1.xls',nameDeps)


