import xlrd
from openpyxl import load_workbook
import time

#openpyxl will spend too much time
def testOpenpyxl(filename):
    start=time.time()
    wb2=load_workbook(filename)
    print wb2.get_sheet_names() 
    print 'openpyxl: ',time.time()-start
    

#xlrd is much more efficent
def testXlrd(filename):
    start=time.time()
    book=xlrd.open_workbook(filename)
    print "Worksheet name(s): ",book.sheet_names()
    print 'xlrd: ', time.time()-start
    '''
    sh=book.sheet_by_index(0)
    print sh.name, sh.nrows, sh.ncols
    print "Cell D30 is", sh.cell_value(rowx=29, colx=3)
    for rx in range(sh.nrows):
        print sh.row(rx)
    '''

def pickMember(filename):
    book=xlrd.open_workbook(filename)
    sh=book.sheet_by_index(0)
    print 'rows: ', sh.nrows
    print 'columns: ', sh.ncols
    #print 'A15: ', sh.cell_value(rowx=14,colx=0).encode('gbk')
    
    for rx in range(sh.nrows):
        print 'A'+str(rx)+': ', sh.cell_value(rowx=rx,colx=0).encode('gbk')
    
    
if __name__=='__main__':
    #testXlrd('June.xls')
    #testOpenpyxl('June.xlsx')
    pickMember('June.xls')

