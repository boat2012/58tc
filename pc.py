#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlwt,xlrd
from xlutils.copy import copy
import time

def readformxls(fn):    # 将excel 文件中的数据读取到一个list中
    price1=[]
    new_price=[]
    wbook = xlrd.open_workbook(fn)
    ws = wbook.sheet_by_index(0)
    for i in range(ws.nrows):
        price1.append(ws.row(i))
    return price1

def writetoxls(list,fn):
    wbook = xlwt.Workbook()
    wsheet = wbook.add_sheet('sheet1')
    for i in range(len(list)):
        wsheet.write(i,0,xlwt.Formula('HYPERLINK("%s";"%s")'%(list[i][6].value,list[i][0].value)))
        wsheet.write(i,1,list[i][1].value.decode('utf-8'))
        wsheet.write(i,2,list[i][2].value)
        wsheet.write(i,3,list[i][3].value.decode('utf-8'))
        wsheet.write(i,4,list[i][4].value.decode('utf-8'))
        wsheet.write(i,5,list[i][5].value)
        wsheet.write(i,6,list[i][6].value.decode('utf-8'))
        for j in range(7,len(list[i]),2):
            if list[i][j+1].value > list[i][2].value:
                cellxf = xlwt.easyxf('font: color green, bold true')
            elif list[i][j+1].value == list[i][2].value:
                cellxf = xlwt.easyxf('font: bold false')
            else:
                cellxf = xlwt.easyxf('font: color red, bold true')
            wsheet.write(i,j,list[i][j].value,cellxf)
            wsheet.write(i,j+1,list[i][j+1].value,cellxf)
    wsheet.col(0).width = 13000
    wsheet.col(2).width = 3000
    wsheet.col(6).hidden = True
    wbook.save(fn)


if __name__ == '__main__':
    list1 = readformxls("test0812.xls")
    list2 = readformxls("test0810.xls")
    for rec in list1:
        for rec2 in list2:            #print rec[0]
            if rec[0].value == rec2[0].value:
                rec.append(rec2[1])
                rec.append(rec2[2])
                find = True
    writetoxls(list1,"comp%s.xls" % time.strftime("%m%d"))

