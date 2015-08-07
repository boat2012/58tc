#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlwt,xlrd
from xlutils.copy import copy

def r_(fn1):    # 比较两个excel文件中的纪录，将相同的行合并，并进行价格比较
    price1=[]
    ws1 = xlrd.open_workbook(fn1).sheet_by_index(0)
    for i in range(ws1.nrows):
        price1.append(ws1.row(i))
    print len(price1)

if __name__ == '__main__':
    price_com("test0730.xls","test0729.xls")