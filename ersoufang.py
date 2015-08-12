#!/usr/bin/env python
# -*- coding: utf-8 -*-

import urllib2,httplib
import os,datetime,string
import chardet
import sys
import re
import time
from bs4 import BeautifulSoup
import sqlite3
import operator
import xlwt,xlrd
from xlutils.copy import copy

httplib.HTTPConnection._http_vsn = 10
httplib.HTTPConnection._http_vsn_str = 'HTTP/1.0'

exceptionStr=["分校","延安","省体","大儒","安华","文博","十八中"]
# 将抓取到的有用信息封装成类
class Info(object):
    def __init__(self,title='',href = '',zj = '',dj = '',mj='',fx = '',putday='',description=''): # 总价，单价，面积，房型
        self.title = title
        self.href = href
        self.zj = int(zj)
        self.dj = dj
        self.mj = mj
        self.fx = fx
        self.putday = putday
        self.description=description
    def get_addr(self):
        return self.addr
    def to_string(self):
#        res =  '<br><a href = \'%s\' >%s</a><br>' %(self.href ,self.title)
#        res = res +self.description + '<br><hr>'
        res = '%s:%s\/n' % (self.title,self.zj)
        return res


##定义抓取网页的类
class tc58(object):
    def __init__(self,url):
        self.__url = url
        self.addr_infos = []
    ##抓取网页到本地
    def load_page(self):
        opener = urllib2.build_opener()
        f = opener.open(self.__url)
        html = f.read()
        self.__html = html
        f.close()
        return html
    ##解析网页元素，找到自己要找的信息，这里要根据实际情况分析标签
    def paser_html(self):

        #print html
        soup = BeautifulSoup(self.__html,'html.parser')
        tables =soup.html.body.find('div',id="main").find('div',id = 'infolist').find_all('table',class_='tbimg')
        table = tables[0].find_all("tr")
        addr_infos = []
        for element in table:
            html = element.find_all('td')
            #print "len (td):%d"%len(html)
            t1 = html[1]
            #raw_input("pause")
            title =  t1.a.get_text()
            #print "title:" + title
            href =  t1.a.get("href")
            #print "href:" + href
            description =  t1.get_text()
            pric_all =  t1.find("div",class_="qj-listright btall")
            putday = t1.find("span",class_="qj-listjjr")
            pric_all = pric_all.get_text()
            title = title.encode("utf-8")
            href = href.encode("utf-8")
            pric_all = pric_all.encode("utf-8")
            pric_all = re.sub(re.compile('\n|\xa0|\xc2')," ",pric_all)
            m = re.search(r'\s*(\d*)\s*万\s*(\d*)元/㎡\s*(\S*)\s*\((\d*.*\d*)㎡\)',pric_all)
            if putday.find_all('a') and m :
               putday = putday.find_all('a')[-1].next_sibling  #发表时间
               putday = re.sub(re.compile('\n|\xa0|\xc2|&nbsp'),"",putday)
               if "今天" in putday or "小时" in putday:
                   putday = time.strftime("%m-%d")
                   zj = m.group(1)
                   dj = m.group(2)
                   fx = m.group(3)
                   mj = m.group(4)
                   description = description.encode("utf-8")
                   description = re.sub(re.compile('\n|\xa0|\xc2')," ",description)
                   i = Info(title,href,zj,dj,fx,mj,putday,description)
                   addr_infos.append(i)
        self.addr_infos = addr_infos


    def send_mail(self):
        pass
    def get_addr_infos(self):
        return self.addr_infos
    def run(self):
        self.load_page()
        self.paser_html()

#这是搜索的基地址
__INITURL__ = 'http://fz.58.com/ershoufang/?PGTID=14373771295430.968437073752284&ClickID=1&key=%s'
def Info2Excel(myInfo,filename):
    if os.path.isfile(filename):
        rbook = xlrd.open_workbook(filename,formatting_info=True)
        wbook = copy(rbook)
        wsheet = wbook.get_sheet(0)
        inserRow = rbook.sheet_by_index(0).nrows
    else:
        wbook = xlwt.Workbook()
        wsheet = wbook.add_sheet("test1")
#        wsheet.write(0,0,u"简介")
#        wsheet.write(0,1,u"发布日期")
#        wsheet.write(0,2,u"总价")
#        wsheet.write(0,3,u"单价")
#        wsheet.write(0,4,u"房型")
#        wsheet.write(0,5,u"面积")
        inserRow = 0
    wsheet.col(0).width = 13000
    wsheet.col(2).width = 3000
    for item in myInfo:
       link = 'HYPERLINK("%s";"%s")'%(item.href.decode('utf-8'),item.title.decode('utf-8'))
       exceTag = True
       for s in exceptionStr:
           if s in item.title.decode('utf-8'):
               exceTag = False
       if len(link) < 255 and exceTag:
          wsheet.write(inserRow,0,xlwt.Formula('HYPERLINK("%s";"%s")'%(item.href.decode('utf-8'),item.title.decode('utf-8'))))
          wsheet.write(inserRow,1,item.putday.decode('utf-8'))
          wsheet.write(inserRow,2,item.zj)
          wsheet.write(inserRow,3,item.dj.decode('utf-8'))
          wsheet.write(inserRow,4,item.fx.decode('utf-8'))
          wsheet.write(inserRow,5,item.mj.decode('utf-8'))
          wsheet.write(inserRow,6,item.href.decode('utf-8'))
          inserRow = inserRow + 1
    print filename
    wsheet.col(6).hidden=True
    wbook.save(filename)


def Info2db(myInfo,dbname):
    conn = sqlite3.connect(dbname)
    conn.execute("create table if not EXISTS fang(title varchar(128) ,href VARCHAR (128),\
                 zj FLOAT ,dj FLOAT ,fx VARCHAR (40),mj FLOAT,putday date)")
    cur = conn.cursor()
    for item in myInfo:
        conn.execute("replace into fang VALUES (\'%s\',\'%s\',\'%s\',\'%s\',\'%s\',\'%s\',\'%s\')" % (item.title,item.href,item.zj,item.dj,item.fx,item.mj,item.putday))

    conn.commit()
    conn.close()

def printdb(dbname):
    conn = sqlite3.connect(dbname)
    cur = conn.cursor()
    cur.execute("select title,zj,putday from fang")
    rows = cur.fetchall()
    for r in rows:
        print r[0],r[1],r[2]
    print len(rows)
    conn.close()

if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf-8')
    key = u'划片钱塘小'
    content = urllib2.quote(key.encode('utf-8'))
    allcont = []
    for i in range(15,0,-1):
        api_url = 'http://fz.58.com/ershoufang/pn%d/?key=%s'%(i,content)
        print api_url
        tc = tc58(api_url)
        tc.run()
    #开始抓取网页
        for item in tc.get_addr_infos():
            allcont.append(item)
    allcont=sorted(allcont,key=operator.attrgetter("zj"))

    newcont=[]
    for i in range(1,len(allcont)):
        if allcont[i].title != allcont[i-1].title:
            newcont.append(allcont[i])
    #sorted(allcont,key=operator.attrgetter("putday"),reverse=True)
    Info2Excel(newcont,"test%s.xls" % time.strftime("%m%d") )
    print len(newcont)

#    Info2db(allcont,"test.db")
#    printdb("test.db")
