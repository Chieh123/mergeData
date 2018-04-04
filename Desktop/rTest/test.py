#!/usr/bin/env python
#-*- coding:utf-8 -*-


import xlrd
import types
import xlwt

#打開xlsx文件
ad_wb = xlrd.open_workbook("bda2018_hw1_table.xlsx")

#獲取第一張表的名稱

row_data1 = ad_wb.sheets()[0] #2gram
row_data2 = ad_wb.sheets()[0] #3gram

#print ("表單數量：", ad_wb.nsheets)

#print ("表單名稱：", ad_wb.sheet_names())
firstR = 3
strC = 1
DFC = 3
TFC = 2
indexC = 0
rate = 0.01
#獲取第一個目標表單

sheet_0 = ad_wb.sheet_by_index(0) #2gram
sheet_1 = ad_wb.sheet_by_index(1) #3gram

#創建xls文件對象
wb = xlwt.Workbook()
#新建表單
sh = wb.add_sheet('merge')
count = 0

#x = "你好嗎"
#y = "你"
#y1 = unicode(y,'utf8')
#x1 = unicode(x,'utf8')
#if x1.find(y1) != -1:
#   print("got it")


print("before first for loop")
for i in range(firstR, 2000):
#  print("in first for loop")
  s = unicode(sheet_0.cell_value(i,strC).encode('utf8'),'utf8')
  df1 = sheet_0.cell_value(i, DFC)
  tf1 = sheet_0.cell_value(i, TFC)
  replace = 0
  for j in range(firstR, sheet_1.nrows):
    s1 = unicode(sheet_1.cell_value(j,strC).encode('utf8'),'utf8')
    df2 = sheet_1.cell_value(j, DFC)
    tf2 = sheet_1.cell_value(j, TFC)
#    print(sheet_0.cell_value(i,strC))
#    print(sheet_1.cell_value(j,strC))
    if s1.find(s) != -1:
#      print("find key word")
      if df1 <= df2*(1+rate) and df1 >= df2*(1-rate):
          replace = 1
          print(replace,count,sheet_1.cell_value(i,strC))
          break
      else:
        count = count + 1
        sh.write(count+firstR,indexC, count+1)
        sh.write(count+firstR,strC, sheet_0.cell_value(i,strC))
        sh.write(count+firstR,DFC, df1-df2)
        sh.write(count+firstR,TFC, tf1-tf2)
        replace = 2
        print(replace,count,sheet_1.cell_value(i,strC))
        break
    else:
        continue
  if replace == 0:
    print(replace,count,sheet_1.cell_value(i,strC))
    count = count + 1
    sh.write(count+firstR,indexC, count+1)
    sh.write(count+firstR,strC, sheet_0.cell_value(i,strC))
    sh.write(count+firstR,DFC, df1)
    sh.write(count+firstR,TFC, tf1)

#for i in range(firstR, sheet_1.nrows):
 # sh.write(count+firstR,indexC, count+1)
 # sh.write(count+firstR,strC, sheet_1.cell_value(i,strC))
 # sh.write(count+firstR,DFC, sheet_1.cell_value(i,DFC))
 # sh.write(count+firstR,TFC, sheet_1.cell_value(i,TFC))
 # count = count + 1

wb.save('merge.xls')

#print (u"表單 %s 共 %d 行 %d 列" % (sheet_0.name, sheet_0.nrows, sheet_0.ncols))

#print("第三列", 12)

#print ("third row and third column:", sheet_0.cell_value(6, 3))
#print(type(s))
#直接輸出日期

#date_value = xlrd.xldate_as_tuple(sheet_0.cell_value(2,2),ad_wb.datemode)

#date1 = xlrd.xldate.xldate_as_datetime(sheet_0.cell_value(2, 2), ad_wb.datemode)

#print (date_value)#元組

#print (date1)#日期
