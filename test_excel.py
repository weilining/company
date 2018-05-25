#coding : utf-8
import xlwt  # 写入文件
import xlrd  # 打开excel文件

from xlutils.copy import copy




fopen = open("test.txt", 'r')
lines = fopen.readlines()
fopen.close()
w=xlwt.Workbook(encoding='utf-8',style_compression=0)
ws = w.add_sheet('1') #创建一个工作表
i=0
for line in lines:
    li=line.split(' ')
    j=0
    for l in li:
        ws.write(i,j,l)
        j=j+1
    i=i+1
# w =copy('xqtest.xls')
# w.get_sheet(0).write(0,0,"foo")
w.save('xqtest.xls')
