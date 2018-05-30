#coding : utf-8
# import xlwt  # 写入文件
# import xlrd  # 打开excel文件
# from xlutils.copy import copy
# import os.path
# file_path='template.xls'
# rb = xlrd.open_workbook(file_path,formatting_info=True)
# r_sheet = rb.sheet_by_index(0)
# wb = copy(rb)
# ws = wb.get_sheet(0)
# fopen = open("test.txt", 'r')
# lines = fopen.readlines()
# fopen.close()
#
# i=0
# for line in lines:
#     li=line.split(' ')
#     j=0
#     for l in li:
#         ws.write(i,j,l)
#         j=j+1
#     i=i+1
# wb.save(file_path)
# input("Tip: press Enter , close window!")
#coding:utf-8
import numpy as np
import xlwings as xw

app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False
filepath=r'template.xlsx'
wb=app.books.open(filepath)
wb.sheets['template'].range('A4').value='韦立宁'
wb.sheets['template'].range('B22').value='0.4'
wb.sheets['template'].range('C22').value='0.8'
wb.sheets['template'].range('D22').value='1.2'
wb.save()
wb.close()
app.quit()
input("结束")

