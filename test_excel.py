#coding : utf-8
import xlwings as xw

def display_rotated(n):#纵坐标变成数字变成字母，如11是A，27是AA
    ch='A'
    if n >26:
        count=int(n/26)
        return chr(ord(ch) + count)+chr(ord(ch) + n%26)
    else:
        return chr(ord(ch) + n-1)

fopen = open("pm_data", 'r')
lines = fopen.readlines()
fopen.close()
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False
filepath=r'template.xlsx'
wb=app.books.open(filepath)
row=20
column=display_rotated(0)
count=11
for line in lines:
    column=display_rotated(count)
    s=column+str(row)
    wb.sheets['template'].range(s).value = line
    count+=1
wb.save()
wb.close()
app.quit()
input("按任意键结束结束")

