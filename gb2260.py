#!/usr/bin/python
# -*- coding: UTF-8 -*- 
import openpyxl

wb = openpyxl.Workbook()
wb = openpyxl.load_workbook('gb2260.xlsx')

# print(wb.sheetnames)

# 打开一个文件
fo = open("gb2260.sql", "w", encoding="utf-8")

dict ={}
for sheet in wb:
    print(sheet.title)
    ws  = wb[sheet.title]
    for rx in range(1,ws.max_row+1):
        w1 = str(ws.cell(rx,1).value).replace(' ','')
        w2 = str(ws.cell(rx,2).value).replace(' ','')

        w = 'insert into gb2260 (revision,code,name) values (\''+ sheet.title+'\',\''+ w1+'\',\''+w2+ '\');\n'
        # print(w)
        fo.write(w)
        dict[w1] = w2

for key in dict:
     print(key,dict[key])  
    # fo.write(ww)
# 关闭打开的文件
fo.close()