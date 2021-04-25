# -*- coding: utf-8 -*-
"""
Created on Sat Apr 17 16:15:43 2021

@author: 18801

这个文件用于随机生成用户数据和服务器的数据
"""
import random
import xlrd
import xlwt
from xlutils.copy import copy
import pandas as pd

NumberofUsers = 1000
NumberofServers = 20

def drop():
    book = xlrd.open_workbook('compare.xlsx')
    sheet = book.sheet_by_name('Sheet1')
    cols = sheet.ncols
    rows = sheet.nrows
    names = []
    for i in range(cols):
        names.append(sheet.cell_value(0,i))
    print(names)
    
    Workbook = copy(book)
    # book.sheet_by_index(0)
    # sheet_1 = Workbook.sheet_by_index(0)
    sheet = Workbook.get_sheet(0)
    for j in range(rows):
        for i in range(cols):
            # sheet.write(0,i,names[i])
            sheet.write(j+1,i,'')
    Workbook.save('compare.xlsx')
    
    sheet = book.sheet_by_name('Sheet2')
    cols = sheet.ncols
    rows = sheet.nrows
    names = []
    for i in range(cols):
        names.append(sheet.cell_value(0,i))
    print(names)
    
    Workbook = copy(book)
    sheet = Workbook.get_sheet(1)
    for j in range(rows):
        for i in range(cols):
            # sheet.write(0,i,names[i])
            sheet.write(j+1,i,'')
    Workbook.save('compare.xlsx')
    print("compare.xlsx数据清零")
    
def write():
    Workbook = xlrd.open_workbook('compare.xlsx')
    booknames=Workbook.sheet_names()                   # 以列表的形式返回
    new_workbook = copy(Workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(1,NumberofUsers):
        for j in range(0,10):
            new_worksheet.write(i, j, random.uniform(0,0.5))
        new_worksheet.write(i,10, random.randint(100,200))
        new_worksheet.write(i,11, 0)
        new_worksheet.write(i,12, 10000)
        new_worksheet.write(i,13, 0)
    new_workbook.save('compare.xlsx')  # 保存工作簿
    print("xlxs表格sheet1写入数据成功！")
    
    Workbook = xlrd.open_workbook('compare.xlsx')
    new_workbook = copy(Workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(1)  # 获取转化后工作簿中的第一个表格
    for i in range(1,NumberofServers):
        for j in range(0,10):
            new_worksheet.write(i, j, random.uniform(0,1))
        new_worksheet.write(i,10, random.uniform(0,1))
        new_worksheet.write(i,11, random.randint(200,1000))
        new_worksheet.write(i,12, random.uniform(0,0.2))
        new_worksheet.write(i,13, 0)
    new_workbook.save('compare.xlsx')  # 保存工作簿
    print("xlxs表格sheet2写入数据成功！")

def main():
    drop()
    write()
    

if __name__ == "__main__":
    main()