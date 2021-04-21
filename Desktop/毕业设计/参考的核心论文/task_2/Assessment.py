# -*- coding: utf-8 -*-
"""
Created on Sun Apr 18 17:18:13 2021

@author: 18801

本方法主要用于评估匹配的结果，和随机匹配相比较
"""
import xlrd
import xlwt
import numpy as np
import random
import pandas as pd
from xlutils.copy import copy

def readData():
    work_book = xlrd.open_workbook('compare.xlsx')
    sheet_1 = work_book.sheet_by_index(0)    #每个用户的属性
    sheet_2 = work_book.sheet_by_index(1)    #每个服务器的属性
    usersnumber = sheet_1.nrows-1
    servernumber = sheet_2.nrows-1
    # print(sheet_1.nrows)
    users = np.empty(shape=(usersnumber,13))
    for i in range(1,usersnumber+1):
        for j in range(0,13):
            # print(sheet_1.cell_value(i,j))
            # print(j)
            users[i-1,j]=sheet_1.cell_value(i,j)
            # print(sheet_1.cell_value(i,j))
            
    servers = np.empty(shape=(servernumber,14))
    for i in range(1,servernumber+1):
        for j in range(0,14):
            servers[i-1,j]=sheet_2.cell_value(i,j)
            # print(sheet_2.cell_value(i,j))
    # user k1时延	k2价格	k3可靠性	k4故障率	k5私密性	β1分享系数	β2	β3	β4	β5	任务量r	连接状态
    # server e1努力程度	e2	e3	e4	e5	b1努力成本	b2	b3	b4	b5	ρ	m	σ是一个随机量（0-0.5）
    return [users,servers]

def userUtility(user,server):
    # 传入的是一维数组，计算某个用户连接某个主机时的效用函数
    seita = random.normalvariate(0, server[12])
    Utility = 0
    for i in range(0,6):
        Utility = Utility + ((1-user[i+5])*(user[i]*server[i]+seita))
    # print(Utility)
    return Utility

def serverUtility(user,server):
    # 传入的是一维数组，计算某个主机连接某个用户时的效用函数
    seita = random.normalvariate(0, server[12])
    Utility = 0;
    for i in range(0,6):
        Utility = Utility + (user[i+5]*(user[i]*server[i]+seita)-0.5*(server[5+i])*(server[i]*server[i])-0.5*server[10]*(user[i+5]*user[i+5])*(server[12]*server[12]))
    # print(Utility)
    return Utility

def avgSatiUsers(users,servers,strA):
    # 传入所有用户和所有服务器列表
    # strA = 'matching'则是匹配博弈的结果,否则为随机分配结果
    pos = 12
    if(strA == 'matching'):
        pos = 11
    avgSatUser = 0
    sumSatUser = 0
    avgSatServer = 0
    sumSatServer = 0
    count = 0
    for i in users:
        sumSatUser = sumSatUser + userUtility(i,servers[int(i[pos])])
        sumSatServer = sumSatServer + serverUtility(i,servers[int(i[pos])])
        count = count + 1
    avgSatUser = sumSatUser / count
    avgSatServer = sumSatServer / count
    return [avgSatUser,avgSatServer]
    
def main():
    users,servers = readData()
    print(avgSatiUsers(users,servers,'matching'))
    print(avgSatiUsers(users,servers,'random'))
    
if __name__ == "__main__":
    main()
