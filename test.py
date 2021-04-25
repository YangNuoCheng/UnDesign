# -*- coding: utf-8 -*-
"""
Created on Wed Apr 21 16:52:03 2021

@author: 18801

重新写匹配博弈算法
"""
import xlrd
import xlwt
import numpy as np
import random
import pandas as pd
from xlutils.copy import copy
window = 10
# 所有用户、主机标号都从0开始

def readData():
    work_book = xlrd.open_workbook('compare.xlsx')
    sheet_1 = work_book.sheet_by_index(0)    #每个用户的属性
    sheet_2 = work_book.sheet_by_index(1)    #每个服务器的属性
    usersnumber = sheet_1.nrows-1
    servernumber = sheet_2.nrows-1
    # print(sheet_1.nrows)
    users = np.empty(shape=(usersnumber,14))
    for i in range(1,usersnumber+1):
        for j in range(0,14):
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
def WriteBackToxlsx(user):
    Workbook = xlrd.open_workbook('compare.xlsx')
    sheet_1 = Workbook.sheet_by_index(0)    #每个用户的属性
    usersnumber = sheet_1.nrows-1
    booknames=Workbook.sheet_names()                   # 以列表的形式返回
    new_workbook = copy(Workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格,用户表格
    # print(type(new_worksheet))
    for i in range(1,usersnumber):
        new_worksheet.write(i,11,user[i][11])
        new_worksheet.write(i,12,user[i][12])
    new_workbook.save('compare.xlsx')  # 保存工作簿
    # print("xlxs表格sheet1写入数据成功！")
    print("结果重新写回compare.xlsx")
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

def findOutBestMatchUser(user1,server):
    # 传入的是一个用户和所有的server
    i=0
    maxMatch = -10000
    matchNumber = -1
    for j in server:
        if(userUtility(user1,j)>maxMatch and j[11]>user1[10]):
            maxMatch=userUtility(user1,j)
            matchNumber = i
        i = i+1
    return matchNumber

def findOutBestMatchSer(user,server,servernumber):
    # 传入所有用户和一个服务器,此服务器的标号；返回一个连接序列，一个剩余算力（然后在main中修改user的状态，循环传入
    # 读取user中保存的待连接目标j[11]，验证j[13]的连接状态
    # 函数要做的事情:做好排序，匹配窗口，修改算力
    prepareNumberList = []  #用户序列表
    prepareUtility = []  #用户序列对应的效用表
    connectList = []  #要建立连接的用户表
    count = 0  
    NewConnect = False  #是否有新的连接建立
    for j in user:
        if(j[13] == 0 and j[11] == servernumber and server[11] >= j[10]):  #问题出在这里
            prepareNumberList.append(count)
            NewConnect = True
        count = count + 1
    if(NewConnect):
        for i in range(np.shape(prepareNumberList)[0]):
            # print(i)
            prepareUtility.append(serverUtility(user[i],server))
        # array = ((np.array(prepareUtility,dtype = float)).argsort()).tolist().reverse() 
        array = ((np.array(prepareUtility,dtype = float)).argsort()).tolist()
        array.reverse() 
        #输出了从大到小的用户索引值
        times = min(window,np.shape(prepareNumberList)[0])
        for i in range(times):
            connectList.append(prepareNumberList[array[i]])
            server[11] = server[11] - user[prepareNumberList[array[i]]][10]
    return [connectList,server[11]]

def not_empty(setA):
    return bool(setA)
    
def main():
    user,server=readData()
    connectEDlist = []
    userList = []
    for i in range(np.shape(user)[0]):
        userList.append(i)
    loop = 0
    while(userList != connectEDlist and loop < 100):
        for i in user:
            if(i[13] == 0):
                i[11] = findOutBestMatchUser(i,server)
        count = 0  #服务器标号
        for i in server:
            NewConnect,power = findOutBestMatchSer(user,i,count)  #NewConnect应当是要加入此服务器的用户编号
            count = count + 1
            i[11] = power
            for j in NewConnect:
                user[j][13] = 1
                connectEDlist.append(j)
        loop = loop + 1

    for i in range(np.shape(user)[0]):
        if(int(user[i][11]) != -1):
            print("经过",loop,"轮的匹配,用户",i,"匹配到了",int(user[i][11]))
    count = 0
    for i in user:
        if(i[13] == 1):
            count = count + 1
    rest = set(userList).difference(set(connectEDlist))
    # print("全体用户userList",userList)
    print(count,"个用户成功匹配","connectEDlist",connectEDlist)
    print("未匹配到的用户有",rest)
    print("*************")
    # ------------随机匹配部分
    user2,server2=readData()
    count = 0
    for i in server2:
        for j in user:
            if(i[11]>j[10] and j[12]==10000):
                j[12] = count
                # print(count)
                i[11] = i[11]-j[10]
        count = count + 1
    count = 0
    for i in user:
        print("用户",count+1,"匹配到了",int(i[12]))
        count = count + 1
    # ------------随机匹配部分
    WriteBackToxlsx(user)
if __name__ == "__main__":
    main()