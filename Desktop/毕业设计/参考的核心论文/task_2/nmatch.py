# -*- coding: utf-8 -*-
"""
Created on Sat Apr 17 17:24:13 2021

@author: 18801

此文件进行多轮的匹配，使每个用户都可以得到服务，得到最终的解集
"""
import xlrd
import xlwt
import numpy as np
import random
import pandas as pd
from xlutils.copy import copy
window = 1
#每次匹配的机器数为window个，每次只能添加一个用户到主机列表中
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
def WriteBackToxlsx(user,server,serverList,user2):
    Workbook = xlrd.open_workbook('compare.xlsx')
    sheet_1 = Workbook.sheet_by_index(0)    #每个用户的属性
    sheet_2 = Workbook.sheet_by_index(1)    #每个服务器的属性
    usersnumber = sheet_1.nrows-1
    servernumber = sheet_2.nrows-1
    
    booknames=Workbook.sheet_names()                   # 以列表的形式返回
    new_workbook = copy(Workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格,用户表格
    for i in range(1,usersnumber):
        new_worksheet.write(i,11,user[i][11])
        new_worksheet.write(i,12,user2[i][12])
    new_workbook.save('compare.xlsx')  # 保存工作簿
    # print("xlxs表格sheet1写入数据成功！")
    
    Workbook = xlrd.open_workbook('compare.xlsx')
    new_workbook = copy(Workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(1)  # 获取转化后工作簿中的第二个表格，服务器表格
    for i in range(1,servernumber):
        # print(serverList[i])
        count = 0
        for j in serverList[i]:
            new_worksheet.write(i,13+count,j)
            count = count + 1
    new_workbook.save('compare.xlsx')  # 保存工作簿
    # print("xlxs表格sheet2写入数据成功！")
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
            # 只按照用户的效用排序，完全没有管服务器对用户的效用排序
            maxMatch=userUtility(user1,j)
            matchNumber = i
        i = i+1
    # print(maxMatch,matchNumber)
    return matchNumber
        
def findOutBestMatchSer(user,server,servernumber):
    # 传入的是一个server和所有的用户,回传的是server连接的第一个用户
    ans = [] #用户的列表
    count = 0
    for j in user:
        if(j[13] == 0):  #用户没有确定连接
            # print(j[11])
            if(j[11] == servernumber and server[11]>j[10]): #算力大小匹配且用户目标主机为本主机
                ans.append(count)
                # server[11]=server[11]-j[10]  # 可以先不处理，记录好所有的序列，之后有选择的进行匹配
        count = count + 1
    # 需要对ans中的服务器效用进行排序
    listSort = ans
    # print("-----",ans,"-----")
    # for i in range(len(ans)):
    #     listSort[i] = serverUtility(user[i],server)
    # print(listSort)
    # 循环window次，只连接前window个用户
    return ans
    # 设计思想，回传可以连接的用户编号（数组），第一轮，满不满另说

def not_empty(setA):
    return bool(setA)

def Courrent(listA):
    CurrentList=[]
    for i in listA:
         for j in i:
             CurrentList.append(j)
    CurrentList.sort()
    return CurrentList

def main():
    user,server=readData()
    # print(userUtility(user[1], server[1]))
    # print(serverUtility(user[1], server[1]))
    count = 0
    for i in user:
        user[count][11] = findOutBestMatchUser(i,server) #用户待连接的主机已经赋值
        # print(count,i[11],i[10])
        count = count + 1
    NumberOfUsers = count
    StandardList = []
    for i in range(0,count):
        StandardList.append(i)
    # print(StandardList)

    # -------------服务器匹配用户
    rest= set(StandardList)
    count = 0
    serverList=[]
    for i in server:
        # print("*****",findOutBestMatchSer(user,i,count),"******")
        serverList.append(findOutBestMatchSer(user,i,count))
        count=count+1
    count = 1
    # print(serverList)
    # 演示，每个主机匹配到的用户，以及剩余的算力
    print("第一次匹配结果")
    for i in serverList:
        print(count,i,server[count-1][11])
        count = count + 1
    CurrentList=Courrent(serverList)
    # print(CurrentList) 
    rest = set(StandardList).difference(set(CurrentList))
    # 此时还有很多用户没有能够匹配到自己心中最佳的服务器（没有出现在
    # ------------服务器匹配用户
    print("没有匹配的用户",rest)
    # loop start
    loopNumber = 0
    while(not_empty(rest)):
        loopNumber = loopNumber + 1
        for i in rest:
            user[i][11] = findOutBestMatchUser(user[i],server)
        for i in rest:
            if(server[int(user[i][11])][11]>user[i][10]):
                serverList[int(user[i][11])].append(i)
                server[int(user[i][11])][11] = server[int(user[i][11])][11] - user[i][10]
                print("用户",i,"匹配到了服务器",int(user[i][11]))
        CurrentList=Courrent(serverList)
        rest = set(StandardList).difference(set(CurrentList))
        if(loopNumber > 100):
            break
        # print(rest)
    # 演示，每个主机匹配到的用户，以及剩余的算力
    if(loopNumber != 0):
        print("经过",loopNumber+1,"次的匹配,最终结果为:")
        count = 0
        for i in serverList:
            print(count+1,i,server[count][11])
            count = count + 1
    rest = set(StandardList).difference(set(CurrentList))
    if(not_empty(rest)):
        print("最终依然没有得到匹配的是",rest)
    else:
        print("全部得到了匹配")
    print("-------------------------------")
    print("经过随机的分配算法结果为")
    #----------------随机匹配模型，从上到下的匹配，改user文件的12列
    user2,server2=readData()
    count = 1
    for i in server2:
        # print(i[[11]])
        for j in user2:
            if(i[11]>j[10] and j[12]==10000):
                j[12] = count
                i[11] = i[11]-j[10]
        count = count + 1
        
    count = 0
    for i in user2:
        print("用户",count+1,"匹配到了",int(i[12]))
        count = count + 1
    #----------------随机匹配模型
    # 最后把匹配的结果写回xlxs中
    WriteBackToxlsx(user,server,serverList,user2)


if __name__ == "__main__":
    main()
