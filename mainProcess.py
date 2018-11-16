# -*- coding: utf-8 -*-
"""
Created on Sun May 27 13:19:50 2018

@author: horsehill
"""

from data import dataManager
from table import formTable
from statistical import analyzer
from splitDeparture import split
import time
import os
from overtime import fromOvertime
import numpy as np 

#import sys

month = 201810
inputfile1 = 'E:/Pyproject/HR/origin8/8月原始数据/简单分析原始记录2018100120181031.xlsx'
inputfile2 = 'E:/Pyproject/HR/origin8/8月原始数据/员工年休假申请情况表.xlsx'
inputfile3 = 'E:/Pyproject/HR/origin8/8月原始数据/2018年10月项目合作人员名单.xlsx'
outputPath = 'E:/Pyproject/HR/output_test2/'

#############################################
if not os.path.exists(outputPath):
    os.makedirs(outputPath) 
#-----------------------gai 企业的名字要出现在两个地方 需要一张外协表
t1 = time.time()
print ('-----------开始啦-----------')
print ('-----------本月为：'+ str(month) +'-----------')
print ('-----------输出文件存放在：'+ outputPath +'-----------')
wideTable, holidayThisMonth = dataManager(inputfile1, inputfile2, inputfile3, month,outputPath).main()
#----------------------创建init一个dataManager对象,并执行main函数（main只是普通命名，非入口函数）
monthDay = np.max(wideTable['day'])
###暂时屏蔽
attendanceRecord, attendanceSummary = formTable(wideTable, holidayThisMonth, outputPath).main()
fromOvertime(wideTable, holidayThisMonth, outputPath).main()

#-----------------------
#lateAllDf, late2HDf, late2HomitDf, lateSummary, lateSummaryDep, \
#extraAllDf, extraBossDf, extraDepartureDf, personExtraSummary, extraSummaryDf, \
#missAllDf, missPersionDf, missDepDf\
###暂时屏蔽
analyzer(wideTable, holidayThisMonth,outputPath).main()
#-----------------------
###暂时屏蔽
split(wideTable,attendanceRecord, attendanceSummary,inputfile1, outputPath, str(month),monthDay).main()
#-----------------------
#-----------------------生成第三个表：月迟到情况统计汇总表
#statistic1=analyzer(self, wideTable, holidayThisMonth, outputPath):
#-----------------------
print ('-----------完成，共耗时 ' + str(int(time.time()-t1))+' 秒-----------')
