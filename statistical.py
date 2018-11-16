# -*- coding: utf-8 -*-
"""
Created on Mon May 28 15:21:03 2018

@author: wanghong
"""
import pandas as pd
import numpy as np
from pandas import ExcelWriter

class analyzer():
    
    def __init__(self, wideTable, holidayThisMonth, outputPath):
        wideTable['date'] = wideTable['date'].apply(lambda x: str(x)[:10])
        self.wideTable = wideTable
        self.monthDay = np.max(wideTable['day'])
        self.thisYear = str(list(set(self.wideTable['year']))[0])
        self.thisMonth = str(list(set(self.wideTable['month']))[0]) 
        self.holidayNumber = holidayThisMonth
        self.workDayNumber = self.monthDay - self.holidayNumber 
        self.op = 'E:/Pyproject/HR/output_test2/'
        
    def main(self):
        lateAllDf, late2HDf, late2HomitDf, lateSummary, lateSummaryDep = self.lateAnalyzer()
        lateXlsx = ExcelWriter(self.op+self.thisYear+'年'+self.thisMonth+'月迟到情况统计汇总表.xlsx')
        lateAllDf.columns = ['考勤号码','姓名','部门','日期','时间','签到时间','上班时间','迟到时长','迟到时长（分钟）']
        lateAllDf.to_excel(lateXlsx,'个人明细', index = False)
        late2HDf.columns = ['考勤号码','姓名','部门','日期','时间','签到时间','上班时间','迟到时长','迟到时长（分钟）']
        late2HDf.to_excel(lateXlsx,'迟到2小时以内', index = False)
        late2HomitDf.columns = ['考勤号码','姓名','部门','日期','时间','签到时间','上班时间','迟到时长','迟到时长（分钟）']
        late2HomitDf.to_excel(lateXlsx,'去除免责因素以外迟到', index = False)
        lateSummary.columns = ['考勤号码','姓名','部门','去除免责是否月均迟到超过3次','去除免责迟到次数','是否月均迟到2小时以内超过3次','迟到2小时以内次数']
        lateSummary.to_excel(lateXlsx,'个人汇总', index = False)
        lateSummaryDep.columns = ['部门','人数',self.thisMonth+'月工作日数','应出勤总天数','去除免责总迟到人次',
                                  '去除免责月均迟到率','去除免责月均迟到超过3次人数','去除免责月均迟到超过3次人数占部门人数比例','去除免责月均迟到超过3次人次','去除免责月均迟到超过3次人次占部门迟到人次比例','迟到2小时以内总人次','月均2小时以内迟到率','月均迟到2小时以内超过3次人数','月均迟到2小时以内超过3次人数占部门人数比例','月均迟到2小时以内超过3次人次','月均迟到2小时以内超过3次人次占部门迟到人次比例']
        lateSummaryDep.to_excel(lateXlsx,'部门汇总', index = False)
        lateXlsx.save()
        lateXlsx.close()
        print ('-----------'+self.thisYear+'年'+self.thisMonth+'月迟到情况统计汇总表.xlsx'+' 已生成-----------')
        
#        extraAllDf, extraBossDf, extraDepartureDf, personExtraSummary, extraSummaryDf = self.extraAnalyzer()
#        extraXlsx = ExcelWriter(self.op+self.thisYear+'年'+self.thisMonth+'月加班情况统计汇总表.xlsx')
#        extraSummaryDf.columns = ['总人数','去除各种假日','去除全月不考勤数值','去除全月不考勤占比','17：05以内打卡数值','17：05以内打卡占比',
#                                  '17：10以内打卡数值','17：10以内打卡占比','17：15以内打卡数值','17：15以内打卡占比','平均加班30分钟以上数值','平均加班30分钟以上占比','平均加班60分钟以上数值','平均加班60分钟以上占比']
#        extraSummaryDf.to_excel(extraXlsx,'汇总表', index = False)
#        extraDepartureDf.columns = ['考勤所属部门','部门人数','加班总时长（分钟）','出勤天数',self.thisMonth+'月工作日数','人均加班时长（分钟）','日均加班时长（分钟）','人均日均加班时长（分钟）','人均考勤率']
#        extraDepartureDf.to_excel(extraXlsx,'部门汇总', index = False)
#        extraBossDf.columns = ['考勤号码','姓名','加班总时长（分钟）','出勤天数',self.thisMonth+'月工作日数','日均加班时长（分钟）','人均考勤率']
#        extraBossDf.to_excel(extraXlsx,'管理序列汇总', index = False)
#        extraAllDf.columns = ['考勤号码','姓名','部门','日期','时间','签退时间','下班时间','加班时长','加班时长（分钟）']
#        extraAllDf.to_excel(extraXlsx,'个人明细', index = False)
#        personExtraSummary.columns = ['姓名','该月平均每天加班时长（分钟）']
#        personExtraSummary.to_excel(extraXlsx,'员工加班时长排序', index = False)
#        extraXlsx.save()
#        extraXlsx.close()
#        print ('-----------'+self.thisYear+'年'+self.thisMonth+'月加班情况统计汇总表.xlsx'+' 已生成-----------')
        
        #######################屏蔽掉missAnalyzer在wideTable中暂时缺失的问题
        missAllDf, missPersionDf, missDepDf = self.missAnalyzer()
        missXlsx = ExcelWriter(self.op+self.thisYear+'年'+self.thisMonth+'月无考勤情况统计汇总表.xlsx')
        missDepDf.columns = ['部门','部门人数','部门工作日无考勤人次','人均无考勤人次','当月累计无考勤天数超过10个工作日的人数']
        missDepDf.to_excel(missXlsx,'部门汇总', index = False)
#########bug
        missPersionDf.columns = ['考勤号码','姓名','部门','无考勤天数','当月累计无考勤天数是否超过10个工作日']
        missPersionDf.to_excel(missXlsx,'个人汇总', index = False)
        missAllDf.columns = ['考勤号码','姓名','部门','日期','时间','签到时间','上班时间','迟到时长','加班时长（分钟）']
        missAllDf.to_excel(missXlsx,'考勤原始明细', index = False)
        missXlsx.save()
        missXlsx.close()
        print ('-----------'+self.thisYear+'年'+self.thisMonth+'月无考勤情况统计汇总表.xlsx'+' 已生成-----------')



#        return lateAllDf, late2HDf, late2HomitDf, lateSummary, lateSummaryDep, \
#                extraAllDf, extraBossDf, extraDepartureDf, personExtraSummary, extraSummaryDf, \
#                missAllDf, missPersionDf, missDepDf
        
    def lateAnalyzer(self):
        #统计该月每个人迟到明细（一条一次 ） ok
        #在上面的表中统计出每个人迟到2小时以内的明细（一条一次）
        #在上面的表中去除晚签到的因素后每个人迟到2小时内的明细（一条一次）
        #汇总个人迟到2小时内的包含和不包含晚签到的信息（一人一条）        
        #部门汇总
        #包含晚签到 即超过30分就算迟到
        #返回5个dataFrame
        df = self.wideTable
        lateAllDf = df[(df['isHoliday']==False) & (df['lateMins']>0)][['userid',
                       'name','departure_x','date','time','firstTime','standardonWorkTime','lateTime', 'lateMins']]
        
        late2HDf = lateAllDf[lateAllDf['lateMins'] <= 120]
        def filterLate(x):
            tempX1 = x[x['lateMins'] <= 15].iloc[2:,:]
            tempX2 = x[x['lateMins'] > 15]
            return pd.concat([tempX1, tempX2])
        late2HomitDf = pd.DataFrame(late2HDf.groupby(['userid','name','departure_x']).apply(filterLate).values)
        late2HomitDf.columns = ['userid', 'name', 'departure_x', 'date', 'time', 'firstTime','standardonWorkTime', 'lateTime', 'lateMins']
        late2HomitDf['date'] = late2HomitDf['date'].apply(lambda x :str(x)[:10]) 
        def calc(x):
            x['isOver3'] = 1 if len(x) > 3 else 0
            x['times'] = len(x) 
            return x
        df1 = late2HomitDf.groupby(['userid','name','departure_x']).apply(calc)
        df2 = late2HDf.groupby(['userid','name','departure_x']).apply(calc)
        lateSummary = pd.merge(df2, df1, on = ['userid','name','departure_x'], how = 'left')
        lateSummary = lateSummary[['userid','name','departure_x','isOver3_x', 'times_x', 'isOver3_y','times_y']]
        lateSummary.columns = ['userid','name','departure_x', 'isOver3', 'times', 'isOver3Omit', 'timesOmit']
        lateSummary = lateSummary[['userid','name','departure_x', 'isOver3Omit', 'timesOmit', 'isOver3', 'times']]
        lateSummary = lateSummary.fillna(0)
        lateSummary = lateSummary.drop_duplicates()
        lateSummary[['isOver3Omit', 'timesOmit', 'isOver3', 'times']] = lateSummary[['isOver3Omit', 'timesOmit', 'isOver3', 'times']].astype(int)
        def f(x):
            a = len(set(x['name']))
            c = len(x)
            return pd.DataFrame({'menber':[a],'workDayNumber':[self.workDayNumber], 'allAttenceDays':[c]})     
        df1 = df[df['isHoliday']==False].groupby(['departure_x']).apply(f).reset_index()[['departure_x', 'menber','workDayNumber','allAttenceDays']]
        def f1(x):
            timesOmit = np.sum(x['timesOmit'])
            personOver3Omit = len(x[x['isOver3Omit'] == 1])
            timsOver3Omit = np.sum(x[x['isOver3Omit'] == 1]['timesOmit'])
            times = np.sum(x['times'])
            personOver3 = len(x[x['isOver3'] == 1])
            timsOver3 = np.sum(x[x['isOver3'] == 1]['times'])
            return pd.DataFrame({'timesOmit':[timesOmit],'personOver3Omit':[personOver3Omit], 'timsOver3Omit':[timsOver3Omit],'times':[times],'personOver3':[personOver3], 'timsOver3':[timsOver3]})      
        df2 = lateSummary.groupby(['departure_x']).apply(f1).reset_index()[['departure_x', 'timesOmit','personOver3Omit','timsOver3Omit','times','personOver3','timsOver3']]
        lateSummaryDep = pd.merge(df1, df2, how = 'left', on = ['departure_x'])
        lateSummaryDep = lateSummaryDep.fillna(0)
        lateSummaryDep = lateSummaryDep.drop_duplicates()
        lateSummaryDep[['timesOmit','personOver3Omit','timsOver3Omit','times','personOver3','timsOver3']] = lateSummaryDep[['timesOmit','personOver3Omit','timsOver3Omit','times','personOver3','timsOver3']].astype(int)
        def f2(x):
            x['latePerMonthOmit'] = format(x['timesOmit']/x['allAttenceDays'] ,'.2%')
            x['Per1Omit'] = format(x['personOver3Omit']/x['menber'] ,'.2%')
            x['Per2Omit'] = format(float(x['timsOver3Omit']/x['timesOmit']) if x['timesOmit'] != 0 else 0.0 , '.2%')
            x['latePerMonth'] = format(float(x['times']/x['allAttenceDays']),'.2%')
            x['Per1'] = format(float(x['personOver3']/x['menber']),'.2%')
            x['Per2'] = format(float(x['timsOver3']/x['times']) if x['times'] != 0 else 0.0 ,'.2%')
            return x
        lateSummaryDep = lateSummaryDep.apply(f2, axis=1)
        lateSummaryDep = lateSummaryDep[['departure_x','menber','workDayNumber','allAttenceDays','timesOmit','latePerMonthOmit','personOver3Omit','Per1Omit','timsOver3Omit','Per2Omit','times','latePerMonth','personOver3','Per1','timsOver3','Per2']]
        return lateAllDf, late2HDf, late2HomitDf, lateSummary, lateSummaryDep
        
#    def extraAnalyzer(self):
#        #在考勤正常的日子里选择工作日统计加班情况
#        #统计该月加班明细（一次一条）（所有的，去除假日的，去除不考勤的各一张）
#        #统计管理人员加班信息（一人一条）
#        #统计部门加班信息（一个部门一条）
#        #排序员工加班时长
#        #汇总全院加班情况
#        #出7个DataFrame
#        df = self.wideTable
#        extraAllDf = df[(df['isHoliday']==False) & (df['extraMins']>=0)][['userid','name','departure_x','date','time','lastTime','standardoffWorkTime','extraTime', 'extraMins']]
#        def statics1(x):
#            a = np.sum(x['extraMins'])
#            b = len(x)
#            c = self.workDayNumber
#            d = int(a/c)
#            e = float(b/c)
#            f = list(x['userid'])[0]
#            return pd.DataFrame({'extraTimeAll':[a],'attendenceDayNumber':[b], 'workDayNumber':[c],'extraTimeAvg':[d],'attendencePer':[e], 'userid':[f]}) 
#        extraBossDf = extraAllDf[extraAllDf['departure_x'] == '管理序列'].groupby(['name']).apply(statics1).reset_index()[['userid', 'name','extraTimeAll','attendenceDayNumber','workDayNumber','extraTimeAvg','attendencePer']]
#        def f(x):
#            a = len(set(x['name']))
#            b = np.sum(x['extraMins'])
#            c = len(x)
#            d = self.workDayNumber
#            e = int(b/a)
#            f = int(b/d)
#            g = int(b/d/a)
#            h = float(c/a/d)
#            return pd.DataFrame({'menber':[a],'extraAll':[b], 'allAttenceDays':[c],'workDayNumber':[d],'extraPerPerson':[e],'extraPerDay':[f],'extraPerDayandPerson':[g],'attendencePer':[h]}) 
#        extraDepartureDf = extraAllDf.groupby(['departure_x']).apply(f).reset_index()[['departure_x', 'menber','extraAll','allAttenceDays','workDayNumber','extraPerPerson','extraPerDay','extraPerDayandPerson','attendencePer']]
#        personExtraSummary = extraAllDf[['name','extraMins']].groupby(['name']).apply(lambda x:int(np.sum(x['extraMins'])/len(x))).sort_values(ascending=False)                                                                         
#        extraSummaryDf = pd.DataFrame([len(df.groupby(['name'])), len(df.groupby(['name'])), len(personExtraSummary), \
#         1.0*len(personExtraSummary)/len(df.groupby(['name'])), len(personExtraSummary[personExtraSummary<5]), \
#         1.0*len(personExtraSummary[personExtraSummary<5])/len(personExtraSummary), len(personExtraSummary[personExtraSummary<10]), \
#         1.0*len(personExtraSummary[personExtraSummary<10])/len(personExtraSummary), len(personExtraSummary[personExtraSummary<15]), \
#         1.0*len(personExtraSummary[personExtraSummary<15])/len(personExtraSummary), len(personExtraSummary[personExtraSummary>30]), \
#         1.0*len(personExtraSummary[personExtraSummary>30])/len(personExtraSummary), len(personExtraSummary[personExtraSummary>60]), \
#         1.0*len(personExtraSummary[personExtraSummary>60])/len(personExtraSummary)])       
#        personExtraSummary = personExtraSummary.reset_index()
##        personExtraSummary.columns = ['姓名','该月平均每天加班时长（分钟）']
#        return extraAllDf, extraBossDf, extraDepartureDf, personExtraSummary, extraSummaryDf.T
    
    def missAnalyzer(self):
        #miss算不算半天缺勤的？应该也要算
        df = self.wideTable
####isAttendence的意义所在，无数据表征出勤，悖论       
        #missAllDf = df[(df['isAttandence']==0)&(df['isHoliday']==False)&(df['isVacation']==False)][['userid','name','departure_x','date','time','firstTime','standardonWorkTime','lateTime','lateMins']]
        missAllDf = df[(df['isAttandence']==0)&(df['isHoliday']==False)&(df['isVacation']==False)][['userid','name',
                       'departure_x','date','time','firstTime','standardonWorkTime','lateTime','lateMins']]
        missAllDf['lateMins'] = 0

        def f(x):
            a = 1 if len(x)>=10 else 0
            return pd.DataFrame({'missDay':[len(x)],'missOver10':[a]})  
        missPersionDf = missAllDf.groupby(['userid','name','departure_x']).apply(f).reset_index()[['userid','name','departure_x','missDay','missOver10']]
        
        def f1(x):
            #集合对象是一组无序排列的可哈希的值，集合成员可以做字典中的键。
            #亦可用groupby实现
            a = len(set(x['name']))
            b = np.sum(x['missDay'])    
            return pd.DataFrame({'member':[a],'missPersonTimes':[b],'missPerPerson':[format(1.0*b/a ,'.2f')],'missOver10All':[np.sum(x['missOver10'])]})                   
        missDepDf = missPersionDf.groupby(['departure_x']).apply(f1).reset_index()[['departure_x','member','missPersonTimes','missPerPerson','missOver10All']]
        return missAllDf, missPersionDf, missDepDf

    