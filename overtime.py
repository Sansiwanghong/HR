# -*- coding: utf-8 -*-
"""
Created on Sun Sep 23 16:10:43 2018

@author: wanghong
"""

import numpy as np
import pandas as pd
from pandas import ExcelWriter
import os

class fromOvertime():
    def __init__(self,wideTable, holidayThisMonth, outputPath,law_days=22):######默认法定上班天数
        self.wideTable = wideTable
        self.thisYear = str(list(set(self.wideTable['year']))[0])
        self.thisMonth = str(list(set(self.wideTable['month']))[0]) 
       # self.op = outputPath
        ###做一次暂时替换
        self.op = 'E:/Pyproject/HR/output_test2/'
        self.monthDay = np.max(wideTable['day'])
        self.holidayNumber = holidayThisMonth
        self.law_days=law_days
    def main(self):
        isExists=os.path.exists(self.op)
        if not isExists:
            os.makedirs(self.op)
        self.Record()
        print('-----------加班统计表生成完毕----------')
       
    def Record(self):   
        df = self.wideTable                     
        #manager_over =pd.DataFrame({'userid':[0],'name':['str'],'extra_sum':[0],'attend_days':[0],'days':[0],'extra_means'：[0],'percent':[0]})
        manager_over =pd.DataFrame({'userid':[0],'name':['str'],'extra_sum':[0],'attend_days':[0],'days':[0],'extra_means':[0],'percent':[0]})
        overtimeRecord = ExcelWriter(self.op+self.thisYear + '年'+ self.thisMonth +'月'+'加班数据统计表.xlsx')
        def record1():
#########to_Datetime() 
            df['date'] = pd.to_datetime(df['date'],unit ='D')
            overtimeSheet1 = df[['userid','name','departure_x','date','time','lastTime','extraTime']]
            overtimeSheet1.columns =['考勤号码','姓名','部门','日期','时间','离开时间','加班时间']                       
            overtimeSheet1.to_excel(overtimeRecord,'个人明细',index = False , header = True)
#            overtimeRecord.save()
#            overtimeRecord.close()
        record1()
        
        def record2(x):
            law_days =len(list(x['isNoRecord']))-self.holidayNumber##########为计算方便，定值，后改为手工输入
            extra_sum=x['extraMins'].sum()
            attend_days=len(list(x['isNoRecord']))-x['isNoRecord'].sum()
            rate =format(attend_days/law_days,'.2%')
            mf=pd.DataFrame([[list(x['userid'])[0],list(x['name'])[0],extra_sum,attend_days,law_days,extra_sum//law_days,rate]])
            return mf        
        mf=df[df['departure_x']=='管理序列'].groupby(['name']).apply(record2)
        #overtimeRecord2 = ExcelWriter(self.op+self.thisYear + '年'+ self.thisMonth +'月'+'加班数据统计表.xlsx')
        mf.columns =[['考勤号码','姓名','加班总时长（单位：分钟）','出勤天数','工作天数','日均加班时长','人均考勤率']]                       
######--------------??????排序失败
        # mf.sort_index(axis = 0,ascending = False,by =['加班总时长（单位：分钟）'])
        mf.to_excel(overtimeRecord,'管理序列汇总',index = False , header = True)
#        overtimeRecord2.save()
#        overtimeRecord2.close()
        
        def record3(x):###既是管理序列，又是部门的计算办法。
            law_days =len(list(x['isNoRecord']))-self.holidayNumber##########为计算方便，定值，后改为手工输入
            people_num=len(x.groupby(['name']))
            extra_sum=x['extraMins'].sum()
            attend_days=len(list(x['isNoRecord']))-x['isNoRecord'].sum()
            rate=format((len(list(x['isNoRecord']))-x['isNoRecord'].sum())/len(list(x['isNoRecord'])),'.2%')
            af=pd.DataFrame([[list(x['workDept'])[0],people_num,extra_sum,attend_days,law_days,extra_sum//people_num,extra_sum//law_days,
                            extra_sum//law_days//people_num,rate]])       
            return af
        af=df.groupby(['workDept']).apply(record3)               
        #overtimeRecord3 = ExcelWriter(self.op+self.thisYear + '年'+ self.thisMonth +'月'+'加班数据统计表_部门汇总.xlsx')
        af.columns =[['考勤所属部门','部门人数','加班总时长（单位：分钟）','出勤天数','工作日数','人均加班时长（单位：分钟）','日均加班时长（单位：分钟）',
                      '人均日均加班时长（单位：分钟）','人均考勤率']]                       
        af.to_excel(overtimeRecord,'部门汇总',index = False , header = True)
        overtimeRecord.save()
        overtimeRecord.close()
        
                  
#            extra_sum=x['extraMins']+extra_sum
#            if str(x['isNoRecord']) == '0':
#                attend_days=attend_days+1
          
            #overtimeRecord = ExcelWriter(self.op+self.thisYear + '年'+ self.thisMonth +'月'+'加班数据统计表.xlsx')
#           dept = list(x['departure_x'])[0].strip().replace('/','和').replace('?','')
            #if list(x['departure_x'])[0] =='管理序列':


#            if str(x['isNoRecord'])[-1] == '0':
#                x['isNoRecord']=1
#            else:
#                x['isNoRecord']= 0
#                
#            overtimesheet2 = x[['userid','name','departure_x','extraTime','isNoRecord']]
#            extra_sum =x['extraTime'].sum()
#            attend_days =overtimesheet2['isNoRecord'].sum()
#            extra_means=extra_sum//attend_days
#            temp={'userid':x['userid'],'name':x['name'],'extra_sum':[extra_sum],'attend_days':[attend_days],'days':[22],'extra_means':[extra_means],'percent':[0]}       
#            #添加新列
#            manager_over=pd.concat(manager_over,temp)
            
            #overtimesheet2.columns = [['考勤号码','姓名','部门','加班时间']]
#            print(overtimesheet2)
            #overtimesheet2.to_excel(overtimeRecord,'管理序列',index = False , header = True)
            #overtimeRecord.save()
            #overtimeRecord.close()
            

                