# -*- coding: utf-8 -*-
"""
Created on Sun May 27 14:37:24 2018

@author: horsehill
"""
import pandas as pd
from pandas import ExcelWriter
import os

class split():
    
    def __init__(self, wideTable,attendanceRecord, attendanceSummary,originalFile, outputPath, month,monthDay ):
        self.wideTable = wideTable
        self.attRecord = attendanceRecord
        self.attSummary = attendanceSummary
        self.of = originalFile
        self.op = outputPath
        self.thisYear = month[:4]
        self.thisMonth = str(int(month[4:]))
        self.monthDay =monthDay
        
        pass
    
    def main(self):
        self.splitOriginal()
        print ('-----------原始考勤数据表拆分完毕-----------')
        self.splitAttRecord()
        print ('-----------考勤记录表拆分完毕-----------')
        self.splitAttSummary()
        print ('-----------考勤汇总表拆分完毕-----------')
        
    def splitAttRecord(self):
        df = self.attRecord
        path= self.op + '各部门考勤记录表/'
        isExists=os.path.exists(path)
        if not isExists:
            os.makedirs(path)
        
        list_temp1=['部门','考勤号码','姓名','1','2','3','4','5','6','7',\
                                    '8','9','10','11','12','13','14','15','16','17','18','19','20',\
                                    '21','22','23','24','25','26','27','28','29','30','31']
        if(self.monthDay == 31):
            list_valuable =list_temp1[0:]
        elif(self.monthDay ==30):
            
            list_valuable =list_temp1[0:33]
        elif(self.monthDay==28):
            
            list_valuable =list_temp1[0:31]
        elif(self.monthDay==29):
            
            list_valuable =list_temp1[0:32]
        def saveFile1(x):
            dept = list(x['工作部门'])[0].strip().replace('/','和').replace('?','')
            if dept != '管理序列':
                xlsx = ExcelWriter(path+dept+self.thisYear+'年'+self.thisMonth+'月员工考勤记录表.xlsx')
                x.iloc[:,1:].to_excel(xlsx,'员工考勤记录表', index = False, header = True)
                xlsx.save()
                xlsx.close()
        df.groupby(['工作部门']).apply(saveFile1)
        
        def saveFile(x):
            dept = list(x['部门'])[0].strip().replace('/','和').replace('?','')
            if not os.path.exists(path+dept+self.thisYear+'年'+self.thisMonth+'月员工考勤记录表.xlsx'):
#                print (dept)
                xlsx = ExcelWriter(path+dept+self.thisYear+'年'+self.thisMonth+'月员工考勤记录表.xlsx')
                temp = x[list_valuable]
                temp.to_excel(xlsx,'员工考勤记录表', index = False, header = True)
                xlsx.save()
                xlsx.close()
        df.groupby(['部门']).apply(saveFile)
        def saveFile2(x):
            name = list(x['姓名'])[0]
            xlsx = ExcelWriter(path+name+self.thisYear+'年'+self.thisMonth+'月员工考勤记录表.xlsx')
            temp = x[list_valuable]
            temp.to_excel(xlsx,'员工考勤记录表', index = False, header = True)
            ##index=false 不写行名（索引） 
            ##header=true 写出列名，如果是给定字符串列表，则假定它是列表名称的别名
            xlsx.save()
            xlsx.close()
        df[df['部门']=='管理序列'].groupby(['姓名']).apply(saveFile2)

    def splitAttSummary(self):
        df = self.attSummary
        path= self.op + '各部门考勤汇总表/'
        isExists=os.path.exists(path)
        if not isExists:
            os.makedirs(path)
        def saveFile1(x):
            dept = list(x['工作部门'])[0].strip().replace('/','和').replace('?','')
#            if dept != '管理序列':
            xlsx = ExcelWriter(path+dept+self.thisYear+'年'+self.thisMonth+'月员工考勤汇总表.xlsx')
            temp = x[['工作部门','考勤号码','姓名','出勤天数',\
            '说明1','出差、会议、培训等天数','说明2','迟到次数','说明3','早退次数','说明4','缺勤天数','说明5',\
            '法定+企业年休假天数','说明6','福利积点兑换年休假','说明7','病假','说明8','事假','说明9','产假','说明10',\
            '其他假期','说明11','备注']]
            temp.to_excel(xlsx,'员工考勤汇总表', index = False, header = True)
            xlsx.save()
            xlsx.close()
        df.groupby(['工作部门']).apply(saveFile1)    
            
        def saveFile(x):
            dept = list(x['部门'])[0].strip().replace('/','和').replace('?','')
            if not os.path.exists(path+dept+self.thisYear+'年'+self.thisMonth+'月员工考勤汇总表.xlsx') and dept != '管理序列':
#                print (dept)
                xlsx = ExcelWriter(path+dept+self.thisYear+'年'+self.thisMonth+'月员工考勤汇总表.xlsx')
                temp = x[['部门','考勤号码','姓名','出勤天数',\
        '说明1','出差、会议、培训等天数','说明2','迟到次数','说明3','早退次数','说明4','缺勤天数','说明5',\
        '法定+企业年休假天数','说明6','福利积点兑换年休假','说明7','病假','说明8','事假','说明9','产假','说明10',\
        '其他假期','说明11','备注']]
#                print(x.head(1))
#                print (temp.head(1))
                temp.to_excel(xlsx,'员工考勤汇总表', index = False, header = True)
                xlsx.save()
                xlsx.close()
        df.groupby(['部门']).apply(saveFile)

        
        def saveFile2(x):
            name = list(x['姓名'])[0]
            xlsx = ExcelWriter(path+name+self.thisYear+'年'+self.thisMonth+'月员工考勤汇总表.xlsx')
            temp = x[['部门','考勤号码','姓名','出勤天数',\
        '说明1','出差、会议、培训等天数','说明2','迟到次数','说明3','早退次数','说明4','缺勤天数','说明5',\
        '法定+企业年休假天数','说明6','福利积点兑换年休假','说明7','病假','说明8','事假','说明9','产假','说明10',\
        '其他假期','说明11','备注']]
            temp.to_excel(xlsx,'员工考勤汇总表', index = False, header = True)
            xlsx.save()
            xlsx.close()
        df[df['部门']=='管理序列'].groupby(['姓名']).apply(saveFile2)
    
    def splitOriginal(self):
        df = self.wideTable
        #df.columns = ['userid','name','departure','date','time']
        def saveFile1(x):
            path= self.op + '管理序列原始数据拆分表/'
            isExists=os.path.exists(path)
            if not isExists:
                os.makedirs(path)    
            name = list(x['name'])[0]
            xlsx = ExcelWriter(path+name+self.thisYear+'年'+self.thisMonth+'月'+'原始数据表.xlsx')            
            temp = x[['userid','name','departure_x','date','output_time']]
            temp.columns =[['考勤号码','姓名','部门','日期','时间']]
            temp.to_excel(xlsx,'管理序列', index = False, header = True)
            xlsx.save()
            xlsx.close()
        df[df['departure_x']=='管理序列'].groupby(['name']).apply(saveFile1)
        def saveFile2(x):
            departure=list(x['departure_x'])[0].strip().replace('/','和').replace('?','')
            path= self.op + '合同制各部门原始数据拆分表/'
            isExists=os.path.exists(path)
            if not isExists:
                os.makedirs(path)
            
            xlsx = ExcelWriter(path+departure+self.thisYear+'年'+self.thisMonth+'月'+'原始数据表.xlsx')
            temp = x[['userid','name','departure_x','date','output_time']]
            temp.columns =[['考勤号码','姓名','部门','日期','时间']]
            temp.to_excel(xlsx,'合同制部门原始数据', index = False, header = True)
            xlsx.save()
            xlsx.close()
        df[(df['departure_x']!='AMT') & (df['departure_x']!='浩方')&(df['departure_x']!='微企')].groupby(['departure_x']).apply(saveFile2)
        def saveFile3(x):
            
            path= self.op + '外协各部门原始数据拆分表/'
            isExists=os.path.exists(path)
            if not isExists:
                os.makedirs(path)
            workDept =list(x['workDept'])[0]
            xlsx = ExcelWriter(path+workDept+self.thisYear+'年'+self.thisMonth+'月'+'原始数据表.xlsx')
            temp = x[['userid','name','departure_x','date','output_time']]
            temp.columns =[['考勤号码','姓名','部门','日期','时间']]
            temp.to_excel(xlsx,'合同制部门原始数据', index = False, header = True)
            xlsx.save()
            xlsx.close()       
        df[(df['departure_x']=='AMT') | (df['departure_x']=='浩方')|(df['departure_x']=='微企')].groupby(['workDept']).apply(saveFile3)
        def saveFile4(x):
            path= self.op + '实习生各部门原始数据拆分表/'
            isExists=os.path.exists(path)
            if not isExists:
                os.makedirs(path)
            workDept =list(x['workDept'])[0]
            xlsx = ExcelWriter(path+workDept+self.thisYear+'年'+self.thisMonth+'月'+'原始数据表.xlsx')
            temp = x[['userid','name','departure_x','date','output_time']]
            temp.columns =[['考勤号码','姓名','部门','日期','时间']]
            temp.to_excel(xlsx,'合同制部门原始数据', index = False, header = True)
            xlsx.save()
            xlsx.close()
        df[df['departure_x']=='实习生'].groupby(['workDept']).apply(saveFile4)
        
        def saveFile5(x):
            path= self.op + '所有管理序列原始数据拆分表/'
            isExists=os.path.exists(path)
            if not isExists:
                os.makedirs(path)                
            xlsx = ExcelWriter(path+self.thisYear+'年'+self.thisMonth+'月'+'原始数据表.xlsx')            
            temp = x[['userid','name','departure_x','date','output_time']]
            temp.columns =[['考勤号码','姓名','部门','日期','时间']]
            temp.to_excel(xlsx,'管理序列', index = False, header = True)
            xlsx.save()
            xlsx.close()
        t=df[df['departure_x']=='管理序列'].groupby(['departure_x']).apply(saveFile5)
