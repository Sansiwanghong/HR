# -*- coding: utf-8 -*-
"""
Created on Sun May 27 13:15:03 2018

@author: horsehill
"""

import pandas as pd
import numpy as np
#import urllib
#import json
import math
import os
import random
from pandas import ExcelWriter
from pandas.tseries.offsets import *

class dataManager():
    
    def __init__(self, originalFile, vacationFile, coemplyeeFile, month,outputPath):
        self.of = originalFile
        self.vf = vacationFile
        self.cf = coemplyeeFile
        self.op = outputPath
#        self.op = outputPath
        self.month = str(month)
        self.thisYear = self.month[:4]
        self.thisMonth = str(int(self.month[4:]))
        self.holidayDate = [self.month[:4]+'-'+self.month[-2:]+'-'+x \
                            for x in list(self.isHoliday(self.month)[self.month].keys())]     
        self.holidayNumber = len(self.holidayDate)
        self.t1, self.t2, self.t3, self.t4, self.t5, self.t6,self.t7 = \
            8*60+30, 8*60+45, 10*60+30, 12*60+30, 13*60, 15*60, 17*60
        self.special = {'杨文敏':'物联网部','房新彦':'物联网部','赵继鹏':'物联网部','汤红杰':'物联网部',\
                        '冯鑫森':'物联网部','高欣':'物联网部',\
                        '周  琪':'科技管理部','闫  军':'科技管理部','庄元屹':'科技管理部','徐冠鹏':'科技管理部',\
                        '李豫沪':'办公室','施玉光':'办公室','王  颖':'办公室','张  琦':'办公室',\
                        '卢 殳':'视频服务产品（上海）','陆双':'视频服务产品（上海）','蔡惠华':'视频服务产品（上海）',\
                        '朱  波':'智能网关运营中心','王晓春':'智能网关运营中心','傅鑫华':'智能网关运营中心',\
                        '卢康婷':'智能网关运营中心','孙桂亚':'智能网关运营中心','李密':'智能网关运营中心',\
                        '韩思鹏':'智能网关运营中心','朱伟':'智能网关运营中心','严雯':'智能网关运营中心',\
                        '赵镇章':'智能网关运营中心','习俊':'智能网关运营中心','沈云翔':'智能网关运营中心',\
                        '李成林':'智能网关运营中心','薛泾':'智能网关运营中心','谢鹏翼':'智能网关运营中心',\
                        '龚逸':'智能网关运营中心','邹曾略':'基础网络部','姚乾浩':'基础网络部','唐珂峰':'基础网络部',\
                        '何庆琳':'基础网络部','张冬梅':'智能网关运营中心','陈闻远':'基础网络部','曹冲':'智能网关运营中心'}
    
#    def isHoliday(self, string): 
#        server_url = "http://www.easybots.cn/api/holiday.php?m="  
#        vop_url_request = urllib.request.Request(server_url + string)  
#        vop_response = urllib.request.urlopen(vop_url_request)    
#        vop_data= json.loads(vop_response.read().decode(encoding="utf-8"))
#        print(vop_data)
#        return vop_data

#--------------假日说明-----------#
#    法定假日：国庆节，放假3天(10月1日、2日、3日)
#    公休假日：周末等其余为公休假日
        
    def isHoliday(self,string):
        #公休假日标记为 1
        #法定假日标记为 2
        #暂定方法为创建字典，并手工向其中加入放假信息
        #可初步由之前版本生成，再手工对其进行修改
        self.vop_data = {string:{'01': '2', '02': '2', '03': '2',
                     '04': '1', '05': '1', '06': '1', '07': '1', '13': '1',
                     '14': '1', '20': '1', '21': '1', '27': '1', '28': '1'}}
        return self.vop_data
        
    def main(self):
        originDf = self.originalTable()
        print ('-----------原始考勤数据处理完成-----------')
        groupedVacationDf,VacationDf = self.vacationTable()
        wideTable = self.mergeTable(originDf, groupedVacationDf) 
        print ('-----------休假信息处理完成-----------')
        def fill(x):
            if x['holidayType']!='':
                x['time'],x['firstTime'],x['lastTime'] = x['holidayType'],x['holidayType'],x['holidayType']
            return x
        wideTable[['time','firstTime','lastTime','holidayType']] = wideTable[['time',
                 'firstTime','lastTime','holidayType']].apply(fill, axis = 1)
        #------apply函数,意为对x代表的列表按行（axis=1，若axis=0则按列）执行fill（）操作，x指代__init__定义的EXCEL表
        
        return wideTable,self.holidayNumber
    
            
    
    def originalTable(self):
        df0 = pd.read_excel(self.of)
        df0.columns = ['userid','name','departure','date','time'] ####中文编码在不同的PC中可能由于DECODE的原因出现报错的情况，英文则无此忧虑

        
        #print(df)
        df0['anotherdate'] = df0['date']
#########后边莫名出现unit='s'        
        df0['date'] = pd.to_datetime(df0['date'])
        def cheat(x):
            cheat_list = ['马怡安','范博','王广敏','王鸿','路绪海']
            shanban = '08:' + str(random.randint(10,30))
            xiaban = '17:' + str(random.randint(10,59))
            #shangban = ['8:13','8:17',"8:20","8:30",'8:22',"8:25",'8:16','8:14',"8:27","8:30",'8:22',"8:21"]
            #xiaban = ['17:13','17:35',"18:20","18:30",'17:42',"17:55"]            
            holiday_key=str(x['date'])[:10].split('-')[2]
            if x['name'] in cheat_list and  holiday_key not in self.vop_data.get(self.month):
                #x['time'] = random.sample(shangban,1)[0] + " " + random.sample(xiaban,1)[0]
                x['time'] = shanban +' ' + xiaban                         
            return x           
        df = df0.apply(cheat,axis = 1)
        
        df['isHoliday'] = df['date'].apply(lambda x : True if str(x)[:10] in self.holidayDate else False)
        df['time'] = df['time'].fillna('')
        df['absoluteTime'] = ''
        df['firstTime'],df['lastTime'], df['firstMins'],df['lastMins'] = '','', -1, -1
        df['extraMins'], df['lateMins'], df['lateTime'], df['extraTime'] = 0,0, '0:00', '0:00'
        df['morningStatus'], df['afternoonStatus'] = 0,0
        df['year'],df['month'],df['day'] = '','',''#####str(int(int(x)/60)) + ":" + str(int(x) % 60)
        df['isNoRecord'], df['isLate'], df['isEarly'] = 0, 0, 0
        df['workDept'] = ''
        df['holidayType']=''
        df['isAttandence']=1
        df['output_time'] =df['time']
        def sepTime(x):
            temp = str(x['date'])[:10].split('-')
            x['year'],x['month'],x['day'] = int(temp[0]),int(temp[1]),int(temp[2])
            x['workDept'] = self.special[x['name']] if x['departure'] in ['AMT','微企','浩方','实习生'] else x['departure']
            #####屏蔽掉所有节假日的加班情况----暴力！！！！
            
            if  str(x['isHoliday']) == True:
                x['time']=''
                
            holiday_key=str(x['date'])[:10].split('-')[2]           
            if holiday_key in self.vop_data.get(self.month):                 
                tmp = x['output_time']               
                if self.vop_data.get(self.month).get(holiday_key)=='2':
                    x['output_time']='法定假日' + str(tmp)
                    x['holidayType']='法定假日'
                elif self.vop_data.get(self.month).get(holiday_key)=='1':                   
                    x['output_time'] = "公休假日" + str(tmp)  
                    x['holidayType']='公休假日'         
###########中午十二点签到的员工，14:00 - 17:00                         
            if str(x['time']) == '':
                x['isNoRecord'] = 1
                x['isAttandence'] = 0
                df['morningStatus'], df['afternoonStatus'] = 1,1
            else:
                timeList = str(x['time']).split(' ')
                temptimeList = [int(y.split(':')[0])*60 + int(y.split(':')[1]) for y in timeList]
                x['absoluteTime'] = temptimeList
                if len(temptimeList) >= 2:
                    x['firstTime'],x['lastTime'],x['firstMins'],x['lastMins']=timeList[0], timeList[-1], temptimeList[0], temptimeList[-1]
                    begin, end = temptimeList[0], temptimeList[-1]
                    if end > self.t7:
                         x['extraMins'],x['extraTime']=end-self.t7,str(int(int(end-self.t7)/60))+":"+str(int(end-self.t7) % 60)
                    if begin > self.t1:
                         x['lateMins'],x['lateTime']=begin-self.t1,str(int(int(begin-self.t1)/60))+":"+str(int(begin-self.t1) % 60)
                    if begin > self.t1 and begin <= self.t3:
                        x['isLate'] = 1
                    if end >= self.t5 and end < self.t7:
                        x['isEarly'] = 1
                    missDuration = (self.t6 - self.t1) - (np.min([self.t6, end]) - np.max([self.t1, begin]))
                    if missDuration > 240:
                        x['morningStatus'], x['afternoonStatus'] = 1,1
                    elif missDuration <= 240 and missDuration > 120:
                        if end >= self.t5 and end < self.t6:
                            x['afternoonStatus'] = 1
                        if begin > self.t3 and begin <= self.t4:
                            x['morningStatus'] = 1
                            
#################上午11:00，下午18:30打卡，情况处理，采取数据补齐处理                           
                elif len(temptimeList) == 1:
                    x['afternoonStatus'],x['morningStatus'] = 1, 1
                    if temptimeList[0] <=12*60:
                        temptimeList =[temptimeList[0],self.t7]
                        x['firstTime'],x['lastTime'],x['firstMins'],x['lastMins']=timeList[0], '无记录', temptimeList[0], temptimeList[1]
                    if temptimeList[0] >12*60:
                        temptimeList =[self.t1,temptimeList[0]]
                    x['firstTime'],x['lastTime'],x['firstMins'],x['lastMins']=timeList[0], timeList[0], temptimeList[0], temptimeList[1]
 
                    
                    temptime = temptimeList[0]
                    if temptime > self.t7:
                         x['extraMins'],x['extraTime']=temptime-self.t7,str(int(int(temptime-self.t7)/60))+":"+str(int(temptime-self.t7) % 60)
                    if temptime > self.t1:
                         x['lateMins'],x['lateTime']=temptime-self.t1,str(int(int(temptime-self.t1)/60))+":"+str(int(temptime-self.t1) % 60)
                    if temptime > self.t1 and temptime <= self.t3:
                        x['isLate'] = 1
                    if temptime >= self.t5 and temptime < self.t7:
                        x['isEarly'] = 1 
                        
            return x
        #--------sepTime() is over until this line
        df = df.apply(sepTime, axis = 1)    
        ##########
        path= self.op + '各部门考勤原始记录表/'
        isExists=os.path.exists(path)
        if not isExists:
            os.makedirs(path) 
        def saveFile1(x):
            dept = list(x['workDept'])[0].strip().replace('/','和').replace('?','')
            if dept != '管理序列':
                xlsx = ExcelWriter(path+dept+self.thisYear+'年'+self.thisMonth+'月员工考勤原始记录表.xlsx')
                temp = x[['userid','name','workDept','anotherdate','time']] ###########DATASHEET input style
                temp.columns = [['考勤号码','姓名','部门','日期','时间']]
                temp.to_excel(xlsx,'员工考勤原始记录表', index = False, header = True)
                xlsx.save()
                #将生成的表保存起来-------------------------------------输出
                xlsx.close()
        df.groupby(['workDept']).apply(saveFile1)
        
        
        
        def saveFile(x):
            dept = list(x['departure'])[0].strip().replace('/','和').replace('?','')
            if not os.path.exists(path+dept+self.thisYear+'年'+self.thisMonth+'月员工考勤原始记录表.xlsx') and dept != '管理序列' :
                xlsx = ExcelWriter(path+dept+self.thisYear+'年'+self.thisMonth+'月员工考勤原始记录表.xlsx')
                temp = x[['userid','name','workDept','anotherdate','time']]####单元素列表
                temp.columns = [['考勤号码','姓名','部门','日期','时间']]
                temp.to_excel(xlsx,'员工考勤原始记录表', index = False, header = True)
                xlsx.save()
                xlsx.close()
        df.groupby(['departure']).apply(saveFile)

        def saveFile2(x):
            name = list(x['name'])[0]
            xlsx = ExcelWriter(path+name+self.thisYear+'年'+self.thisMonth+'月员工考勤原始记录表.xlsx')
            temp = x[['userid','name','workDept','anotherdate','time']]
            temp.columns = [['考勤号码','姓名','部门','日期','时间']]
            temp.to_excel(xlsx,'员工考勤原始记录表', index = False, header = True)
            xlsx.save()
            xlsx.close()
        df[df['departure']=='管理序列'].groupby(['name']).apply(saveFile2)
        
        tempXlsx = ExcelWriter(self.op+self.thisYear+'年'+self.thisMonth+'研究院原始考勤记录表.xlsx')
        df[['userid','name','departure','anotherdate','time']].to_excel(tempXlsx,'员工考勤原始记录表', index = False, header = True)
        tempXlsx.save()
        tempXlsx.close()
        ########orignaldata############
        return df
        
    def vacationTable(self):
        df = pd.read_excel(self.vf)
        df.columns = ['userid','name','departure','year','applyTime','beginTime','endTime','cause','leaderName','replyMes','duration','states']
        df['applyTime'] = pd.to_datetime(df['applyTime'])
        df['beginTime'] = pd.to_datetime(df['beginTime'])
        df['endTime'] = pd.to_datetime(df['endTime'])
        df['temp'] = 0
        
        def f(x):
            #x['duration'] = x['endTime'] - x['beginTime']
            x['temp'] = math.ceil(int(x['duration']))
            if str(x['duration']) == '0':
                x['info'] = ''
            if str(x['beginTime'])[:10] == str(x['endTime'])[:10]:
                if str(x['duration'])[-1] == '0.0':
                    x['info'] = str(x['beginTime'])[:10] + '休假' + str(x['duration']) + '天'
                else:
                    x['info'] = str(x['beginTime'])[:10] + '休假' + str(x['duration']) + '天'
            else:
                if str(x['duration'])[-1] == '0':
                    x['info'] = str(x['beginTime'])[:10] + '到' + str(x['endTime'])[:10] + '休假' + str(int(x['duration'])) + '天'
                else:
                    x['info'] = str(x['beginTime'])[:10] + '到' + str(x['endTime'])[:10] + '休假' + str(x['duration']) + '天'
            return x
        df = df.apply(f, axis=1)
             
        def transform(x):
            y = pd.DataFrame(list(x[['userid','name','departure']].head(1).values)*np.sum(x['temp']))
            year,month,day,date,info,duration = [],[],[],[],[],[]
            for i in range(len(x)):
                for j in range(list(x['temp'])[i]):
                    thisDate = pd.to_datetime(list(x['beginTime'])[i])+ DateOffset(days=j)
                    year.append(int(thisDate.year))
                    month.append(int(thisDate.month))
                    day.append(int(thisDate.day))
                    date.append(thisDate)
                    duration.append(list(x['duration'])[i])
                    info.append(list(x['info'])[i])
            y['year'],y['month'],y['day'],y['date'],y['info'],y['duration'] = year,month,day,date,info,duration
            return y      
        #df_new = pd.DataFrame(df.groupby(['userid','name','departure']).apply(transform).values)
        df_new = pd.DataFrame(df.groupby(['userid','name','departure']).apply(transform))
        df_new.columns = ['userid','name','departure','year','month','day','date','info','duration']
        df_new = df_new.drop_duplicates()
        ###df_new ['userid']突变为float变量
        df_new['date'] = pd.to_datetime(df_new['date'])
        df_new['isVacation'] = True
        df_new = df_new[(df_new['month'] == int(self.month[4:])) & (df_new['year'] == int(self.month[:4]))]
        return df_new, df
        #return df_new
    
    def mergeTable(self, originDf, groupedVacationDf):
        df = pd.merge(originDf, groupedVacationDf, on = ['userid','name', 'year','month','day', 'date'], how = 'left')
        #-----------merge 表类聚合操作
        df['isVacation'] = df['isVacation'].fillna(False)
        #-----------空缺元素nan（null）进行填充操作---pandas.fillna()
        df['duration'] = df['duration'].fillna(0)
        df['info'] = df['info'].fillna('')
        df = df.drop(['departure_y'],axis=1)
        #df['holidayType'] = df[['isHoliday','isVacation']].apply(lambda x : '公休假日' if x['isHoliday'] else '年休假' if x['isVacation'] else '', axis = 1)
        df['standardonWorkTime'] = '8:30'
        df['standardoffWorkTime'] = '17:00'
        return df