# -*- coding: utf-8 -*-
"""
Created on Sun May 27 15:28:09 2018

@author: horsehill
"""

import pandas as pd
import numpy as np
from pandas import ExcelWriter

class formTable():
    
    def __init__(self, wideTable, holidayThisMonth, outputPath):
        self.wideTable = wideTable
        self.thisYear = str(list(set(self.wideTable['year']))[0])
        self.thisMonth = str(list(set(self.wideTable['month']))[0]) 
        self.op = outputPath
        self.op = 'E:/Pyproject/HR/output_test3/'
        self.monthDay = np.max(wideTable['day'])
        self.holidayNumber = holidayThisMonth
        
    def main(self):
        attendanceDf = self.attendance()
        attendanceRecord, attendanceSummary = self.split(attendanceDf)
        return attendanceRecord, attendanceSummary
#        return attendanceDf
        
    def split(self, attendanceDf):
#        print (attendanceDf.columns)
#        print (len(list(attendanceDf.columns)))
        print(self.monthDay)
        list_temp = ['部门','工作部门','考勤号码','姓名','1','2','3','4','5','6','7',\
                                    '8','9','10','11','12','13','14','15','16','17','18','19','20',\
                                    '21','22','23','24','25','26','27','28','29','30','31']
        list_temp1=['部门','考勤号码','姓名','1','2','3','4','5','6','7',\
                                    '8','9','10','11','12','13','14','15','16','17','18','19','20',\
                                    '21','22','23','24','25','26','27','28','29','30','31']
        if(self.monthDay == 31):
            list_head =list_temp[0:]
            list_valuable =list_temp1[0:]
        elif(self.monthDay ==30):
            list_head =list_temp[0:34]
            list_valuable =list_temp1[0:33]
        elif(self.monthDay==28):
            list_head =list_temp[0:32]
            list_valuable =list_temp1[0:31]
        elif(self.monthDay==29):
            list_head =list_temp[0:33]
            list_valuable =list_temp1[0:32]
            
        attendanceRecord = attendanceDf.iloc[:,0:self.monthDay+4]
        attendanceRecord.columns = list_head
#未考虑30天的情况
        attendanceSummary = attendanceDf.iloc[:,self.monthDay+4:].dropna()
        attendanceSummary.columns = ['部门','工作部门','考勤号码','姓名','出勤天数',\
        '说明1','出差、会议、培训等天数','说明2','迟到次数','说明3','早退次数','说明4','缺勤天数','说明5',\
        '法定+企业年休假天数','说明6','福利积点兑换年休假','说明7','病假','说明8','事假','说明9','产假','说明10',\
        '其他假期','说明11','备注','lateTime', 'monthDay', 'holidayNumber']
#        print (attendanceSummary.head(1))
        attendanceRecordXlsx = ExcelWriter(self.op+self.thisYear+'年'+self.thisMonth+'月研究院各部门员工考勤记录表.xlsx')
        attendanceRecord[attendanceRecord['部门']!=attendanceRecord['工作部门']][list_valuable].to_excel(attendanceRecordXlsx,'员工考勤记录表', index = False, header = True)
        attendanceRecordXlsx.save()
        attendanceRecordXlsx.close()
        
        attendanceRecordXlsx = ExcelWriter(self.op+self.thisYear+'年'+self.thisMonth+'月研究院项目合作员工考勤记录表.xlsx')
        attendanceRecord[(attendanceRecord['部门']=='微企')|(attendanceRecord['部门']=='浩方')|(attendanceRecord['部门']=='AMT')][list_valuable].to_excel(attendanceRecordXlsx,'员工考勤记录表', index = False, header = True)
        attendanceRecordXlsx.save()
        attendanceRecordXlsx.close()
        
        attendanceRecordXlsx = ExcelWriter(self.op+self.thisYear+'年'+self.thisMonth+'月研究院实习生考勤记录表.xlsx')
        attendanceRecord[attendanceRecord['部门']=='实习生'][list_valuable].to_excel(attendanceRecordXlsx,'员工考勤记录表', index = False, header = True)
        attendanceRecordXlsx.save()
        attendanceRecordXlsx.close()
        print ('-----------'+self.thisYear+'年'+self.thisMonth+'月研究院各部门员工考勤记录表.xlsx'+' 已生成-----------')
        
        
        
        
        attendanceSummaryXlsx = ExcelWriter(self.op+self.thisYear+'年'+self.thisMonth+'月研究院各部门员工考勤汇总表.xlsx')
        attendanceSummary.to_excel(attendanceSummaryXlsx,'员工考勤汇总表', index = False, header = True)
        attendanceSummaryXlsx.save()
        attendanceSummaryXlsx.close()
        
        
        attendanceSummaryXlsx = ExcelWriter(self.op+self.thisYear+'年'+self.thisMonth+'月研究院项目合作员工考勤汇总表.xlsx')
        attendanceSummary[(attendanceSummary['部门']=='微企')|(attendanceSummary['部门']=='浩方')|(attendanceSummary['部门']=='AMT')].to_excel(attendanceSummaryXlsx,'员工考勤汇总表', index = False, header = True)
        attendanceSummaryXlsx.save()
        attendanceSummaryXlsx.close()
        
        
        attendanceSummaryXlsx = ExcelWriter(self.op+self.thisYear+'年'+self.thisMonth+'月研究院实习生考勤汇总表.xlsx')
        attendanceSummary[attendanceSummary['部门']=='实习生'].to_excel(attendanceSummaryXlsx,'员工考勤汇总表', index = False, header = True)
        attendanceSummaryXlsx.save()
        attendanceSummaryXlsx.close()
        
        print ('-----------'+self.thisYear+'年'+self.thisMonth+'月研究院各部门员工考勤汇总表.xlsx'+' 已生成-----------')

        return attendanceRecord, attendanceSummary
        
    def attendance(self):
        df = self.wideTable
        def tarns(x):
            ret = [[],[]]
            ret[0].append(list(x['departure_x'])[0])
            ret[1].append(list(x['departure_x'])[0])
            ret[0].append(list(x['workDept'])[0])
            ret[1].append(list(x['workDept'])[0])
            ret[0].append(list(x['userid'])[0])
            ret[0].append(list(x['name'])[0])
            ret[1].append(list(x['userid'])[0])
            ret[1].append(list(x['name'])[0])
            lateTime = 0
            attendance,miss,holiday,vacation= 0, 0, 0, 0
            early,late = 0, 0
            

            for i in range(len(x)):
                add1, add2 = '8', '8'
                if list(x['holidayType'])[i] == '公休假日':
                    add1 = '公休假日'
                    add2 = '公休假日'
                    holiday += 1
                if list(x['holidayType'])[i] == '法定假日':
                    add1 = '法定假日'
                    add2 = '法定假日'
                    holiday += 1
                    
                elif list(x['holidayType'])[i] == '年休假':
                    if list(x['duration'])[i] == 0.5:
                        vacation += 0.5
                        if list(x['morningStatus'])[i] != 0:
                            add1 = '年休假'
                        elif list(x['afternoonStatus'])[i] != 0:
                            add2 = '年休假'
                        elif list(x['isLate'])[i] != 0:
                            add1 = '年休假'
                        elif list(x['isEarly'])[i] !=0:
                            add2 = '年休假'
                    else:
                        vacation += 1
                        add1 = '年休假'
                        add2 = '年休假'
                else:
                    if list(x['morningStatus'])[i] != 0:
                        add1 = '缺勤'
                        miss += 0.5
                    if list(x['afternoonStatus'])[i] != 0:
                        add2 = '缺勤'
                        miss += 0.5
                    if list(x['isLate'])[i] != 0:
                        if lateTime < 2:
                            add1 = '晚签到'
                            lateTime  += 1
                        else:
                            add1 = '迟到'
                            late += 1
                    if list(x['isEarly'])[i] != 0:
                        add2 = '早退'
                        early += 1
                    
                ret[0].append(add1)
                ret[1].append(add2)  
            attendance = self.monthDay - miss - holiday- vacation
            information = set(df[df['name']==list(x['name'])[0]]['info'])
            information = str(information)[1:-1] if len(information) != 0 else ''
            ret[0].extend([list(x['departure_x'])[0],list(x['workDept'])[0], str(list(x['userid'])[0]), list(x['name'])[0],\
                          attendance,'', 0,'',late,'',early,'',miss,'',holiday+vacation,information[3:],0,'',0,'',0,'',0,'',0,'','',lateTime, self.monthDay, self.holidayNumber])       
            return pd.DataFrame(ret)
        return pd.DataFrame(df.groupby(['departure_x','userid','name']).apply(tarns).values)
        
        
        