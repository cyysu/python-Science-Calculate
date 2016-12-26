# -*- coding: utf-8 -*-
# @Author: Kali
# @Date:   2016-12-24 10:25:13
# @Last Modified by:   Kali
# @Last Modified time: 2016-12-26 11:07:44

import xlrd
import matplotlib.pyplot as plt
import numpy as np
from   scipy.optimize import leastsq
import sys  


plt.rcParams['font.sans-serif'] = ['SimHei']


#打开excel文件
data = xlrd.open_workbook('DataReal.xlsx')

#获取第一张工作表（通过索引的方式）
table = data.sheets()[0]

# 获取时间列
TimeColomn = []
for rownum in range(1,table.nrows):
     row = table.row_values(rownum)
     TimeColomn.append(row[0])

# 获取温度列
TemperatureColomn = []
for rownum in range(1,table.nrows):
     row = table.row_values(rownum)
     TemperatureColomn.append(row[5])

# 获取功率列
PowerColomn = []
for rownum in range(1,table.nrows):
     row = table.row_values(rownum)
     PowerColomn.append(row[4])


# 提取所需图像数据   设定阈值为 75
ExpectedPowerWave = []
ExpectedTimeWave  = []
ExpectedTemperatureWave = []
IndexCount = 0
for Index in range(1,len(TemperatureColomn)):
    if(TemperatureColomn[Index - 1] == TemperatureColomn[Index]):
            IndexCount += 1
            if(IndexCount >= 75):
                IndexCount = 0
                if (PowerColomn[Index] - PowerColomn[Index - 1]) < 0:
                    ExpectedPowerWave.append(PowerColomn[Index])
                    ExpectedTimeWave.append(TimeColomn[Index])
                    ExpectedTemperatureWave.append(TemperatureColomn[Index])

# 时间提取写入 时间.txt 文件
wf = open(unicode("时间.txt","utf-8"), 'w')
for i in range(0,len(ExpectedTimeWave)):
    wf.write(str(ExpectedTimeWave[i]))
    wf.write('\n')

# 功率提取写入 功率.txt 文件
wf = open(unicode("功率.txt","utf-8"), 'w')
for i in range(0,len(ExpectedPowerWave)):
    wf.write(str(ExpectedPowerWave[i]))
    wf.write('\n')

# 功率提取写入 温度.txt 文件
wf = open(unicode("温度.txt","utf-8"), 'w')
for i in range(0,len(ExpectedTemperatureWave)):
    wf.write(str(ExpectedTemperatureWave[i]))
    wf.write('\n')


# 存储分段的索引值
IndexSectionFirstArray   = []
IndexSectionSecondArray  = []
IndexSectionThirdArray   = []
SecondFlagStart          = 0
ThirdFlagStart           = 0
SecondFlag               = 0
ThirdFlag                = 0
# IndexAArray 索引值
IndexFirstArray          = []
IndexSecondArray         = []
IndexThirdArray          = []

# 索引标志位
IndexCountNum  = 0
IndexCountNum1 = 0
IndexCountNum2 = 0
# 获取
for Index in range(1,len(ExpectedTemperatureWave)):
   # 初次判断
   if(ExpectedTemperatureWave[Index] - ExpectedTemperatureWave[Index - 1] <= 50):
            SecondFlag = 1

            if ThirdFlagStart == 1:
                if(ExpectedTemperatureWave[Index] - ExpectedTemperatureWave[Index - 1] <= 50):
                    IndexCountNum2 += 1
                    if IndexCountNum2 > 6:
                        IndexThirdArray.append(Index)

            elif SecondFlagStart == 1 and ThirdFlagStart != 1:
                IndexSecondArray.append(Index)

            elif ThirdFlagStart != 1:
                IndexFirstArray.append(Index)

   # 开始进行第二次判断                 设置权值为30               并且设置容错数值为10
   if SecondFlag == 1:
        IndexCountNum += 1
        if(ExpectedTemperatureWave[Index] - ExpectedTemperatureWave[Index - 1] > 30 and IndexCountNum >= 10):
            SecondFlagStart = 1
            ThirdFlag = 1
            # SectionSecondTemperature.append(ExpectedTemperatureWave[Index])

   #  开始进行第三次判断                设置权值为50               并且设置容错数值为10
   if ThirdFlag == 1:
        IndexCountNum1 += 1
        if(ExpectedTemperatureWave[Index] - ExpectedTemperatureWave[Index - 1] > 30 and IndexCountNum1 >= 10):
            ThirdFlagStart  = 1
            # SectionThirdTemperature.append(ExpectedTemperatureWave[Index])



'''   分段函数对应的 x y 坐标
        SectionFirstTime        --->      SectionFirstPower
        SectionSecondTime       --->      SectionSecondPower
        SectionThirdTime        --->      SectionThirdPower
'''

# 根据相应的索引值提取分段函数
# section time variable
SectionFirstTime   = []
SectionSecondTime  = []
SectionThirdTime   = []
# section power variable
SectionFirstPower  = []
SectionSecondPower = []
SectionThirdPower  = []

# first section time
for i in IndexFirstArray:
    SectionFirstTime.append(ExpectedTimeWave[i - 1])
    SectionFirstPower.append(ExpectedPowerWave[i])

# second section time
for j in IndexSecondArray:
    SectionSecondTime.append(ExpectedTimeWave[j - 1])
    SectionSecondPower.append(ExpectedPowerWave[j])

# third section time
for k in IndexThirdArray:
    SectionThirdTime.append(ExpectedTimeWave[k - 1])
    SectionThirdPower.append(ExpectedPowerWave[k])


# 原始数据和提取数据对比图像
plt.figure(1)
plt.plot(TimeColomn,TemperatureColomn,'r')
plt.plot(TimeColomn,PowerColomn,'g')
plt.plot(ExpectedTimeWave,ExpectedPowerWave,'b')
plt.xlabel('time')
plt.ylabel('temprature/power')
plt.savefig(u'数据显示.png')


'''   分段函数对应的 x y 坐标
        SectionFirstTime        --->      SectionFirstPower
        SectionSecondTime       --->      SectionSecondPower
        SectionThirdTime        --->      SectionThirdPower
'''

# 分段曲线拟合
Polynomial1 = np.poly1d(np.polyfit(SectionFirstTime,SectionFirstPower,3))
Polynomial2 = np.poly1d(np.polyfit(SectionSecondTime,SectionSecondPower,3))
Polynomial3 = np.poly1d(np.polyfit(SectionThirdTime,SectionThirdPower,3))



# 打印预测函数
print(Polynomial1)
print(Polynomial2)
print(Polynomial3)


# 画出散点图和原坐标系比较
plt.scatter(SectionFirstTime,SectionFirstPower,marker='o',color='m',label='1',s=80)
plt.scatter(SectionSecondTime,SectionSecondPower,marker='o',color='c',label='2',s=80)
plt.scatter(SectionThirdTime,SectionThirdPower,marker='o',color='r',label='3',s=80)
plt.savefig(u'综合图像显示.png')



# 分段函数散点图以及集合曲线图
plt.figure(2)
plt.plot(SectionFirstTime,SectionFirstPower,'r*')
Coff1 = np.polyfit(SectionFirstTime,SectionFirstPower,3)
y = np.polyval(Coff1,SectionFirstTime)
plt.plot(SectionFirstTime,y,linestyle='-')
plt.savefig(u'第一段拟合效果和原数据.png')


plt.figure(3)
plt.plot(SectionSecondTime,SectionSecondPower,'b*')
Coff2 = np.polyfit(SectionSecondTime,SectionSecondPower,3)
y = np.polyval(Coff2,SectionSecondTime)
plt.plot(SectionSecondTime,y,linestyle='-')
#plt.plot(SectionSecondTime,SectionSecondPower,linestyle='*')
plt.savefig(u'第二段拟合效果和原数据.png')


plt.figure(4)
plt.plot(SectionThirdTime,SectionThirdPower,'g*')
Coff3 = np.polyfit(SectionThirdTime,SectionThirdPower,3)
y = np.polyval(Coff3,SectionThirdTime)
plt.plot(SectionThirdTime,y,linestyle='-')
#plt.plot(SectionThirdTime,SectionThirdPower,linestyle='*')
plt.savefig(u'第三段拟合效果和原数据.png')



# 打印三个阶段的表达式
print ("\n\n\n")
print ("first")
print (np.poly1d(Coff1))
print ("\n")
print ("second")
print (np.poly1d(Coff2))
print ("\n")
print ("third")
print (np.poly1d(Coff3))


''' 写入文件测试Demo
    IndexFirstArray IndexSecondArray IndexThirdArray 文件操作
'''
# IndexFirstArray 提取写入 IndexFirstArray.txt 文件
wf = open("IndexFirsArray.txt", 'w')
for i in range(0,len(IndexFirstArray)):
    wf.write(str(IndexFirstArray[i]))
    wf.write('\n')

# IndexSecondArray 提取写入 IndexSecondArray.txt 文件
wf = open("IndexSecondArray.txt", 'w')
for i in range(0,len(IndexSecondArray)):
    wf.write(str(IndexSecondArray[i]))
    wf.write('\n')

# IndexThirdArray 提取写入 IndexThirdArray.txt 文件
wf = open("IndexThirdArray.txt", 'w')
for i in range(0,len(IndexThirdArray)):
    wf.write(str(IndexThirdArray[i]))
    wf.write('\n')


''' 写入文件测试Demo
      SectionFirstPower SectionSecondPower SectionThirdPower 文件操作
'''
# SectionFirstPower 提取写入 SectionFirstPower.txt 文件
wf = open("SectionFirstPower.txt", 'w')
for i in range(0,len(SectionFirstPower)):
    wf.write(str(SectionFirstPower[i]))
    wf.write('\n')

# SectionSecondPower 提取写入 SectionSecondPower.txt 文件
wf = open("SectionSecondPower.txt", 'w')
for i in range(0,len(SectionSecondPower)):
    wf.write(str(SectionSecondPower[i]))
    wf.write('\n')

# SectionThirdTime 提取写入 SectionThirdTime.txt 文件
wf = open("SectionThirdPower.txt", 'w')
for i in range(0,len(SectionThirdPower)):
    wf.write(str(SectionThirdPower[i]))
    wf.write('\n')


''' 写入文件测试Demo
       SectionFirstTime  SectionSecondTime   SectionThirdTime  文件操作
'''
# SectionFirstTime 提取写入 SectionFirstTime.txt 文件
wf = open("SectionFirstTime.txt", 'w')
for i in range(0,len(SectionFirstTime)):
    wf.write(str(SectionFirstTime[i]))
    wf.write('\n')

# SectionSecondTime 提取写入 SectionSecondTime.txt 文件
wf = open("SectionSecondTime.txt", 'w')
for i in range(0,len(SectionSecondTime)):
    wf.write(str(SectionSecondTime[i]))
    wf.write('\n')

# SectionThirdTime 提取写入 SectionThirdTime.txt 文件
wf = open("SectionThirdTime.txt", 'w')
for i in range(0,len(SectionThirdTime)):
    wf.write(str(SectionThirdTime[i]))
    wf.write('\n')




