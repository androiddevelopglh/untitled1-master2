# -*- coding:utf-8 -*-
#coding: unicode_escape
import xlrd
import sys
import os
import xlwt
import numpy as np
from xlutils.copy import copy
import csv
import numpy as np
import matplotlib.pyplot as plt
import linecache
import pandas as pd
import re
import pandas as pd




lrb=r'G:\数据\福晶科技\lrb002222.csv'
zcfzb=r'G:\数据\福晶科技\zcfzb002222.csv'
xjllb=r'G:\数据\福晶科技\xjllb002222.csv'
plt.rcParams['font.sans-serif'] = ['SimHei']#让lable可以是中文
plt.rcParams['axes.unicode_minus'] = False #让lable可以是中文
#----------------导入csv文件---------------#
data_zcfzb=pd.read_csv(zcfzb,error_bad_lines=False,encoding='gbk',
                  header=0,index_col=0,dtype= str, na_filter=True)#header=0,index_col=0表示将第一行和第一列作为索引，
data_lrb=pd.read_csv(lrb,error_bad_lines=False,encoding='gbk',
                  header=0,index_col=0,dtype=str, na_filter=True)#header=0,index_col=0表示将第一行和第一列作为索引，
data_xjllb=pd.read_csv(xjllb,error_bad_lines=False,encoding='gbk',
                  header=0,index_col=0,dtype=str, na_filter=True)#header=0,index_col=0表示将第一行和第一列作为索引，
data_zcfzb=data_zcfzb.replace('--','0')#将所有'--'替换为零
data_lrb=data_lrb.replace('--','0')#将所有'--'替换为零
data_xjllb=data_xjllb.replace('--','0')#将所有'--'替换为零

data_zcfzb=data_zcfzb.replace(' --','0')#将所有'--'替换为零
data_lrb=data_lrb.replace(' --','0')#将所有'--'替换为零
data_xjllb=data_xjllb.replace(' --','0')#将所有'--'替换为零

data_zcfzb=data_zcfzb.replace(' ',np.nan)#将所有'空字符串'替换为Nan
data_lrb=data_lrb.replace(' ',np.nan)#将将所有'空字符串'替换为Nan
data_xjllb=data_xjllb.replace(' ',np.nan)#将所有'空字符串'替换为Nan


data_zcfzb=data_zcfzb.dropna(axis = 0,how='all')#去掉全部为一行全部为NaN，删除该行，axis=0表示行
data_lrb=data_lrb.dropna(axis = 0,how='all')    #去掉全部为一行全部为NaN，删除该行，axis=0表示行
data_xjllb=data_xjllb.dropna(axis = 0,how='all')#去掉全部为一行全部为NaN，删除该行，axis=0表示行

data_zcfzb=data_zcfzb.dropna(axis = 1)#去掉包含nan值的列，axis=1表示列
data_lrb=data_lrb.dropna(axis = 1)#去掉包含nan值的列，axis表示列
data_xjllb=data_xjllb.dropna(axis = 1)#去掉包含nan值的列，axis表示列

x=data_zcfzb.columns

x1=x[x1_index]
x2=x[x2_index]
x3=x[x3_index]
x4=x[x4_index]

x1=list(map(lambda x: x[2:4], x1))
x2=list(map(lambda x: x[2:4], x2))
x3=list(map(lambda x: x[2:4], x3))
x4=list(map(lambda x: x[2:4], x4))

data_zcfzb=data_zcfzb.astype('float')
data_lrb=data_lrb.astype('float')
data_xjllb=data_xjllb.astype('float')


货币资金4=data_zcfzb.iloc[0,x4_index]
应收票据4=data_zcfzb.iloc[5,x4_index]
应收账款4=data_zcfzb.iloc[6,x4_index]
货币票据账款合计4=货币资金4+应收票据4+应收账款4
流动负债合计4=data_zcfzb.iloc[83,x4_index]
#货币资金4=list(map(eval,货币资金4))#现将所有非字符串变为字符串，而后再去掉。
货币资金与流动负债比例=货币资金4/流动负债合计4
应收票据与流动负债比例=应收票据4/流动负债合计4
应收账款与流动负债比例=应收账款4/流动负债合计4
货币票据账款合计与流动负债比例4=货币票据账款合计4/流动负债合计4

plt.figure(1)
plt.plot(x4, 货币资金4,'-', label='货币资金4')
plt.plot(x4, 应收票据4, 'o-',label='应收票据4')
plt.plot(x4, 应收账款4, 's-',label='应收账款4')
plt.legend(loc = 'upper right')
plt.xlabel("年份")#x轴上的名字
plt.ylabel("万元")#y轴上的名字

plt.figure(2)

plt.plot(x4, 货币资金与流动负债比例, '--', label='货币资金与流动负债比例')
plt.plot(x4, 应收票据与流动负债比例, 's-', label='应收票据与流动负债比例')
plt.plot(x4, 应收账款与流动负债比例, '*-', label='应收账款与流动负债比例')
plt.plot(x4, 货币票据账款合计与流动负债比例4, 'd-' ,label='货币票据账款合计与流动负债比例4')
plt.legend(loc = 'upper right')
plt.show()
m=1
#--------------第一季度--------------#

#--------------第二季度--------------#

#--------------第三季度--------------#

#--------------年度报告--------------#
