#!/usr/bin/env python
# -*- coding:utf-8 -*-
import datetime
import sys
import os
import numpy as np
import matplotlib.pyplot as plt
import linecache
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages
import textwrap
import time

curPath = os.path.abspath(os.path.dirname(__file__))[:2]#确定路径所在的盘符
dir=curPath+'\全部A股数据\\'#确定数据所在具体文件夹
dir_fig=curPath+'\个股指标图\\'
dir_replace=curPath+'\替换后数据\\'
plt.rcParams['font.sans-serif'] = ['SimHei']#让lable可以是中文
plt.rcParams['axes.unicode_minus'] = False #让lable可以是中文

filepaths = []
def all_files_path(dir):
    for root, dirs, files in os.walk(dir):     # 分别代表根目录、文件夹、文件
        for file in files:                         # 遍历文件
            file_path = os.path.join(root, file)   # 获取文件绝对路径
            filepaths.append(file_path)            # 将文件路径添加进列表
            tempdata=pd.read_excel(file_path,header=0,index_col=0)
            tempdata=tempdata.replace('——',0)#将所有'--'替换为零
            tempdata.to_excel(os.path.join(dir_replace, file))
        for dir in dirs:                           # 遍历目录下的子目录
            dir_path = os.path.join(root, dir)     # 获取子目录路径
            all_files_path(dir_path)               # 递归调用
#all_files_path(dir)


gp_daima=['生益科技', '深南电路','旗滨集团', '正川股份','山东药玻','三安光电','正泰电器','福晶科技','华灿光电','明泰铝业','金发科技','汤臣倍健','鲁阳节能','久立特材','振芯科技']
#gp_daima=['生益科技', '深南电路']


有息负债率= pd.read_excel(dir_replace +  r'有息负债率.xls', header=None, index_col=0,skiprows=[0])  # header=0,index_col=0表示将第一行和第一列作为索引
负债合计= pd.read_excel(dir_replace +  r'负债合计.xls', header=None, index_col=0,skiprows=[0])  # header=0,index_col=0表示将第一行和第一列作为索引
有息负债=有息负债率.iloc[1:-6, 1:]*负债合计.iloc[1:-6, 1:]/100
有息负债=pd.concat([有息负债率.iloc[1:-6,0],有息负债],axis=1)
有息负债.to_excel(os.path.join(dir_replace, r'有息负债.xls'))


指标集合=[]
指标列举原因=[]

指标集合.append(['净利润','经营活动产生的现金流量净额','投资活动产生的现金流量净额','筹资活动产生的现金流量净额'])
指标列举原因.append(textwrap.fill('''经营活动现金流净额持续优于同行，同时投资活动现金净额持续为大额负数时，
    需要考虑是否存在投资活动现金宽出转换为经营活动现金流入的可能。
    好：经营活动产生的现金流量净额>净利润>0
    好：投资活动产生的现金流量净额<0,且主要用于投入新项目''',width=50).strip())

指标集合.append(['销售商品提供劳务收入','营业收入'])
指标列举原因.append('''好：销售商品、提供劳务收到的现金>=营业收入>0''')

#指标集合.append(['存货','毛利率'])
#指标列举原因.append('''库存增加导致毛利率提升需要警惕''')

指标集合.append(['现金及现金等价物净增加额','期末现金及现金等价物余额','有息负债'])
指标列举原因.append('''好：现金及现金等价物净额增加额>0,可放宽位排出分红因素，该科目>0
好：期末现金及现金等价物余额+(银行承兑汇票)>=有息负债''')

指标集合.append(['营业总收入','货币资金','应收票据','应收账款','存货'])
指标列举原因.append('''应收账款大幅增长，且增幅超过同期收入增幅；或应收账款周转率低于同行水平，
且呈明显下降趋势，均可能预示财报操纵''')

指标集合.append(['存货','营业收入','存货周转率','销售毛利率'])
指标列举原因.append('''1存货，尤其是数量和价值不易确定的存货大幅增长，且超过营业成本的增长;
或者存货周转率低于同行水平，并呈现下降趋势，可能是操纵。
2如果存货周转率明显下降，却同时伴生着毛利率的显著上升，投资者基本可以按照造假对待了。''')

指标集合.append(['其他业务收入','其他应收款','其他应付款'])
指标列举原因.append('''1财报上出现金额较大的“其他××”，或者一些很少见过的科目名称，
都是值得投资者警惕的信号。公司其他业务收入在营业收入中的占比突然发生很大提升
2、其他应收款和其他应付款科目，数额极小，甚至为零。值得警惕''')

指标集合.append([ '其他流动资产','流动资产','其他非流动资产','非流动资产', '其他流动负债','流动负债', '其他非流动负债','非流动负债'])
指标列举原因.append(''',金额较大的“其他××”，或者一些很少见过的科目名称，都是值得投资者警惕的信号''')

指标集合.append(['固定资产减值损失','存货跌价损失', '资产减值损失合计','净利润'])
指标列举原因.append(''',资产减值损失包含（固定资产、存货、应收账款、金融资产损失）,
如果资产减值损失同比大增，如果按照上年同期的资产减值损失水平算，
利润同比波动并不大。那么，可能需要考虑企业在“洗大澡”，存在故意做低盈利基数的可能。''')

指标集合.append([ '坏账准备合计','存货跌价准备合计','固定资产减值准备合计', '无形资产减值准备合计','资产减值准备合计', '净利润'])
指标列举原因.append('''固定资产、存货、应收账款、金融资产等大幅计提，乃至价值归零，
                     也要考虑“洗大澡”的可能。''')

指标集合.append(['应付票据','应付账款','预付款项'])
指标列举原因.append('''1预付账款大幅增加，尤其是预付工程款或预付专利或非专利技术内采购款大幅增加
，可能存在通过预付款流出资金，最终以营业收入斗，虚增利润的情况。
2应付账款和应付票据。这两个科目数额大增，要么说明公司在整个商业链条
上地位大增，对供应商变得更加强势了；''')

指标集合.append(['货币资金' ,'流动负债'],)
指标列举原因.append(''' 货币资金比流动负责小很多''')

指标集合.append(['营业总收入','销售费用','管理费用','财务费用'])
指标列举原因.append('''销售管理占比 ''')

varDict = locals()
for 指标组 in 指标集合:
    for 指标 in 指标组:
        try:
            varDict[指标] = pd.read_excel(dir_replace + 指标 + r'.xls', header=0, index_col=0)  # header=0,index_col=0表示将第一行和第一列作为索引

        except:
            print(指标+'读取出错')

线型=['--','o-', 's-','kx-','d-','1-','m-','h-']
x=list(range(1990,2020))
for gp in gp_daima:
    指标原因序号 = 0
    with PdfPages(dir_fig+gp+r'.pdf') as pdf:
        for 指标组 in 指标集合:
            序号 = 0
            fig = plt.figure(figsize=[10, 6.18])
            ax = fig.add_subplot(111)
            ax2 = ax.twinx()
            for 指标 in 指标组:
                print(指标)
                if '率' in 指标:
                    ax2.plot(x, varDict[指标][varDict[指标].iloc[:, 0] == gp].iloc[0, 1:], 线型[序号], label=指标)
                    ax2.legend(loc='center left')
                    ax2.set_ylabel(r"比率")
                else:
                    ax.plot(x, varDict[指标][varDict[指标].iloc[:, 0] == gp].iloc[0, 1:], 线型[序号], label=指标)
                    ax.legend(loc='upper left')
                    ax.grid()
                序号=序号+1
            ax.set_xlabel("时间")
            ax.set_ylabel(r"元")
            plt.title(指标列举原因[指标原因序号], fontsize='large',loc='center', fontweight='bold',color='blue',wrap=True,bbox=dict(facecolor='g', edgecolor='blue', alpha=0.65))  # 设置字体大小与格式
            指标原因序号=指标原因序号+1
            #plt.show()
            pdf.savefig()  # saves the current figure into a pdf page
            plt.close()


'''货币资金=pd.read_excel(dir_replace+r'货币资金.xls',header=0,index_col=0)#header=0,index_col=0表示将第一行和第一列作为索引，
应收票据=pd.read_excel(dir_replace+r'应收票据.xls',header=0,index_col=0)#header=0,index_col=0表示将第一行和第一列作为索引，
应收账款=pd.read_excel(dir_replace+r'应收账款.xls',header=0,index_col=0)#header=0,index_col=0表示将第一行和第一列作为索引，
营业总收入=pd.read_excel(dir_replace+r'营业总收入.xls',header=0,index_col=0)#header=0,index_col=0表示将第一行和第一列作为索引，
经营活动产生的现金流量净额=pd.read_excel(dir_replace+r'经营活动产生的现金流量净额.xls',header=0,index_col=0,dtype= str)#header=0,index_col=0表示将第一行和第一列作为索引，
投资活动产生的现金流量净额=pd.read_excel(dir_replace+r'投资活动产生的现金流量净额.xls',header=0,index_col=0,dtype= str)#header=0,index_col=0表示将第一行和第一列作为索引，
筹资活动产生的现金流量净额=pd.read_excel(dir_replace+r'筹资活动产生的现金流量净额.xls',header=0,index_col=0,dtype= str)#header=0,index_col=0表示将第一行和第一列作为索引，'''


with PdfPages('multipage_pdf.pdf') as pdf:
    for gp in gp_daima:
        plt.figure(1)
        plt.plot(x, 货币资金[货币资金.iloc[:,0]==gp].iloc[0,1:],'--', label='货币资金')
        plt.plot(x, 应收票据[应收票据.iloc[:,0]==gp].iloc[0,1:], 'o-',label='应收票据')
        plt.plot(x, 应收账款[应收账款.iloc[:,0]==gp].iloc[0,1:], 's-',label='应收账款')
        #plt.plot(x, 营业总收入[营业总收入.iloc[:, 0] == gp].iloc[0, 1:], 'kx-', label='营业总收入')
        plt.title(gp, fontsize='large', fontweight = 'bold') #设置字体大小与格式
        plt.title(gp, color='blue')    #设置字体颜色
        #plt.title(gp, loc='left')    #设置字体位置
        #plt.title(gp, verticalalignment='bottom')    #设置垂直对齐方式
        #plt.title(gp, rotation=45)    #设置字体旋转角度
        plt.title(gp, bbox=dict(facecolor='g', edgecolor='blue', alpha=0.65))    #标题边框
        plt.legend(loc = 'upper left')
        plt.xlabel("年份")#x轴上的名字
        plt.ylabel("元")#y轴上的名字
        #plt.show()
        pdf.savefig()  # saves the current figure into a pdf page
        plt.close()
m=1

d = pdf.infodict()
d['Title'] = 'Multipage PDF Example'
d['Author'] = u'Jouni K. Sepp\xe4nen'
d['Subject'] = 'How to create a multipage pdf file and set its metadata'
d['Keywords'] = 'PdfPages multipage keywords author title subject'
d['CreationDate'] = datetime.datetime(2009, 11, 13)
d['ModDate'] = datetime.datetime.today()


pd.DataFrame.plot()
plt.rc('text', usetex=True)


zhibia=['货币资金', '应收票据', '应收账款']

lrb=r'G:\数据\福晶科技\lrb002222.csv'
zcfzb=r'G:\数据\福晶科技\zcfzb002222.csv'
xjllb=r'G:\数据\福晶科技\xjllb002222.csv'
#----------------导入csv文件---------------#
data_zcfzb=pd.read_csv(zcfzb,error_bad_lines=False,encoding='gbk',
                       header=0,index_col=0,dtype= str, na_filter=True)#header=0,index_col=0表示将第一行和第一列作为索引，
data_zcfzb=data_zcfzb.replace('--','0')#将所有'--'替换为零
data_zcfzb=data_zcfzb.replace(' ',np.nan)#将所有'空字符串'替换为Nan
data_zcfzb=data_zcfzb.dropna(axis = 0,how='all')#去掉全部为一行全部为NaN，删除该行，axis=0表示行
data_zcfzb=data_zcfzb.dropna(axis = 1)#去掉包含nan值的列，axis=1表示列


#data_zcfzb=data_zcfzb.replace('\'', '')#去掉数据中字符串所有引号
#data_lrb.replace('"', '')  #去掉数据中字符串所有引号
x=data_zcfzb.columns
x=list(map(lambda x: x[2:4], x))

data_zcfzb=data_zcfzb.astype('float')



货币资金4=data_zcfzb.iloc[0,:]
应收票据4=data_zcfzb.iloc[5,:]
应收账款4=data_zcfzb.iloc[6,:]
货币票据账款合计4=货币资金4+应收票据4+应收账款4
流动负债合计4=data_zcfzb.iloc[83,:]
#货币资金4=list(map(eval,货币资金4))#现将所有非字符串变为字符串，而后再去掉。
货币资金与流动负债比例=货币资金4/流动负债合计4
应收票据与流动负债比例=应收票据4/流动负债合计4
应收账款与流动负债比例=应收账款4/流动负债合计4
货币票据账款合计与流动负债比例4=货币票据账款合计4/流动负债合计4

plt.figure(1)
plt.plot(x, 货币资金4,'-', label='货币资金4')
plt.plot(x, 应收票据4, 'o-',label='应收票据4')
plt.plot(x, 应收账款4, 's-',label='应收账款4')
plt.legend(loc = 'upper right')
plt.xlabel("年份")#x轴上的名字
plt.ylabel("万元")#y轴上的名字

plt.figure(2)

plt.plot(x, 货币资金与流动负债比例, '--', label='货币资金与流动负债比例')
plt.plot(x, 应收票据与流动负债比例, 's-', label='应收票据与流动负债比例')
plt.plot(x, 应收账款与流动负债比例, '*-', label='应收账款与流动负债比例')
plt.plot(x, 货币票据账款合计与流动负债比例4, 'd-' ,label='货币票据账款合计与流动负债比例4')
plt.legend(loc = 'upper right')
plt.show()
m=1
#--------------第一季度--------------#

#--------------第二季度--------------#

#--------------第三季度--------------#

#--------------年度报告--------------#
