#!/usr/bin/env python
# -*- coding:utf-8 -*-
import datetime
import sys
import numpy as np
import matplotlib.pyplot as plt
import linecache
import pandas as pd
import win32ui
from matplotlib.backends.backend_pdf import PdfPages
import textwrap
import time
import http.client, urllib.parse
import json
import fun_readallfile

###--------合并excel表格--------####

#def


dir=r'G:\教务处工作\师范生学业荣誉制度\2021年上半年\收集的名单'
typeall=['.xls','.xlsx']
file_list=fun_readallfile.check_file(dir,typeall)





m=1


#短信发送功能
def sendmessage(Message,allPhone,SendDate='',SendTime=''):
    Len_Phone=int(len(allPhone)/1000)+1
    Message = u'【华南师大教务处】'+Message
    Message = Message.encode('gb2312')
    httpClient = None
    i=0
    while i<Len_Phone:
        Phone1000=allPhone[i:i+1000]
        Phone = ";".join(Phone1000)
        try:
            params = urllib.parse.urlencode({'UserID': '837086',
                                       'Account': 'jsjyk',
                                       'Password': 'DC4584ABD7E06A1FC928DF01C37B7E48BEA257EE',
                                       'Content': Message,
                                       'Phones': Phone,
                                       'SendDate': SendDate,
                                       'SendTime': SendTime,
                                       'ReturnXJ': '1'})

            headers = {"Content-type": "application/x-www-form-urlencoded",
                       "Accept": "text/plain"}

            httpClient = http.client.HTTPConnection("dxjk.51lanz.com", 80, timeout=30)
            httpClient.request("POST", "/LANZGateway/DirectSendSMSs.asp", params, headers)
            response = httpClient.getresponse()
            result=json.loads(response.read())
            #print(response.status)
            #print(response.reason)
            #print(response.read())
            #print(response.getheaders())
            #json.loads(response.read())
            return result
        except Exception as e:
            #print(e)
            return e
        finally:
            if httpClient:
                httpClient.close()
        i=i+1

#-----------短信回复查询功能-------------------
def Fetchmessage():
    try:
        params = urllib.parse.urlencode({'UserID': '837086',
                                   'Account': 'jsjyk',
                                   'Password': 'DC4584ABD7E06A1FC928DF01C37B7E48BEA257EE',
                                   'ReturnXJ': '1'})

        headers = {"Content-type": "application/x-www-form-urlencoded",
                   "Accept": "text/plain"}

        httpClient = http.client.HTTPConnection("dxjk.51lanz.com", 80, timeout=30)
        httpClient.request("POST", "/LANZGateway/DirectFetchSMS.asp", params, headers)
        response = httpClient.getresponse()
        result=json.loads(response.read())
        return result
        NN=1
    except Exception as e:
        # print(e)
        return e
    finally:
        if httpClient:
            httpClient.close()



Message='您好，这是测试信息'
Phone=['18818399096','18392508007']
#result=sendmessage(Message,Phone)#即时发送
#result=sendmessage(Message,Phone,'2021-06-23','15:53:00')#定时发送
#result=Fetchmessage()#查询结果




##--------------常见分类信息--------------------
varDict = locals()
联系人类别=['所有学院','师范专业','省赛','田家炳','东芝杯']
所有学院=['教育科学学院', '历史文化学院', '哲学与社会发展学院', '马克思主义学院', '马院（思政理论课）', '外国语言文化学院',
       '大英部', '教育信息技术学院', '数学科学学院', '地理科学学院', '计算机学院', '生命科学学院', '美术学院',
       '心理学院', '旅游管理学院', '政治与公共管理学院', '体育科学学院', '文学院', '物理与电信工程学院',
       '信息光电子科技学院', '化学学院', '环境学院', '音乐学院', '经济与管理学院', '经济与管理学院', '法学院',
       '软件学院', '国际商学院', '城市文化学院', '职业教育学院']
师范专业=['文学院','数学科学学院','外国语言文化学院','物理与电信工程学院', '化学学院','生命科学学院',
           '哲学与社会发展学院','历史文化学院', '地理科学学院','音乐学院','体育科学学院', '美术学院',
          '教育科学学院', '教育信息技术学院',  '马克思主义学院', '马院（思政理论课）', '计算机学院', '心理学院']
省赛=['文学院','数学科学学院','外国语言文化学院','物理与电信工程学院', '化学学院','生命科学学院',
        '哲学与社会发展学院','历史文化学院', '地理科学学院','音乐学院', '美术学院', '教育科学学院', '教育信息技术学院', '计算机学院', '心理学院']
田家炳=['文学院','数学科学学院','外国语言文化学院','物理与电信工程学院', '化学学院','生命科学学院',
           '哲学与社会发展学院','历史文化学院', '地理科学学院']
东芝杯=['数学科学学院','物理与电信工程学院', '化学学院']

dlg = win32ui.CreateFileDialog(1) # 1表示打开文件对话框
dlg.SetOFNInitialDir('G:\\通讯录') # 设置打开文件对话框中的初始显示目录
dlg.DoModal()
file_path = dlg.GetPathName() # 获取选择的文件名称
Dread_excel=pd.read_excel(file_path, header=0, index_col=None, usecols=list(range(1, 100)), skiprows=2)#header=0表
Dread_excel.loc[:, ['校区', '单位']]= Dread_excel.loc[:, ['校区', '单位']].fillna(method='ffill')
Dread_excel.set_index(["单位"], inplace=True, drop=False)#设定单位为index，并且令drop为
for 类别 in 联系人类别:
    try:
        tempdata = Dread_excel.loc[varDict[类别], :]
        tempdata=tempdata.loc[:,['姓名','手机号码','工作邮箱','校区','单位','办公电话']]#
        tempdata.to_excel('G:\我的坚果云\联系方式'+'\\'+类别 + '教学院长联系方式.xls',sheet_name = "联系方式",index = False)
    except:
        print(类别 + '出错')
m=1





