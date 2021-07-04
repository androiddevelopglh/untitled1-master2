import xlrd
import sys
import os
import xlwt
import numpy as np
import pandas as pd

def bankname(file_path,col1,bank_path, col2 ):

	file_excel = pd.read_excel(file_path, header=0, index_col=None,nrows=88)  # header=0表
	#bank_excel = pd.read_excel(bank_path, header=0, index_col=None)  # header=0表
	#bank_excel.to_pickle('samples')
	bank_excel=pd.read_pickle('samples')
	index = 0
	for namef in file_excel.iloc[:, col1]:
		#print(index)
		sim=0
		simstr=''
		#print(namef)
		for nameb in bank_excel.iloc[:, col2]:
			res = []
			for x in namef:
				if x in nameb:
					res.append(x)
			lenstr=len(res)
			if lenstr>sim:
				sim=lenstr
				simstr=nameb
		file_excel.loc[index, '匹配后银行']=simstr
		file_excel.loc[index, '相同字个数'] = sim
		#print(simstr)
		#print(sim)
		index=index+1
	return file_excel

file_path=r'G:\教务处工作\远程实习工作坊\2020年\【总表】校外劳务申报录入表.xls'
bank_path=r'G:\我的坚果云\学生助理工作\银行支行列表20210618(1)(1).xlsx'

file=bankname(file_path,7,bank_path,1)
file.to_excel(r'G:\教务处工作\远程实习工作坊\2020年\【总表】校外劳务申报录入表匹配银行后2.xls')
m=-1



def open_excel():
    try:
        book = xlrd.open_workbook(r'E:\教务处工作\教指委申报\2019-2023年省教指委委员推荐汇总表教指委汇总表excel.xls');  #文件名，把文件与py文件放在同一目录下
    except:
        print("open excel file failed!")
    try:
        sheets=book.sheet_names()
        sheet = book.sheet_by_name(sheets[0])   #execl里面的worksheet1
        return sheet
    except:
        print("locate worksheet in excel failed!")

book = xlrd.open_workbook(r'E:\教务处工作\19年校级质量工程\result.xls');  # 文件名，把文件与py文件放在同一目录下
workbook = xlwt.Workbook(encoding = 'ascii');
worksheet = workbook.add_sheet('My Worksheet1')


sheet = book.sheet_by_name('文件读取记录表')  # execl里面的worksheet1
sheet=open_excel()
for i in range(1,73):
	listexcel = [];
	result=0;
	for j in range(2,31,2):
		print(j)
		listexcel.append(sheet.cell(i,j).value)
		if sheet.cell(i,j+1).value=='' or sheet.cell(i,j+1).value=='否':
			result=result+1
	listexcel.sort()
	row_mean = np.mean(listexcel[2:13])
	worksheet.write(i, 0, row_mean)
	worksheet.write(i, 1, result)
workbook.save('Excel_Workbook1.xls')
aa=sheet.cell(1,0).value
m=1;


