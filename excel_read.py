import xlrd
import sys
import os
import xlwt
import numpy as np

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


