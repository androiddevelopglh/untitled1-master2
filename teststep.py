import pymysql
import xlrd
import sys
def open_excel():
    try:
        book = xlrd.open_workbook(r'E:\教务处工作\教指委申报\2019-2023年省教指委委员推荐汇总表教指委汇总表excel.xls');  #文件名，把文件与py文件放在同一目录下
    except:
        print("open excel file failed!")
    try:
        sheet = book.sheet_by_name("Sheet2")   #execl里面的worksheet1
        return sheet
    except:
        print("locate worksheet in excel failed!")


sheet = open_excel()
a=1;
b=['aadf','dafadf','daaf']
b.append('daf')
print(b)
