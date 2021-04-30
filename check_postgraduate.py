from xlutils.copy import copy
import xlrd
import os,sys,re
filepath=r'E:\教务处工作\2020年推免（含学科提升计划专项推免）相关材料提交通知\学科提升计划各学院生成绩表\附件3-1教育信息技术学院推荐免试攻读硕士学位研究生综合得分表(学科提升计划类).xls';
sheet_name=u'验收通过汇总表'
row=1
col=1
book = xlrd.open_workbook(filepath)
#read excel
sheet = book.sheet_by_name(sheet_name)
sheet_value=sheet.cell(row,col).value
#write excel
write_book = copy(book)
write_sheet = write_book.get_sheet(sheet_name)
write_sheet.write(row+1,col+1,'ll')


write_book.save(filepath)