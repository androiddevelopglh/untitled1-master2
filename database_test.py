# conding=utf8mb4
import pymysql
import xlrd
import sys
import os
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


# 建立数据库连接
try:
	connect = pymysql.Connect(
	host='localhost',
	port=3306,
	user='root',
	passwd='romance1126'
)
except:
	print("创建链接失败")

cursor = connect.cursor()
print('获取光标完成')

cursor.execute("create database if not exists test_db character set utf8;")
print('创建test_db库完成')
cursor.execute("use test;")
print('进入test库完成')
cursor.execute("show tables;")
sql_createTb = """CREATE TABLE if not exists MONEY (
                 name_id INT NOT NULL AUTO_INCREMENT,
                 LAST_NAME  CHAR(20),
                 AGE INT,
                 SEX CHAR(1),
                 PRIMARY KEY(name_id))
                 """
cursor.execute("create table if not exists test_tab(name1 char,name char(20))character set utf8;")
print('创建test_tab表完成')


def search_count():
	cursor = connect.cursor()
	select = "select count(id) from XXXX"  # 获取表中xxxxx记录数
	cursor.execute(select)  # 执行sql语句
	line_count = cursor.fetchone()
	print(line_count[0])


def insert_deta():
	cursor = connect.cursor()
	sheet = open_excel()
	row_data1=[];
	row_data1 = sheet.row_values(0);# 按行获取excel的值
	#row_data = sheet.sheets()[0]  # 按行获取excel的值
   #读取excel中的第一行作为数据库的field
	list_field=[];
	for j in range(0, sheet.ncols):
		list_field.append(str(row_data1[j]));
	str_field_type = ' char(20),'.join(list_field);
	print(str_field_type);
	#sql_createTb1 = 'CREATE TABLE if not exists excel_db1 ('+str(row_data[1])+' char(20),'+str(row_data[2]) + ' char(20))';
	cursor.execute("drop table if exists excel_db")
	sql_createTb1 = 'CREATE TABLE if not exists excel_db('+str_field_type+' char(20))';
	print(sql_createTb1);
	cursor.execute(sql_createTb1+"character set utf8;")
	connect.commit()
	for i in range(1, sheet.nrows):  # 第一行是标题名，对应表中的字段名所以应该从第二行开始，计算机以0开始计数，所以值是1
		excel_data=sheet.row_values(i);
		data2str=[]
		symbol=[];
		for j in range(0, sheet.ncols):
			#print(excel_data[j]);
			if isinstance(excel_data[j],float):
				cursor.execute("ALTER table excel_db modify "+row_data1[j]+" float(20)");
				#excel_data[j]=str(excel_data[j]);
				symbol.append('%f');
			else:
			   symbol.append('%s');
		connect.commit()
		symbol_str=','.join(symbol)
		#data=','.join(excel_data);
		#data.replace(',','\',\'');
		#print(data);
		str_field = ','.join(list_field);
		#insertdata="INSERT INTO excel_db("+str_field+") values("+symbol_str+")";
	#sql = "INSERT INTO XXX(name,data)VALUES(%s,%s)"
		#insertdata = "INSERT INTO excel_db(序号) values(%f)";

		#cursor.execute(insertdata, [1.0]); # 执行sql语句
		sql = "INSERT INTO EMPLOYEE(FIRST_NAME, \
		       LAST_NAME, AGE, SEX, INCOME) \
		       VALUES (%s, %s, %s, %s, %s )" % \
		      ('Mac', 'Mohan', 20, 'M', 2000)

		#cursor.execute("insert into user(id,age,name,create_time,update_time)values('%d','%d','%s','%s','%s')" % (user.getId(), user.getAge(), user.getName(), dt, dt))

		query = "INSERT INTO excel_db (序号,教指委名称) VALUES ('%f','%s')"%tuple([1.2,'fajsdlf']);
	cursor.execute(query);  # 执行sql语句
	connect.commit()
insert_deta()
print("ok ")




#cursor = connect.cursor()
#cursor.execute('create database if not exists ' + 'excel2database')
#  创建数据表的sql 语句  并设置name_id 为主键自增长不为空
sql_createTb = """CREATE TABLE if not exists MONEY (
                 name_id INT NOT NULL AUTO_INCREMENT,
                 LAST_NAME  CHAR(20),
                 AGE INT,
                 SEX CHAR(1),
                 PRIMARY KEY(name_id))
                 """
# 插入一条数据到moneytb 里面。
sql_insert = "insert into money(LAST_NAME,AGE,SEX) values('de1',18,'0')"
cursor.execute('drop database if exists ' + 'dbcreate')
cursor.execute('SHOW CREATE database excel2database')
# 在 execute里面执行SQL语句
cursor.execute(sql_createTb)
cursor.execute(sql_insert)
print(cursor.rowcount)
connect.commit()

connect.close()
cursor.close()

'''## 插入数据
sql = "INSERT INTO trade (name, account, saving) VALUES ( '%s', '%s', %.2f )"
data = ('雷军', '13512345678', 10000)
cursor.execute(sql % data)
connect.commit()
print('成功插入', cursor.rowcount, '条数据')

# 修改数据
sql = "UPDATE trade SET saving = %.2f WHERE account = '%s' "
data = (8888, '13512345678')
cursor.execute(sql % data)
connect.commit()
print('成功修改', cursor.rowcount, '条数据')

# 查询数据
sql = "SELECT name,saving FROM trade WHERE account = '%s' "
data = ('13512345678',)
cursor.execute(sql % data)
for row in cursor.fetchall():
    print("Name:%s\tSaving:%.2f" % row)
print('共查找出', cursor.rowcount, '条数据')

# 删除数据
sql = "DELETE FROM trade WHERE account = '%s' LIMIT %d"
data = ('13512345678', 1)
cursor.execute(sql % data)
connect.commit()
print('成功删除', cursor.rowcount, '条数据')

# 事务处理
sql_1 = "UPDATE trade SET saving = saving + 1000 WHERE account = '18012345678' "
sql_2 = "UPDATE trade SET expend = expend + 1000 WHERE account = '18012345678' "
sql_3 = "UPDATE trade SET income = income + 2000 WHERE account = '18012345678' "

try:
    cursor.execute(sql_1)  # 储蓄增加1000
    cursor.execute(sql_2)  # 支出增加1000
    cursor.execute(sql_3)  # 收入增加2000
except Exception as e:
    connect.rollback()  # 事务回滚
    print('事务处理失败', e)
else:
    connect.commit()  # 事务提交
    print('事务处理成功', cursor.rowcount)

# 关闭连接
'''
