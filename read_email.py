import poplib
import email
import os
import ssl
import zip
import shutil
import fun_read_word
from xlutils.copy import copy
import xlrd
import time
from datetime import datetime
import sys

from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr

ssl._create_default_https_context = ssl._create_unverified_context
path_file=r'E:\教务处工作\19年校级质量工程';
def decode_str(s):
	value, charset = decode_header(s)[0]
	if charset:
		if charset == 'gb2312':
			charset = 'gb18030'
		value = value.decode(charset)
	return value


def get_email_headers(msg):
	headers = {}
	for header in ['From', 'To', 'Cc', 'Subject', 'Date']:
		value = msg.get(header, '')
		if value:
			if header == 'Date':
				headers['Date'] = value
			if header == 'Subject':
				subject = decode_str(value)
				headers['Subject'] = subject
			if header == 'From':
				hdr, addr = parseaddr(value)
				name = decode_str(hdr)
				from_addr = u'%s <%s>' % (name, addr)
				headers['From'] = from_addr
			if header == 'To':
				all_cc = value.split(',')
				to = []
				for x in all_cc:
					hdr, addr = parseaddr(x)
					name = decode_str(hdr)
					to_addr = u'%s <%s>' % (name, addr)
					to.append(to_addr)
				headers['To'] = ','.join(to)
			if header == 'Cc':
				all_cc = value.split(',')
				cc = []
				for x in all_cc:
					hdr, addr = parseaddr(x)
					name = decode_str(hdr)
					cc_addr = u'%s <%s>' % (name, addr)
					cc.append(to_addr)
				headers['Cc'] = ','.join(cc)
	return headers


def get_email_content(message, savepath):
	attachments = []
	for part in message.walk():
		filename = part.get_filename()
		if filename:
			filename = decode_str(filename)
			data = part.get_payload(decode=True)
			abs_filename = os.path.join(savepath, filename)
			attach = open(abs_filename, 'wb')
			attachments.append(filename)
			attach.write(data)
			attach.close()
	return attachments
# 取出附件中的文件正文内容
'''def get_file(msg):
	for part in msg.walk():
		filename = part.get_filename()
		if filename != None:  # 如果存在附件
			filename = decode_str(filename)  # 获取的文件是乱码名称，通过一开始定义的函数解码
			data = part.get_payload(decode=True)  # 取出文件正文内容
			# 此处可以自己定义文件保存位置
			path = filename
			f = open(path, 'wb')
			f.write(data)
			f.close()
			print(filename, 'download')'''
# 获取邮件的字符编码，首先在message中寻找编码，如果没有，就在header的Content-Type中寻找
def guess_charset(msg):
	charset = msg.get_charset()
	if charset is None:
		content_type = msg.get('Content-Type', '').lower()
		pos = content_type.find('charset=')
		if pos >= 0:
			charset = content_type[pos + 8:].strip()
	return charset
#邮件正文方法一
def parseBody(message):
    """ 解析邮件/信体 """
    # 循环信件中的每一个mime的数据块
    email='';
    for part in message.walk(): 
        # 这里要判断是否是multipart，是的话，里面的数据是一个message 列表
        if not part.is_multipart():
            #charset = part.get_charset()
            # print 'charset: ', charset
            contenttype = part.get_content_type()
            # print 'content-type', contenttype
            name = part.get_param("name") #如果是附件，这里就会取出附件的文件名
            if not name:
                charset = guess_charset(part)
                #print(charset)
                try:
                #不是附件，是文本内容
                    if charset==None and part.get_payload(decode=True)!=b'':
                        print(part.get_payload(decode=True))
                        email.append(part.get_payload(decode=True))
                        return email;
                    else :
                        print(part.get_payload(decode=True).decode(charset)) # 解码出文本内容，直接输出来就可以了。
                        email.append(part.get_payload(decode=True).decode(charset))
                        return email
                except:
                    print('未知编码或者空邮件')
                # pass
            # print '+'*60 # 用来区别各个部分的输出
#邮件正文方法二
def get_content(msg):
	for part in msg.walk():
		content_type = part.get_content_type()
		charset = guess_charset(part)
		# 如果有附件，则直接跳过
		if part.get_filename() != None:
			continue
		email_content_type = ''
		content = ''
		if content_type == 'text/plain':
			email_content_type = 'text'
		elif content_type == 'text/html':
			print('html 格式 跳过')
			continue  # 不要html格式的邮件
			email_content_type = 'html'
		if charset:
			try:
				content = part.get_payload(decode=True).decode(charset)
			except AttributeError:
				print('type error')
			except LookupError:
				print("unknown encoding: utf-8")
		if email_content_type == '':
			continue
		# 如果内容为空，也跳过
		print(email_content_type + ' -----  ' + content)
def rename(Subject,text,path,date):#检查所下载附件是否所需附件，如果不是则直接删除临时文件夹；否者将临时文件夹中的文件重命名，并且将临时文件夹修改后的名字返回in，
	for fpathe, dirs, files in os.walk(path):
		for file in files:
			print(os.path.join(fpathe, file))
			if '质量工程' in Subject and '.doc'in file:
				fun_read_word.read_word(os.path.split(os.path.split(path)[0])[0],os.path.join(fpathe, file),date)
if __name__ == '__main__':
	# 账户信息
	email = 'scnujyk@126.com'
	password = '85217673jyk'
	pop3_server = 'pop.126.com'
	# 连接到POP3服务器，带SSL的:
	server = poplib.POP3_SSL(pop3_server)
	# 可以打开或关闭调试信息:
	server.set_debuglevel(0)
	# POP3服务器的欢迎文字:
	print(server.getwelcome())
	# 身份认证:
	server.user(email)
	server.pass_(password)
	# stat()返回邮件数量和占用空间:
	msg_count, msg_size = server.stat()
	print('message count:', msg_count)
	print('message size:', msg_size, 'bytes')
	# b'+OK 237 174238271' list()响应的状态/邮件数量/邮件占用的空间大小
	resp, mails, octets = server.list()

	#for i in range(1, msg_count):
	for i in range(msg_count, 1,-1):
		print(i)
		try:
			resp, byte_lines, octets = server.retr(i)
		except:
			continue;
		# 转码
		str_lines = []
		for x in byte_lines:
			try:
				str_lines.append(x.decode())
			except:
				continue
		# 拼接邮件内容
		msg_content = '\n'.join(str_lines)
		# 把邮件内容解析为Message对象
		msg = Parser().parsestr(msg_content)
		headers = get_email_headers(msg)
		if not '质量工程' in headers['Subject']:
			continue
		date = time.mktime(time.strptime(headers['Date'][5:25].rstrip(), '%d %b %Y %H:%M:%S'))
		date_begin = time.mktime(time.strptime('1 Apr 2019 10:11:36', '%d %b %Y %H:%M:%S'))
		if date-date_begin<0:
			os._exit(0)
		# 判断邮件是否读取过
		book = xlrd.open_workbook(r'E:\教务处工作\19年校级质量工程\result.xls')
		write_book = copy(book)
		write_sheet = write_book.get_sheet(u'邮件读取记录')
		sheet_sta = book.sheet_by_name(u'邮件读取记录')
		nrows_sta = sheet_sta.nrows  # 行数
		flag2=''
		for rownum in range(0, nrows_sta):
			if sheet_sta.cell(rownum, 0).value == headers['Subject'] and sheet_sta.cell(rownum, 1).value == headers['Date']:
				flag2='该邮件已阅读'
				break
		if flag2=='该邮件已阅读':
			continue
		write_sheet.write(nrows_sta, 0, headers['Subject'])
		write_sheet.write(nrows_sta, 1, headers['Date'])
		write_book.save(r'E:\教务处工作\19年校级质量工程\result.xls')
		if not os.path.isdir(path_file+'\邮件附件'):#在目标文件夹下创建临时文件夹，用来存放某个邮件的附件，而后将附件全部解压，并将临时文件夹重命名。
			os.mkdir(path_file+'\邮件附件')
		if not os.path.isdir(path_file + '\邮件附件\临时文件'):
			os.mkdir(path_file + '\邮件附件\临时文件')
		attachments = get_email_content(msg, path_file+'\邮件附件\临时文件')
		print('subject:', headers['Subject'])
		print('from:', headers['From'])
		print('to:', headers['To'])
		if 'cc' in headers:
			print('cc:', headers['Cc'])
		try:
			print('date:', headers['Date'])
		except:
			print('邮件头不包含日期')
		print('attachments: ', attachments)
		emailtext=parseBody(msg)
		zip.getFiles(path_file+'\邮件附件\临时文件', path_file+'\邮件附件\临时文件')#解压文件
		'''
		if attachments!=[]:
			rename(headers['Subject'],emailtext,path_file+'\邮件附件\临时文件',headers['Date'])
		if not os.path.isdir(path_file+'\邮件附件\临时文件'):#如果目标文件夹中的临时文件夹不存在，则可以将临时文件夹转到目标文件夹
			shutil.move(path_file+'\邮件附件\临时文件', path_file)
		if os.path.isdir(r'E:\教务处工作\19年校级质量工程\临时文件'):
			shutil.rmtree(r'E:\教务处工作\19年校级质量工程\临时文件')
		#shutil.move(r'E:\教务处工作\19年校级质量工程\临时文件', r'E:\教务处工作\19年校级质量工程\aa')#重命名文件夹
'''
		print('-----------------------------')
	server.quit()
