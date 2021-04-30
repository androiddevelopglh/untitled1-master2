#! encoding:utf8
'''
环境：
    Win10 64位 Python 2.7.5
参考：
    http://www.pythonclub.org/python-network-application/email-format
    http://blog.sina.com.cn/s/blog_4deeda2501016eyf.html
'''


import imaplib
import email


def parseHeader(message):
    """ 解析邮件首部 """
    subject = message.get('subject')
    h = email.Header.Header(subject)
    dh = email.Header.decode_header(h)
    # 主题
    print(subject)
    print('</br>')
    # 发件人
    print('From:', email.utils.parseaddr(message.get('from'))[1])
    print('</br>')
    # 收件人
    print('To:', email.utils.parseaddr(message.get('to'))[1])
    print('</br>')
    # 抄送人
    print('Cc:',email.utils.parseaddr(message.get_all('cc'))[1])



def parseBody(message):
    """ 解析邮件/信体 """
    # 循环信件中的每一个mime的数据块
    for part in message.walk():
        # 这里要判断是否是multipart，是的话，里面的数据是一个message 列表
        if not part.is_multipart():
            charset = part.get_charset()
            # print 'charset: ', charset
            contenttype = part.get_content_type()
            # print 'content-type', contenttype
            name = part.get_param("name") #如果是附件，这里就会取出附件的文件名
            if name:
                # 有附件
                # 下面的三行代码只是为了解码象=?gbk?Q?=CF=E0=C6=AC.rar?=这样的文件名
                fh = email.Header.Header(name)
                fdh = email.Header.decode_header(fh)
                fname = dh[0][0]
                print('附件名:', fname)
                # attach_data = par.get_payload(decode=True) #　解码出附件数据，然后存储到文件中

                # try:
                #     f = open(fname, 'wb') #注意一定要用wb来打开文件，因为附件一般都是二进制文件
                # except:
                #     print '附件名有非法字符，自动换一个'
                #     f = open('aaaa', 'wb')
                # f.write(attach_data)
                # f.close()
            else:
                #不是附件，是文本内容
                print(part.get_payload(decode=True)) # 解码出文本内容，直接输出来就可以了。
                # pass
            # print '+'*60 # 用来区别各个部分的输出


def getMail(host, username, password, port=143):
    try:
        serv = imaplib.IMAP4_SSL(host, port)
    except:
        serv = imaplib.IMAP4(host, port)

    serv.login(username, password)
    serv.select()
    # 搜索邮件内容
    typ, data = serv.search(None, '(FROM "xx@xxx.com")')

    count = 1
    pcount = 1
    for num in data[0].split()[::-1]:
        typ, data = serv.fetch(num, '(RFC822)')
        text = data[0][1]
        message = email.message_from_string(text)   # 转换为email.message对象
        parseHeader(message)
        print('</br>')
        parseBody(message)
        pcount += 1
        if pcount > count:
            break

    serv.close()
    serv.logout()


if __name__ == '__main__':
    host = "imap.126.com" # "pop.mail_serv.com"
    username = "scnujyk@126.com"
    password = "85217673jyk"
    getMail(host, username, password)
