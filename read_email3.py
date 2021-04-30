
import os
import email
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8')

def mail_to_text(mailname,datapath,index):
#由于批处理的邮件包含的附件名称相同，这里传入一个index作为区别符
    fp=open(mailname,"r")
    msg=email.message_from_file(fp)

    for par in msg.walk():
        if not par.is_multipart():
            name=par.get_param("name")  #获取附件名

            if name:
                h=email.Header.Header(name)
                dh=email.Header.decode_header(h)
                fname = dh[0][0]

                data=par.get_payload(decode=True)

                try:
                    f=open(fname,'wb')
                except:
                    data_name=str(h).replace('/','_')  #附件数据
                    f=open(datapath+'\\'+str(index)+data_name,'wb')
                f.write(data)
                f.close()

if __name__=='__main__':
    dir="D:\\M\\"    #邮件存放路径
    #Dir=unicode(dir,"utf8")
    Dir=dir;
    datapath="D:\\M\\"   #附件存放路径
    #DataPath=unicode(datapath,"utf8")
    DataPath=datapath;
    count=0


    for filename in os.listdir(dir):
        print(filename)
        #filename=unicode(dir+filename,"utf8")   #由于邮件名出现中文，所以统一用utf8编码，便于读取
        filename = dir + filename
        print(filename)
        count+=1
        mail_to_text(filename,datapath,count)