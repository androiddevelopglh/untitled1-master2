#__author__ = 'Joker'
# -*- coding:utf-8 -*-
import urllib
import os
import os.path
import zipfile
from zipfile import *
import sys
import unrar
from unrar import rarfile
#import reload

#reload(sys)
#sys.setdefaultencoding('gbk')

#rootdir = "E:\表格下载"    # 指明被遍历的文件夹
#zipdir = "E:\解压测试文件夹"# 存储解压缩后的文件夹

#Zip文件处理类
class ZFile(object):
    def __init__(self, filename, mode='r', basedir=''):
        self.filename = filename
        self.mode = mode
        if self.mode in ('w', 'a'):
            self.zfile = zipfile.ZipFile(filename, self.mode, compression=zipfile.ZIP_DEFLATED)
        else:
            self.zfile = zipfile.ZipFile(filename, self.mode)
        self.basedir = basedir
        if not self.basedir:
            self.basedir = os.path.dirname(filename)

    def addfile(self, path, arcname=None):
        path = path.replace('//', '/')
        if not arcname:
            if path.startswith(self.basedir):
                arcname = path[len(self.basedir):]
            else:
                arcname = ''
        self.zfile.write(path, arcname)

    def addfiles(self, paths):
        for path in paths:
            if isinstance(path, tuple):
                self.addfile(*path)
            else:
                self.addfile(path)

    def close(self):
        self.zfile.close()

    def extract_to(self, path):
        for p in self.zfile.namelist():
            self.extract(p, path)



    def extract(self, filename, path):
        if not filename.endswith('/'):
            f = os.path.join(path, filename)
            fileext = GetFileNameAndExt(filename)[1]
            if fileext== ".zip":
                try:
                    ZFile(filename).extract_to(path)
                    os.remove(filename)
                except:
                    print(f,'解压失败')
            dir = os.path.dirname(f)
            if not os.path.exists(dir):
                os.makedirs(dir)
            self.zfile.extract(filename,path)

            #file(f, 'wb').write(read(filename))

    '''zFile = zipfile.ZipFile("F:\\txt.zip", "r")
    #ZipFile.namelist(): 获取ZIP文档内所有文件的名称列表
    for fileM in zFile.namelist():
        zFile.extract(fileM, "F:\\work")
    zFile.close();'''

#创建Zip文件
def createZip(zfile, files):
    z = ZFile(zfile, 'w')
    z.addfiles(files)
    z.close()

#解压缩Zip到指定文件夹
def extractZip(zfile, path):
    z = ZFile(zfile)
    z.extract_to(path)
    z.close()

#解压缩rar到指定文件夹
def extractRar(zfile, path):
    try:
        rar = rarfile.RarFile(zfile)
        rar.extractall(path)
        os.remove(zfile)
        print("rar OK.")
    except:
        print(zfile,"rar Error")

#获得文件名和后缀
def GetFileNameAndExt(filename):
    (filepath,tempfilename) = os.path.split(filename);
    (shotname,extension) = os.path.splitext(tempfilename);
    return shotname,extension

#定义文件处理数量-全局变量
fileCount = 0

#递归获得rar文件集合
def getFiles(filepath,zipdir):
#遍历filepath下所有文件，包括子目录
  files = os.listdir(filepath)
  for fi in files:
    fi_d = os.path.join(filepath,fi)
    if os.path.isdir(fi_d):
        getFiles(fi_d,zipdir)
    else:
        global fileCount
        fileCount = fileCount + 1
        # print fileCount
        fileName = os.path.join(filepath,fi_d)
        filenamenoext = GetFileNameAndExt(fileName)[0]
        fileext = GetFileNameAndExt(fileName)[1]
        # 如果要保存到同一个文件夹，将文件名设为空
        filenamenoext = ""
        zipdirdest = zipdir + "/" + filenamenoext + "/"
        if fileext in ['.zip','.rar']:
            if not os.path.isdir(zipdirdest):
                os.mkdir(zipdirdest)#如果没有该文件夹就创建一个
        if fileext == ".zip" :#
            print(str(fileCount) + " -- " + fileName)
           # unzip(fileName,zipdirdest)
            extractZip(fileName,zipdirdest)
        elif fileext == ".rar":
            print(str(fileCount) + " -- " + fileName)
            extractRar(fileName, zipdirdest)

#递归遍历“rootdir”目录下的指定后缀的文件列表
#getFiles(rootdir,rootdir)