# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

#-*- encoding: gb2312 -*-
import os
import poplib
from email import parser  
from datetime import datetime,timedelta

#%%功能函数
#检查数据目录，下载附件
def download_AttachmentFilename(msg,path):
    if (msg.is_multipart()):
        # 如果邮件对象是一个MIMEMultipart,
        # get_payload()返回list，包含所有的子对象:
        parts = msg.get_payload()
        for n, part in enumerate(parts):
            download_AttachmentFilename(part,path)
    else:
        # 邮件对象不是一个MIMEMultipart,
        # 当content_type为application/octet-stream时，就判断为附件
        content_type = msg.get_content_type()
        if content_type=='application/octet-stream':
            AttachmentFilename = msg.get_filename()
            key=AttachmentFilename.split('_')[2]
            #"DYZ3","DYZ5","RYZ3","RYZ4"为卫星关键字，若附件名相应的字段与之无关则不再考虑
            if key not in ["DYZ3","DYZ5","RYZ3","RYZ4"]:
                return
            subdir=AttachmentFilename.split('_')[-1][0:10]
            path_sub=os.path.join(path,key,subdir)
            #若不存在该日期的子目录则创建目录
            if not os.path.exists(path_sub):
                os.mkdir(path_sub)
            path_file=os.path.join(path_sub,AttachmentFilename)
            #若日期子目录不存在和附件同名的文件，则将附件下载到对应的目录下
            if not os.path.exists(path_file):
                fileData=msg.get_payload(decode=True)
                f=open(path_file,'wb')
                f.write(fileData)
                f.close()            

#获取邮件日期
def get_emailTime(i):
    header = pp.top(i,0)
    header_info = '\r\n'.join(header[1])
    msg = parser.Parser().parsestr(header_info)
    return datetime.strptime(msg.get("Date")[5:24],'%d %b %Y %H:%M:%S')
    
#%%主进程
#数据路径        
path=r'D:/data'
#%%登录邮箱
#邮箱服务器地址
host = 'pop.cstnet.cn'
#邮箱用户名及密码
username = 'liuxy14@radi.ac.cn'
password = 'liuxueying123'

pp = poplib.POP3(host)
pp.set_debuglevel(1)
pp.user(username)
pp.pass_(password)
#%%获取邮件总数和当前时间
email_num=pp.stat()[0]
time_now=datetime.now()
#设置搜索邮件的时间间隔，默认搜索前5分钟
time_D=timedelta(0,60*5,0)

for i in range(email_num):
    #获取邮件的接收时间
    email_time=get_emailTime(email_num-i)
    #若邮件接收时间超出搜索搜索范围，则此邮件及之前的邮件不再考虑
    if time_now-email_time>time_D:
        break
    #否则根据附件名及数据目录情况决定是否下载该附件
    else:
        messages = pp.retr(email_num-0)
        msg = '\r\n'.join(messages[1])
        msg = parser.Parser().parsestr(msg)
        AttachmentFilename=download_AttachmentFilename(msg,path)
pp.quit()
