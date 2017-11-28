#!/usr/bin/env python
# -*- coding:utf-8 -*-

__author__ = 'feng'

import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import configparser
from log_config.read_logconfig import loggers
from Control.read_path import read_path

class Sendmail(object):
    def __init__(self):
        self.Email_ser = None
        self.Email_port = None
        self.username = None
        self.password = None
        self.to_list = None
        self.read_file()

    def read_file(self):
        #读取配置文件
        conf_path = read_path('\pingtest\Email_conf.ini')
        config = configparser.ConfigParser()
        # 从配置文件中读取数据
        config.read(conf_path.read_superior_path(), encoding='utf-8')
        self.Email_ser = config['Email_Service']['Email_ser']
        self.Email_port = int(config['Email_Service']['Email_port'])
        self.username = config['send_user_pwd']['username']
        self.password = config['send_user_pwd']['password']
        self.to_list = config['to_list']['to_list']

    def email(self,path_name,file_name):
        msg = MIMEMultipart()
        msg['From'] = formataddr(["PingTest接口自动化测试",'PingTest@kairutech.com'])
        msg['To'] = formataddr(["测试组",'test@kairutech.com'])
        msg['Subject'] = "PingTest接口自动化测试报告"
        message = "Dear all:\n" \
                  "    附件为PingTest接口自动化测试报告，此为自动发送邮件，请勿回复，谢谢！"
        msg.attach(MIMEText(message, 'plain', 'utf-8'))

        #构造附件4,xlsx类型附件
        re_file = path_name + '/' + file_name
        att1 = MIMEApplication(open(re_file,'rb').read())
        att1.add_header('Content-Disposition', 'attachment', filename=file_name)
        msg.attach(att1)

        '''if test_report_fail_file is not None:
            re_fail_fail = path_name + '/' + test_report_fail_file
            att2 = MIMEApplication(open(re_fail_fail,'rb').read())
            att2.add_header('Content-Disposition', 'attachment', filename=test_report_fail_file)
            msg.attach(att2)
        else:
            pass'''
        try:
            server = smtplib.SMTP(self.Email_ser,self.Email_port)
            server.login(self.username,self.password) #login()方法用来登录SMTP服务器
            #sendmail()方法就是发邮件，由于可以一次发给多个人，所以传入一个list，邮件正文是一个str，as_string()把MIMEText对象变成str
            server.sendmail(self.username,eval(self.to_list), msg.as_string())
            loggers.info("邮件发送成功!")
            server.quit()
        except smtplib.SMTPException as e:
            loggers.info("Error: 无法发送邮件!%s" %e)