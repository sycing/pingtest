import requests,urllib.request,urllib,urllib.parse
import ssl
from log_config.read_logconfig import loggers
import json
import time
from View.TaskFrame import *

class interface(object):

    def __init__(self,url,head,body,parent):
        self.url = url
        self.head = head
        self.body = body
        self.parent = parent

    def post(self):
        try:
            #context = ssl._create_unverified_context()
            response = requests.post(self.url,json=eval(str(self.body)),headers=self.head)
            response.raise_for_status()
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "URL:%s\n"%self.url)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "header:%s\n"%self.head)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info("URL:%s"%self.url)
            loggers.info("header:%s"%self.head)
            loggers.info("code:%s"%response)
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "code:%s\n"%response,('b'))
            self.parent.show_log.update()
            self.parent.show_log.see("end")

            re = response.text
            re = json.loads(re)
            #self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "response:%s\n"%re)
            #self.parent.show_log.update()
            #self.parent.show_log.see("end")
            return re
        except Exception as e:
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + 'post error : %s\n' %e,('a'))
            self.parent.show_log.update()
            loggers.error('post error : %s' %e)

    def get(self):
        try:
            #context = ssl._create_unverified_context()
            self.body = urllib.parse.urlencode(self.body)
            data = requests.get(self.url,params=self.body,headers=self.head)
            data.raise_for_status()

            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "urlencode完整地址:%s\n"%data.url)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info(data.url)
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "header:%s\n"%self.head)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info("header:%s"%self.head)
            page_response = data.text
            page_response = json.loads(page_response)
            #self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "page_response:%s\n"%page_response)
            #self.parent.show_log.update()
            #self.parent.show_log.see("end")
            return page_response

        except Exception as e:
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + 'get error : %s\n' %e,('a'))
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info('get error : %s' %e)
            return str(e)

    def put(self):
        try:
            #context = ssl._create_unverified_context()
            response = requests.put(self.url,json=eval(str(self.body)),headers=self.head)
            response.raise_for_status()

            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "URL:%s\n"%self.url)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "header:%s\n"%self.head)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info("URL:%s"%self.url)
            loggers.info("header:%s"%self.head)
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "code:%s\n"%response,('b'))
            self.parent.show_log.update()
            self.parent.show_log.see("end")

            re = response.text
            re = json.loads(re)

            #loggers.info("response:%s"%re)
            #self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "response:%s\n"%re)
            #self.parent.show_log.update()
            #self.parent.show_log.see("end")
            return re
        except Exception as e:
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + 'put error : %s\n' %e,('a'))
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.error('put error : %s' %e)

    def delete(self):
        try:
            #context = ssl._create_unverified_context()
            data = requests.delete(self.url,headers=self.head)
            data.raise_for_status()

            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "URLenCode完整地址:%s\n"%data.url)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info(data.url)
            page_response = data.text
            page_response = json.loads(page_response)

            #self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "page_response:%s\n"%page_response)
            #self.parent.show_log.update()
            #self.parent.show_log.see("end")
            return page_response

        except Exception as e:
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + 'delete error : %s\n' %e,('a'))
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.error('delete error : %s' %e)
            return str(e)