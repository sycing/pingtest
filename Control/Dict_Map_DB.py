#!/usr/bin/env python
# -*- coding:utf-8 -*-
__author__ = 'feng'
import re
from Control.connet_mysql import DB
from log_config.read_logconfig import loggers

class Dict_Map_DB():

    def __init__(self,parent):

        self.json_params = {}
        self.db = DB(parent)

    def Dict_Map_DB_sum(self,json_params,Expect_Type):
        if isinstance(json_params,dict):
            self.json_params = self.dict_map_db(json_params,Expect_Type)
        else:
            loggers.error("您的期望结果既不是json也不是SQL查询语句，请修改！")

        return self.json_params

    def dict_map_db(self,json_params,Expect_Type):

        for key,value in json_params.items():
            if isinstance(value,dict):
                for key_i,value_i in value.items():
                    value[key_i] = self.replace(value[key_i],Expect_Type)
                self.dict_map_db(value,Expect_Type)
            else:
                json_params[key] = self.replace(json_params[key],Expect_Type)

        return json_params


    def replace(self,value,Expect_Type):
        if isinstance(value,list):
            for index in range(0,len(value)):
                if type(value[index]) != int and re.match('SELECT', str(value[index]), flags=0):
                    value[index] = self.db.search(value[index],Expect_Type)
        elif type(value) != int and re.match('SELECT', str(value), flags=0):
            value = self.db.search(value,Expect_Type)
        #self.db.close()
        return value