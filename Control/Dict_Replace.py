#!/usr/bin/env python
# -*- coding:utf-8 -*-
__author__ = 'feng'
from log_config.read_logconfig import loggers
import json

class Dict_Replace():

    def __init__(self):

        self.json_params = {}
        self.replace_value_dict = {}

    def body_replace(self,json_params,replace_value):
        #将replace_value 转化成字符串
        try:
            self.json_params = eval(str(json_params))
            self.replace_value_dict = eval(str(replace_value))
        except:
            loggers.error(u'传入参数为非字典格式')
        for key,value in self.replace_value_dict.items():
            if isinstance(value,dict):
                for key_i,value_i in value.items():
                    for ke,va in self.json_params.items():
                        if str(ke) == str(key_i) and (self.json_params[ke] == "#"):
                            self.json_params[ke] = value[key_i]
                            break
                        else:
                            continue
            else:
                for ke,va in self.json_params.items():
                    if isinstance(va,dict):
                        for k,v in va.items():
                            if str(k) == str(key) and (va[k] == "#"):
                                va[k] = self.replace_value_dict[key]
                                break
                            else:
                                continue
                    elif str(ke) == str(key) and (self.json_params[ke] == "#"):
                        self.json_params[ke] = self.replace_value_dict[key]
                        break
                    else:
                        continue

        return self.json_params

