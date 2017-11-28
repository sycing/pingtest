#!/usr/bin/env python
# -*- coding:utf-8 -*-
__author__ = 'feng'
from log_config.read_logconfig import loggers
import random
from random import choice as randomChoice

class Dict_Radom():

    def __init__(self):

        self.json_params = {}
        self.passData = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm',
            'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z',
            'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
            'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
            '1', '2', '3', '4', '5', '6', '7', '8', '9', '0']
        self.passData_num = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0']

    def radom_replace_sum(self,json_params):
        if isinstance(json_params,dict):
            self.json_params = self.radom_replace(json_params)
        else:
            while -1 != json_params.find("*"):
                json_params = json_params.replace("*",randomChoice(self.passData),1)
            while -1 != json_params.find("$"):
                json_params = json_params.replace("$",randomChoice(self.passData_num),1)
            self.json_params = json_params

        return self.json_params

    def radom_replace(self,json_params):

        for key,value in json_params.items():
            if isinstance(value,dict):
                for key_i,value_i in value.items():
                    value[key_i] = self.replace(value[key_i])
                self.radom_replace(value)
            else:
                json_params[key] = self.replace(json_params[key])

        return json_params

    def replace(self,value):
        if isinstance(value,list):
            for index in range(0,len(value)):
                if isinstance(value[index],str):
                    while -1 != value[index].find("*"):
                        value[index] = value[index].replace("*",randomChoice(self.passData),1)
                    while -1 != value[index].find("$"):
                        value[index] = value[index].replace("$",randomChoice(self.passData_num),1)
                elif isinstance(value[index],dict):
                    self.radom_replace(value[index])
                else:
                    pass

        elif isinstance(value,(int,float)):
            pass
            return value
        elif value is None:
            pass
            return value
        elif isinstance(value,dict):
            self.radom_replace(value)
        else:
            while -1 != value.find("*"):
                value = value.replace("*",randomChoice(self.passData),1)
            while -1 != value.find("$"):
                value = value.replace("$",randomChoice(self.passData_num),1)
        return value
