#!/usr/bin/env python
# -*- coding:utf-8 -*-
__author__ = 'feng'

class Dict_Map():

    def __init__(self):
        pass

    def dict_map_sum(self,old_Dict):

        self.new_keys = []
        self.new_values = []
        self.new_Dict = {}
        self.dict_map(old_Dict)
        return self.new_Dict

    def dict_map(self,old_Dict):
        for k,v in old_Dict.items():
            self.new_keys.append(k)
            self.new_values.append(v)
            self.new_Dict = dict(zip(self.new_keys,self.new_values))
            if isinstance(v,dict):
                self.dict_map(v)
            elif isinstance(v,list):
                for index in range(0,len(v)):
                    if isinstance(v[index],dict):
                        self.dict_map(v[index])