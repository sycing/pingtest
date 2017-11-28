#!/usr/bin/env python
# -*- coding:utf-8 -*-

__author__ = 'feng'

import os
from log_config.read_logconfig import loggers

class read_path(object):
    def __init__(self,file):
        self.file = file
        #获取工程目录
    def read_superior_path(self):
        path = os.getcwd()
        parent_path = os.path.dirname(path)
        path_data_conf = parent_path + self.file
        if ((os.path.exists(path_data_conf))==False):
            loggers.error("请检查是否存在 : %s" %path_data_conf)
        return path_data_conf



