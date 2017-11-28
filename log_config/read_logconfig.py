import logging
from logging.handlers import RotatingFileHandler
import threading
import configparser

import sys,os
os.chdir(sys.path[0])

class LogSignleton(object):
    def __init__(self, log_config):
        pass
    def __new__(cls, log_config):
        mutex=threading.Lock()
        mutex.acquire() # 上锁，防止多线程下出问题
        if not hasattr(cls, '_instance'):
            cls._instance = super(LogSignleton, cls).__new__(cls) #super继承父类LogSignleton
            config = configparser.ConfigParser()
            config.read(log_config,encoding='utf-8')
            #读取logconfig.ini文件
            cls._instance.log_filename = config.get('LOGGING','log_file') #config['LOGGING']['log_file']
            cls._instance.max_bytes_each = int(config['LOGGING']['max_bytes_each'])
            cls._instance.backup_count = int(config.get('LOGGING','backup_count'))
            cls._instance.fmt = config.get('LOGGING','fmt')
            cls._instance.log_level_in_console = int(config.get('LOGGING','log_level_in_console'))
            cls._instance.log_level_in_logfile = int(config.get('LOGGING','log_level_in_logfile'))
            cls._instance.logger_name = config.get('LOGGING','logger_name')
            cls._instance.console_log_on = int(config.get('LOGGING','console_log_on'))
            cls._instance.logfile_log_on = int(config.get('LOGGING','logfile_log_on'))
            cls._instance.logger = logging.getLogger(cls._instance.logger_name)
            cls._instance.__config_logger()


        mutex.release()
        return cls._instance

    def get_logger(self):
        return  self.logger

    def __config_logger(self):
        # 设置日志格式
        fmt = self.fmt.replace('|','%')#  将 | 替换成%
        formatter = logging.Formatter(fmt)

        if self.console_log_on == 1: # 如果开启控制台日志
            console = logging.StreamHandler()
            console.setFormatter(formatter)
            self.logger.addHandler(console)
            self.logger.setLevel(self.log_level_in_console)

        if self.logfile_log_on == 1: # 如果开启文件日志

            rt_file_handler = RotatingFileHandler(self.log_filename, maxBytes=self.max_bytes_each, backupCount=self.backup_count)
            rt_file_handler.setFormatter(formatter)
            self.logger.addHandler(rt_file_handler)
            self.logger.setLevel(self.log_level_in_logfile)


path = os.getcwd()
#parent_path = os.path.dirname(path)
log_file = path + "\log_config\logconfig.ini"
logsignleton = LogSignleton(log_file)

loggers = logsignleton.get_logger()
