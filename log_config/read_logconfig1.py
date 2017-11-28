import logging
from logging.handlers import RotatingFileHandler
import threading
import configparser
import sys,os
#from Control.read_path import read_path
#from Control.read_file import read_file
import time
import sys,os

os.chdir(sys.path[0])

class LogSignleton():
    def __init__(self,log_config):
        #self.read_file = read_file()
        pass

    def __new__(cls, log_config):
        mutex=threading.Lock()
        mutex.acquire() # 上锁，防止多线程下出问题
        if not hasattr(cls, '_instance'):
            cls._instance = super(LogSignleton, cls).__new__(cls)                   #super继承父类LogSignleton
            config = configparser.ConfigParser()
            config.read(log_config,encoding='utf-8')
            #读取logconfig.ini文件
            cls._instance.log_filename = config.get('LOGGING','log_file')
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
        return self.logger

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

    def newlog_task_id(self):                  # 计算新增任务递增ID
        #case_list = self.read_file.read_task().col_values(0)[1:]
        case_list = []
        if case_list != []:
            task_id_ori = max(case_list)
            task_id_ori = task_id_ori.split('k')
            task_id_ori[1] = int(task_id_ori[1]) + 1
            task_id_ori[1] = "%04d" % int(task_id_ori[1])  #前面补0，保留5位
            task_id_new = 'Task' + task_id_ori[1]
        else:
            task_id_new = 'Task0001'
        return task_id_new

    def _set_result_filedir(self):
        file_name = self.newlog_task_id()
        self.file_name = file_name
        #ws_path = read_path("/pingtest/LOG")
        self.log_dir = ".\\Log\\"   #ws_path.read_superior_path()#获取文件路径

        if os.path.exists(self.log_dir) == False:
            os.makedirs(self.log_dir)
            print("自动建立路径：%s" %self.log_dir)
        if '' == self.file_name:
            pass
            print('filename can not be empty')
        # 判断是否为目录
        elif os.path.isdir(self.log_dir):
            parent_path, ext = os.path.splitext(file_name)
            tm = time.strftime('%Y%m%d%H%M%S', time.localtime())
            self.file_name = parent_path + "_" + tm + '.txt'
            os.chdir(self.log_dir)  # 转换目录
            return self.file_name
        else:
            pass
            #raise loggers.error("%s must point to a dir" % self.log_dir)

    def create_log(self):
        file_name = self._set_result_filedir()
        print('测试log文件名：%s ' % file_name)

        return file_name


path = os.getcwd()
#parent_path = os.path.dirname(path)
log_file = path + "\log_config\logconfig.ini"
logsignleton = LogSignleton(log_file)

loggers = logsignleton.get_logger()









