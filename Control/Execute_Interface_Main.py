from Control.read_path import read_path
import xlrd
from log_config.read_logconfig import loggers
from View.TaskFrame import *
from Control.read_file import read_file
from Control.Ordered_Dict import Ordered_Dict
from Control.Report_excel import Report_excel
from Control.Interface_Test import Interface_Test
from Control.connet_mysql import DB
from Control.SendEmail import Sendmail
import time
import os

class Execute_Interface_Main():
    def __init__(self,selected_task,parent):
        self.read_excel = read_file()
        self.testcase_sheet = self.read_excel.read_case_file()
        self.interconf_sheet = self.read_excel.read_conf_file()
        self.task_sheet = self.read_excel.read_task()
        self.project_list = eval(selected_task[3])  # 获取执行接口名称列表
        self.test_result = 'Pass'
        self.test_time = ''
        self.parent = parent
        self.inter_entrance(selected_task)

    def inter_entrance(self,selected_task):

        self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "开始循环接口，找对应的接口配置参数\n",('b'))
        self.parent.show_log.update()
        self.parent.show_log.see("end")
        loggers.info("开始循环接口，找对应的接口配置参数")
        self.thread_poject_name(selected_task)

    def thread_poject_name(self,selected_task):

        self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "接口名称是'%s'\n"%self.project_list)
        self.parent.show_log.update()
        self.parent.show_log.see("end")
        loggers.info("接口名称是%s"%self.project_list)
        # 进入执行项目
        self.test_case_func(selected_task)

    def test_case_func(self,selected_task):
        # 记录开始测试时间
        start_time = time.time()
        # 创建测试报告对象
        test_report = Report_excel(selected_task,self.parent)
        # 执行初始化SQL
        if selected_task[5].replace("\n","") not in ['','\n']:
            db = DB(self.parent)
            result = db.run(selected_task[5])
            if isinstance(result,(dict,list)):
                self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "任务初始化SQL语句执行成功\n")
                self.parent.show_log.update()
                self.parent.show_log.see("end")
                loggers.info("任务初始化SQL语句执行成功")
            else:
                loggers.info("任务初始化SQL语句执行失败%s"%result)
                showerror(title="错误",message="任务初始化SQL语句执行失败%s"%result)
                return

        # 创建发送请求对象
        test_object = Interface_Test(test_report,selected_task,self.parent)
        # 对象调用发送请求方法
        case_to_report,case_row_list = test_object.test_selected_case()
        # 执行任务结束SQL
        if selected_task[6].replace("\n","") not in ['','\n']:
            db = DB(self.parent)
            result = db.run(selected_task[6])
            if isinstance(result,(dict,list)):
                self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "任务结束SQL语句执行成功\n")
                self.parent.show_log.update()
                self.parent.show_log.see("end")
                loggers.info("任务结束SQL语句执行成功")
            else:
                loggers.info("任务结束SQL语句执行失败%s"%result)
                showerror(title="错误",message="任务结束SQL语句执行失败%s"%result)
                return

        end_time = time.time()
        self.test_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())

        test_report.time_caculate(end_time - start_time)

        test_report_file,report_dir = test_report.crate_excel(case_to_report,selected_task[0]+'.xlsx',case_row_list)
        if sorted(eval(selected_task[4])) != sorted(test_report.Success_case_ID):
            self.test_result = "Fail"

        if os.path.exists(report_dir) and selected_task[7] == '是':
            #进入发邮件模块
            SendEmail_object = Sendmail()
            SendEmail_object.email(report_dir,test_report_file)
