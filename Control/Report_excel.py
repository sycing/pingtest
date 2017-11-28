#!/usr/bin/env python
# -*- coding:utf-8 -*-
__author__ = 'Feng'
import xlsxwriter
import os
import time
from log_config.read_logconfig import loggers
from Control.read_path import read_path
from View.TaskFrame import *

class Report_excel(object):
    def __init__(self, selected_task,parent):
        self.parent = parent
        self.title = 'PingTest接口自动化测试报告'    # 网页标签
        self.file_name = ''                # report 文件名称
        self.file_name_fail = ''           # 失败report 文件名称
        self.Test_time = '00:00:00'        # 测试总耗时
        self.Success_num = 0               # 测试成功的用例数
        self.Fail_num = 0                  # 测试失败数
        self.error_num = 0                 # 运行出错用例数
        self.testCase_total = 0            # 运行测试用例总数
        self.block_num = 0                 # 未运行的测试用例总数
        self.task_id = ''
        self.case_list = []                # 执行的用例序号集合
        self.Fail_case_ID = []             # 失败用例数ID集合
        self.Success_case_ID = []          # 执行成功的用例ID
        self.Error_case_ID = []             # 失败用例数ID集合
        self.unblock_case_ID = []          # 实际执行的用例
        self.block_case_ID = []            # 未执行的用例ID
        self.report_dir = ''               # 测试报告路径
        self.selected_task = selected_task   # 执行任务数据

    def get_format(self,wd, option={}):
        return wd.add_format(option)
    # 设置居中
    def get_format_center(self,wd,num=1):
        return wd.add_format({'align': 'center','valign': 'vcenter','border':num})
    def get_format_left(self,wd,num=1):
        return wd.add_format({'align': 'left','valign': 'vcenter','border':num,'bold': True})
    # 写数据
    def _write_center(self,worksheet, cl, data, wd):
        return worksheet.write(cl, data, self.get_format_center(wd))
    def _write_left(self,worksheet, cl, data, wd):
        return worksheet.write(cl, data, self.get_format_left(wd))

    def init(self,worksheet):
        # 设置列行的宽高
        worksheet.set_column("A:A", 20)
        worksheet.set_column("B:B", 10)
        worksheet.set_column("C:C", 20)
        worksheet.set_column("D:D", 10)
        worksheet.set_column("E:E", 30)
        worksheet.set_column("F:F", 70)
        for i in [2,3,4,5,7,8,9]:
            worksheet.set_row(i,25)

        define_format_H1 = self.get_format(self.workbook, {'bold': True, 'font_size': 18,'border':True,'align':'center'})
        define_format_H2 = self.get_format(self.workbook, {'bold': True, 'font_size': 14,'border':True,'align':'center','bg_color':'#4682B4','font_color': '#ffffff'})
        define_format_H3 = self.get_format(self.workbook,{'align': 'left','valign': 'vcenter','border':True,'text_wrap':True})
        define_format_H4 = self.get_format(self.workbook,{'align': 'center','valign': 'vcenter','border':True,'text_wrap':True,'font_color': 'green'})
        define_format_H5 = self.get_format(self.workbook,{'align': 'center','valign': 'vcenter','border':True,'text_wrap':True,'font_color': 'orange'})
        define_format_H5_error = self.get_format(self.workbook,{'align': 'center','valign': 'vcenter','border':True,'text_wrap':True,'font_color': 'red'})
        define_format_H6 = self.get_format(self.workbook,{'align': 'center','valign': 'vcenter','border':True,'text_wrap':True})
        # Create a new Chart object.

        worksheet.merge_range('A1:F1', self.title, define_format_H1)
        worksheet.merge_range('A2:F2', '测试概况', define_format_H2)

        self._write_left(worksheet, "A3", '执行任务ID', self.workbook)
        self._write_left(worksheet, "A4", '用例总数', self.workbook)
        self._write_left(worksheet, "A5", '实际执行用例总数', self.workbook)
        self._write_left(worksheet, "A6", '测试总耗时', self.workbook)

        self._write_left(worksheet, "A8", '成功用例数(pass)', self.workbook)
        self._write_left(worksheet, "A9", '失败用例数(fail)', self.workbook)
        self._write_left(worksheet, "A10", '错误用例数(error)', self.workbook)

        self._write_left(worksheet, "C3", '任务名称', self.workbook)
        self._write_left(worksheet, "C4", '任务描述', self.workbook)
        self._write_left(worksheet, "C5", '执行接口', self.workbook)
        self._write_left(worksheet, "C6", '执行用例ID', self.workbook)
        self._write_left(worksheet, "C8", '执行用例成功率（%）', self.workbook)
        self._write_left(worksheet, "C9", '执行用例失败率（%）', self.workbook)
        self._write_left(worksheet, "C10", '执行用例错误率（%）', self.workbook)

        self._write_left(worksheet, "E8", '成功用例数集合(list)', self.workbook)
        self._write_left(worksheet, "E9", '失败用例数集合(list)', self.workbook)
        self._write_left(worksheet, "E10", '错误用例数集合(list)', self.workbook)

        data = {"task_id":self.task_id,"case_total": self.testCase_total+self.block_num, "testCase_total": self.testCase_total,"test_time":self.Test_time}
        worksheet.write(2, 1, data['task_id'], define_format_H6)
        worksheet.write(3, 1, data['case_total'], define_format_H6)
        worksheet.write(4, 1, data['testCase_total'], define_format_H6)
        worksheet.write(5, 1, data['test_time'], define_format_H6)

        data2 = {"Success_num": self.Success_num, "Fail_num": self.Fail_num,"Error_num":self.error_num}
        worksheet.write(7, 1, data2['Success_num'], define_format_H4)
        worksheet.write(8, 1, data2['Fail_num'], define_format_H5)
        worksheet.write(9, 1, data2['Error_num'], define_format_H5_error)

        data3 = {"Success_percentage": float(self.Success_num)/float(self.testCase_total),
                 "Fail_percentage": float(self.Fail_num)/float(self.testCase_total),
                 "Error_percentage": float(self.error_num)/float(self.testCase_total),
                 }
        worksheet.write(7, 3, '%.2f%%'% (data3['Success_percentage']*100), define_format_H4)
        worksheet.write(8, 3, '%.2f%%'% (data3['Fail_percentage']*100), define_format_H5)
        worksheet.write(9, 3, '%.2f%%'% (data3['Error_percentage']*100), define_format_H5_error)

        data4 = {"Success_case_ID": str(self.Success_case_ID), "Fail_case_ID": str(self.Fail_case_ID),"Error_case_ID": str(self.Error_case_ID),}
        worksheet.write(7, 5, data4['Success_case_ID'], self.get_format(self.workbook,{'align': 'left','valign': 'vcenter','border':True,'text_wrap':True,'font_color': 'green'}))
        worksheet.write(8, 5, data4['Fail_case_ID'], self.get_format(self.workbook,{'align': 'left','valign': 'vcenter','border':True,'text_wrap':True,'font_color': 'orange'}))
        worksheet.write(9, 5, data4['Error_case_ID'], self.get_format(self.workbook,{'align': 'left','valign': 'vcenter','border':True,'text_wrap':True,'font_color': 'red'}))

        data5 = {"任务名称":str(self.selected_task[1]),"任务描述":str(self.selected_task[2]),
                 "执行接口":str(self.selected_task[3]),"执行用例ID":str(self.case_list)}
        worksheet.merge_range('D3:F3', data5["任务名称"], define_format_H3)
        worksheet.merge_range('D4:F4', data5["任务描述"], define_format_H3)
        worksheet.merge_range('D5:F5', data5["执行接口"], define_format_H3)
        worksheet.merge_range('D6:F6', data5["执行用例ID"], define_format_H3)

        self.unblock_case_ID = self.Success_case_ID + self.Fail_case_ID
        self.unblock_case_ID.sort()

    def init_fail(self,worksheet_fail):

        # 设置列行的宽高
        worksheet_fail.set_column("A:A", 25)
        worksheet_fail.set_column("B:B", 125)
        worksheet_fail.set_column("C:C", 25)
        worksheet_fail.set_column("D:D", 25)
        worksheet_fail.set_column("E:E", 25)
        worksheet_fail.set_column("F:F", 25)

        define_format_H1 = self.get_format(self.workbook_fail, {'bold': True, 'font_size': 18,'border':True,'align':'center'})
        define_format_H2 = self.get_format(self.workbook_fail, {'bold': True, 'font_size': 14,'border':True,'align':'center','bg_color':'red','font_color': '#ffffff'})
        # Create a new Chart object.

        worksheet_fail.merge_range('A1:B1', self.selected_task[0] + '--失败测试报告总概况', define_format_H1)
        worksheet_fail.merge_range('A2:B2', '失败测试概况', define_format_H2)

        self._write_center(worksheet_fail, "A3", '失败用例数(fail)', self.workbook_fail)
        self._write_center(worksheet_fail, "A4", '失败用例数集合(list)', self.workbook_fail)
        self._write_center(worksheet_fail, "A5", '测试总耗时', self.workbook_fail)

        data = {"fail_total": self.Fail_num, "Fail_case_ID": str(self.Fail_case_ID),"test_time":self.Test_time}
        self._write_center(worksheet_fail, "B3", data['fail_total'], self.workbook_fail)
        self._write_center(worksheet_fail, "B4", data['Fail_case_ID'], self.workbook_fail)
        self._write_center(worksheet_fail, "B5", data['test_time'], self.workbook_fail)

    # 生成饼形图
    def pie(self,workbook, worksheet):
        chart1 = workbook.add_chart({'type': 'column'})
        chart1.add_series({
        'name':       '接口测试统计',
        'categories': "=" + self.selected_task[0] + "测试总况!$A$8:$A$10",
        'values':   "=" + self.selected_task[0] + "测试总况!$B$8:$B$10",
        })
        chart1.set_title({'name': '接口测试统计'})
        chart1.set_style(10)
        chart1.set_table()
        worksheet.insert_chart('A12', chart1, {'x_offset': 25, 'y_offset': 5})
    def pie_fail(self,workbook_fail, worksheet_fail):
        chart1 = workbook_fail.add_chart({'type': 'column'})
        chart1.add_series({
        'name':       u'接口测试统计',
        'categories': u"=" + self.selected_task[0] + u'失败测试总况!$A$3:$A$3',
        'values':    u"=" + self.selected_task[0] + u'失败测试总况!$B$3:$B$3',
        })
        chart1.set_title({'name': u'接口测试统计'})
        chart1.set_style(10)
        worksheet_fail.insert_chart('A9', chart1, {'x_offset': 25, 'y_offset': 5})
    def test_detail(self,sheet,worksheet,bg_color,workbook,case_row_list):

        # 设置列行的宽高
        worksheet.set_column("A:A", 8)
        worksheet.set_column("B:B", 10)
        worksheet.set_column("C:C", 15)
        worksheet.set_column("D:D", 15)
        worksheet.set_column("E:E", 8)
        worksheet.set_column("F:F", 20)
        worksheet.set_column("G:G", 15)
        worksheet.set_column("H:H", 30)
        worksheet.set_column("I:I", 8)
        worksheet.set_column("J:J", 100)
        worksheet.set_column("K:K", 30)
        worksheet.set_column("L:L", 8)
        worksheet.set_column("M:M", 100)
        worksheet.set_column("N:N", 15)
        worksheet.set_column("O:O", 15)
        worksheet.set_column("P:P", 15)
        worksheet.set_column("Q:Q", 35)
        worksheet.freeze_panes(2, 0)  # 设置窗口冻结区域  距离顶部2行，距离左边6列

        format_detail_title = self.get_format(workbook, {'bold': True, 'font_size': 18 ,'align': 'center','valign': 'vcenter','bg_color': bg_color, 'font_color': '#ffffff'})
        worksheet.merge_range('A1:R1', u'测试详情', format_detail_title)
        self._write_center(worksheet, "A2", u'用例类型', workbook)
        self._write_center(worksheet, "B2", u'所属模块', workbook)
        self._write_center(worksheet, "C2", u'接口名称', workbook)
        self._write_center(worksheet, "D2", u'用例ID', workbook)
        self._write_center(worksheet, "E2", u'请求方式', workbook)
        self._write_center(worksheet, "F2", u'请求地址', workbook)
        self._write_center(worksheet, "G2", u'用例名称', workbook)
        self._write_center(worksheet, "H2", u'包体body', workbook)
        self._write_center(worksheet, "I2", u'期望类型', workbook)
        self._write_center(worksheet, "J2", u'预期结果', workbook)
        self._write_center(worksheet, "K2", u'检查字段', workbook)
        self._write_center(worksheet, "L2", u'测试结果', workbook)
        self._write_center(worksheet, "M2", u'实际返回结果', workbook)
        self._write_center(worksheet, "N2", u'比对失败的字段', workbook)
        self._write_center(worksheet, "O2", u'期望写入DB结果', workbook)
        self._write_center(worksheet, "P2", u'实际写入DB结果', workbook)
        self._write_center(worksheet, "Q2", u'HTTPHeader头信息', workbook)
        self._write_center(worksheet, "R2", u'备注', workbook)

        format_detail = self.get_format(workbook,{'align': 'left','valign': 'vcenter','border':True,'text_wrap':True})
        format_detail_pass = self.get_format(workbook,{'align': 'left','valign': 'vcenter','border':True,'text_wrap':True,'bg_color':'green', 'font_color': '#ffffff'})
        format_detail_fail = self.get_format(workbook,{'align': 'left','valign': 'vcenter','border':True,'text_wrap':True,'bg_color':'orange', 'font_color': '#ffffff'})
        format_detail_error = self.get_format(workbook,{'align': 'left','valign': 'vcenter','border':True,'text_wrap':True,'bg_color':'red', 'font_color': '#ffffff'})
        temp = 2
        for index in case_row_list:
            for i in range(0,sheet.ncols):
                test_result = sheet.row_values(index)[11]
                if test_result == "Pass":
                    worksheet.write(temp, i, sheet.row_values(index)[i], format_detail)
                    worksheet.write(temp, 11, sheet.row_values(index)[11], format_detail_pass)
                elif test_result in ['Fail','Fail,响应包体与期望包体字段不一致','Fail,写入数据库结果不正确']:
                    worksheet.write(temp, i, sheet.row_values(index)[i], format_detail)
                    worksheet.write(temp, 11, sheet.row_values(index)[11], format_detail_fail)
                elif test_result == "Error":
                    worksheet.write(temp, i, sheet.row_values(index)[i], format_detail)
                    worksheet.write(temp, 11, sheet.row_values(index)[11], format_detail_error)
            temp = temp + 1

    def _set_result_filedir(self, file_name):
        self.file_name = file_name
        ws_path = read_path("/pingtest/Report")
        self.report_dir = ws_path.read_superior_path()  # 获取工作空间路径

        if os.path.exists(self.report_dir) is False:
            os.makedirs(self.report_dir)
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "已自动建立测试报告路径：%s\n" % self.report_dir)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info("已自动建立测试报告路径：%s" % self.report_dir)
        if '' == self.file_name:
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + 'filename can not be empty\n')
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.error('filename can not be empty')
        # 判断是否为目录
        elif os.path.isdir(self.report_dir):
            parent_path, ext = os.path.splitext(file_name)
            tm = time.strftime('%Y%m%d%H%M%S', time.localtime())
            self.file_name = parent_path + "_" + tm + ext
            self.file_name_fail = parent_path + 'Fail' + "_" + tm + ext
            os.chdir(self.report_dir)  # 转换目录
        else:
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "%s must point to a dir\n" % self.report_dir)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            raise loggers.error("%s must point to a dir" % self.report_dir)

    def time_caculate(self, seconds):
        self.Test_time = time.strftime('%H:%M:%S', time.gmtime(seconds))
        return self.Test_time

    def crate_excel(self,sheet,report_name,case_row_list):
        self._set_result_filedir(report_name)
        self.workbook = xlsxwriter.Workbook(self.file_name)
        self.worksheet = self.workbook.add_worksheet(self.selected_task[0] + "测试总况")
        self.worksheet2 = self.workbook.add_worksheet(self.selected_task[0] + "测试详情")
        self.init(self.worksheet)
        self.pie(self.workbook, self.worksheet)
        self.test_detail(sheet,self.worksheet2,"#4682B4",self.workbook,case_row_list)
        self.workbook.filename = self.file_name

        self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + '测试报告文件名：%s \n' % self.file_name,('c'))
        self.parent.show_log.update()
        self.parent.show_log.see("end")
        loggers.info('测试报告文件名：%s ' % self.file_name)
        self.workbook.close()
        path = os.getcwd()
        os.chdir(os.path.dirname(path))  # 转换目录到主程序目录
        return self.file_name,self.report_dir

    def crate_excel_fail(self,sheet,report_name):
        self._set_result_filedir(report_name)
        self.workbook_fail = xlsxwriter.Workbook(self.file_name_fail)
        self.worksheet_fail = self.workbook_fail.add_worksheet(self.selected_task[0] + u"失败测试总况")
        self.worksheet2_fail = self.workbook_fail.add_worksheet(self.selected_task[0] + u"失败测试详情")
        self.init_fail(self.worksheet_fail)
        self.pie_fail(self.workbook_fail, self.worksheet_fail)
        self.test_detail(sheet,self.worksheet2_fail,"red",self.workbook_fail,self.Fail_case_ID)
        self.workbook_fail.filename = self.file_name_fail
        self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + '测试报告文件名：%s \n' % self.file_name)
        self.parent.show_log.update()
        self.parent.show_log.see("end")
        loggers.info('测试报告文件名：%s ' % self.file_name_fail)
        self.workbook_fail.close()
        path = os.getcwd()
        os.chdir(os.path.dirname(path))  # 转换目录到主程序目录
        return self.file_name_fail,self.report_dir