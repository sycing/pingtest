import json
import time
import os
from Control.read_path import read_path
import xlrd
from Control.read_file import read_file
from Control.connet_mysql import DB
from Control.Dict_Map_DB import Dict_Map_DB
from Control.Dict_Replace import Dict_Replace
from Control.Dict_Map import Dict_Map
from Control.interface import interface
from Control.Dict_radom import Dict_Radom
from log_config.read_logconfig import loggers
from View.TaskFrame import *

class Interface_Test(object):
    def __init__(self,report,selected_task,parent):
        self.report = report
        self.parent = parent
        self.selected_task = selected_task
        self.read_excel = read_file()
        self.Auth_key = ''
        self.get_id = ''
        self.Auth_key_value = ''
        self.get_id_value = ''
        self.catch_result = {}
        self.catch_result_total = {}
        self.json_params = {}
        self.ndict = Dict_Map()
        self.rdict = Dict_Replace()
        self.radom_dict = Dict_Radom()
        self.db_result = {}
        self.parent_catch_result = {}
        self.dict_db = Dict_Map_DB(self.parent)

    def test_selected_case(self):
        self.testcase_sheet = self.read_excel.read_case_file()
        self.interconf_sheet = self.read_excel.read_conf_file()
        self.task_sheet = self.read_excel.read_task()
        self.report.task_id = self.selected_task[0]
        case_list = eval(self.selected_task[4])
        for i in eval(self.selected_task[4]):
            for j,k in enumerate(self.testcase_sheet.col_values(3)):
                if i == k and self.testcase_sheet.row_values(j)[15] == "无效":
                    case_list.remove(k)
                    break

        self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "测试用例ID列表：%s\n"%case_list)
        self.parent.show_log.update()
        self.parent.show_log.see("end")
        loggers.info("测试用例ID列表：%s"%case_list)
        if case_list is []:
            showerror(title='错误',message="用例ID为空，请检查任务数据！")
            return
        self.report.case_list = case_list

        case_to_report,case_row_list = self.index_range(case_list)
        # 计算未执行用例数和集合
        unblock_list = self.report.Success_case_ID + self.report.Fail_case_ID  # 已执行用例列表
        self.report.block_case_ID = list(set(case_list) ^ set(unblock_list))  # 所有用例列表与已执行列表的差集
        self.report.block_case_ID.sort()
        self.report.block_num = len(self.report.block_case_ID)
        return case_to_report,case_row_list

    def index_range(self,case_list):
        case_row_list = []
        try:
            for i in case_list:
                case_row_list.append(self.testcase_sheet.col_values(3).index(i))
        except BaseException as e:
            showerror(title="错误",message="测试用例中存在无效用例，请检查任务\n%s" %e)
            return

        for index in case_row_list:
            currunt_caseid = self.testcase_sheet.row_values(index)[3]
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "当前执行用例ID是：%s\n"%currunt_caseid,('c'))
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info("当前执行用例ID是：%s"%currunt_caseid)

            params = self.testcase_sheet.row_values(index)[7]
            Expect_Type = self.testcase_sheet.row_values(index)[8]
            expected_result = self.testcase_sheet.row_values(index)[9]

            # 调用http请求方法
            result,response,fail_keys = self.test_post(params,expected_result,index)
            time.sleep(0.1)
            if result == "Error":
                self.report.error_num = self.report.error_num + 1
                self.report.Error_case_ID.append(self.testcase_sheet.row_values(index)[3])
                self.testcase_sheet.put_cell(index,15,1,"",0)

            # 更新EXCEL对象数据
            self.testcase_sheet.put_cell(index,8,1,Expect_Type,0)
            # 记录运行结果
            self.testcase_sheet.put_cell(index,11,1,result,0)
            # 记录错误信息
            self.testcase_sheet.put_cell(index,12,1,str(response).replace("'",'''"'''),0)
            # 记录比对失败的字段名称
            self.testcase_sheet.put_cell(index,13,1,str(fail_keys),0)

            # 变更包体字段，用于打印报告
            # 将unicode转中文
            json_params = self.json_params
            self.testcase_sheet.put_cell(index,7,1,str(json_params).replace("'",'''"'''),0)
            # 测试用例数加1
            self.report.testCase_total = self.report.testCase_total + 1
        return self.testcase_sheet,case_row_list

    def test_post(self,params,expected_result,index):
        fail_keys = []
        response = None
        currunt_caseid = self.testcase_sheet.row_values(index)[3]
        Expect_Type = self.testcase_sheet.row_values(index)[8]

        # 检查期望结果是否为json格式
        try:
            expected_result = self.dict_db.Dict_Map_DB_sum(eval(expected_result.replace("\n"," ")),Expect_Type)
        except BaseException as e:
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "测试用例%s中填写的expected_result字段错误，不是json格式或执行SQL错误\n%s\n"%(currunt_caseid,e),("a"))
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            return "Error","期望结果字段错误，不是json格式或执行SQL错误",[]
        # 更新期望结果字段
        self.testcase_sheet.put_cell(index,9,1,str(expected_result).replace("'",'''"'''),0)

        # 判断catch_result字段是否为字典格式
        if self.testcase_sheet.row_values(index)[12].strip().replace("\n","") == '':
            pass
        else:
            try:
                self.catch_result = eval(self.testcase_sheet.row_values(index)[12].strip().replace("\n",""))
            except:
                self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "测试用例ID:%s 中填写的catch_result错误，不是字典格式\n"%currunt_caseid,('a'))
                self.parent.show_log.update()
                self.parent.show_log.see("end")
                loggers.error("测试用例中填写的catch_result错误，不是字典格式")
                return "Error","catch_result格式错误，不是字典格式",[]
        if isinstance(self.catch_result,dict):
            pass
        else:
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "测试用例ID:%s 中填写的catch_result错误，不是字典格式\n"%currunt_caseid,('a'))
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.error(u"测试用例中填写的catch_result错误，不是字典格式")
            return "Error","catch_result格式错误，不是字典格式",[]

        # 请求包体中替换流程列表中的字段
        self.json_params = params
        try:
            if isinstance(eval(self.json_params),dict):
                pass
            else:
                self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "当前用例ID:%s请求包体格式为非Json格式\n"%currunt_caseid,("a"))
                self.parent.show_log.update()
                self.parent.show_log.see("end")
                return "Error","请求包体格式错误，不是json格式",[]
        except BaseException as e:
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "当前用例ID:%s请求包体格式为非Json格式\n"%currunt_caseid,("a"))
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            return "Error","请求包体格式错误，不是json格式",[]

        # 替换请求包头字段
        if self.catch_result_total != {}:
            try:
                self.json_params = self.rdict.body_replace(self.json_params,self.catch_result_total)
            except:
                pass
            for key,value in self.catch_result_total.items():
                if value == "Auth":
                    self.Auth_key = key
                elif value == "URL":
                    self.get_id = key
                else:
                    continue

        # 替换请求包体中需要加入随机数的字段
        self.json_params = self.radom_dict.radom_replace_sum(self.json_params)

        if self.Auth_key != '' and self.parent_catch_result not in ('',None,{}):
            if self.Auth_key in self.parent_response.keys() and self.Auth_key in eval(self.parent_catch_result).keys():
                if eval(self.parent_catch_result)[self.Auth_key] == 'Auth':
                    self.Auth_key_value = self.parent_response[self.Auth_key]
            else:
                self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "前一个流程接口返回结果没有%s字段\n"%self.Auth_key)
                self.parent.show_log.update()
                self.parent.show_log.see("end")
                loggers.info("前一个流程接口返回结果没有%s字段\n"%self.Auth_key)
        if self.get_id != '' and self.get_id_value == '':
            try:
                self.get_id_value = self.parent_response[self.get_id]
            except:
                self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "当前用例ID:%s前置接口返回包体没有%s字段\n"%(currunt_caseid,self.get_id),("a"))
                self.parent.show_log.update()
                self.parent.show_log.see("end")
                return "Error","当前用例ID:%s前置接口返回包体没有%s字段\n"%(currunt_caseid,self.get_id),[]

        self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "替换后请求包体\n %s\n"%self.json_params)
        self.parent.show_log.update()
        self.parent.show_log.see("end")
        loggers.info("替换后请求包体\n %s"%self.json_params)
        header = {"Accept":"application/json","Content-Type":"application/json;charset=utf-8","Accept-Language": "zh-cn,zh;q=0.5"}
        URL = self.read_excel.read_case_file().row_values(index)[5]
        if -1 != URL.find("URL"):
            URL = URL.replace("URL",self.get_id_value)
        try:
            head_json = eval(self.read_excel.read_case_file().row_values(index)[16])

        except:
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "测试用例ID:%s HTTPHeader字段格式不正确，非Json格式\n"%currunt_caseid)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            return "Error","HTTPHeader字段格式不正确，不是json格式",[]
        head_json = dict(head_json,**header)
        if 'Authentication' in head_json.keys():

            if head_json['Authentication'] == "empty":
                head_json['Authentication'] = ''
            elif head_json['Authentication'] == "None":
                head_json['Authentication'] = None
            elif head_json['Authentication'] is None:
                head_json['Authentication'] = None
            elif head_json['Authentication'] == "":
                head_json['Authentication'] = self.Auth_key_value
            elif head_json['Authentication'] != "":
                pass

        head_json = self.radom_dict.radom_replace_sum(head_json)
        self.testcase_sheet.put_cell(index,5,1,URL,0)
        self.testcase_sheet.put_cell(index,16,1,str(head_json),0)
        test_interface = interface(URL,head_json,self.json_params,self.parent)
        request_Method = self.testcase_sheet.row_values(index)[4]

        if request_Method == 'POST':
            response = test_interface.post()
        elif request_Method == 'GET':
            response = test_interface.get()
        elif request_Method == 'PUT':
            response = test_interface.put()
        elif request_Method == 'DELETE':
            response = test_interface.delete()

        check_field_str = self.testcase_sheet.row_values(index)[10]
        check_field_remove = self.testcase_sheet.row_values(index)[11]

        if check_field_str == "ALL":
            self.ndict.dict_map_sum(expected_result)
            check_field_list = self.ndict.new_keys
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + '需要比对的字段列表: %s\n' %check_field_list)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info('需要比对的字段列表: %s' %check_field_list)
        elif check_field_str != "":
            try:
                check_field_list = eval(check_field_str)
            except:
                self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + '测试用例ID：%s 需要比对的字段为非列表格式,或包含不存在的字段\n' %currunt_caseid)
                self.parent.show_log.update()
                self.parent.show_log.see("end")
                return "Error","需要比对的字段格式不正确，不是列表格式,或包含不存在的字段",[]

            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + '需要比对的字段列表: %s\n' %check_field_list)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info('需要比对的字段列表: %s' %check_field_list)
        elif check_field_remove != "":
            self.ndict.dict_map_sum(expected_result)
            try:
                for i in eval(check_field_remove):
                    self.ndict.new_keys.remove(i)
                check_field_list = self.ndict.new_keys
            except:
                self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + '测试用例ID：%s 不需检查字段为非列表格式,或包含不存在的字段\n' %currunt_caseid)
                self.parent.show_log.update()
                self.parent.show_log.see("end")
                return "Error","不需检查字段格式不正确，不是列表格式,或包含不存在的字段",[]
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + '需要比对的字段列表: %s\n' %check_field_list)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info('需要比对的字段列表: %s' %check_field_list)
        else:
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + '比对字段Check_field和Check_field_remove都没有填写，请检查用例!\n',('a'))
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.error('比对字段Check_field和Check_field_remove都没有填写，请检查用例！')
            return "Error","比对字段Check_field和Check_field_remove都没有填写",[]
        self.testcase_sheet.put_cell(index,10,1,str(check_field_list),0)

        # 按照Execl标记的Check_field进行检查，以列表形式存在
        result = "Pass"
        result_response = None
        # 1.验证结果，如果返回None，则说明发送请求错误
        if isinstance(response,dict) is False:
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + 'ERROR：测试用例%s发送请求失败！%s\n' %(currunt_caseid,response),('a'))
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.error('ERROR：测试用例%s发送请求失败！%s' %(currunt_caseid,response))
            return 'Error','请求失败:%s'%response,[]
        else:
            # 调用字典遍历模块，方便找出所有的键值对
            new_response = self.ndict.dict_map_sum(response)
            new_expected = self.ndict.dict_map_sum(expected_result)
            self.parent_response = new_response
            self.parent_catch_result = self.testcase_sheet.row_values(index)[12].strip().replace("\n","")

        # 2.比对response和expected字典键的个数和键名称是否一致
        #if sorted(new_response.keys()) == sorted(new_expected.keys()):
            # 3.遍历响应包体字典的所有键值对，与期望值逐一检查
        for index_c in range(len(check_field_list)):
                field_name = check_field_list[index_c]
                try:
                    if new_response[field_name] == new_expected[field_name]:
                        result = 'Pass'
                        result_response = response
                        continue
                    else:
                        result = 'Fail'
                        result_response = response
                        fail_keys.append(field_name)
                        break
                except BaseException as e:
                    self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + 'ERROR：测试用例%s 比对结果时出现异常！%s\n' %(currunt_caseid,e),('a'))
                    self.parent.show_log.update()
                    self.parent.show_log.see("end")
                    loggers.error('ERROR：测试用例%s 比对结果时出现异常！%s %s' %(currunt_caseid,e,response))
                    return 'Error','比对结果时出现异常:%s %s'%(e,response),[]
        '''else:
            result = 'Fail,响应包体与期望包体字段不一致'
            result_response = response
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "响应包体与期望包体字段不一致\n")
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info("响应包体与期望包体字段不一致！")'''

        if result == "Pass":

            checkDB_SQL = self.testcase_sheet.row_values(index)[13]
            checkDB_result = self.testcase_sheet.row_values(index)[14]
            if checkDB_SQL.replace("\n","").strip() != "":
                db = DB(self.parent)
                db_result = db.run(checkDB_SQL)
                self.db_result = db_result
                if isinstance(db_result,(dict,list)) is False:
                    return 'Error','期望写入DB结果SQL语句执行失败:%s'%db_result,[]
                self.testcase_sheet.put_cell(index,15,1,str(db_result),0)
                if eval(checkDB_result) == db_result:
                    self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + '写入数据库后结果为%s，比对结果正确\n' %db_result)
                    self.parent.show_log.update()
                    self.parent.show_log.see("end")
                    loggers.info("写入数据库后结果为%s，比对结果正确"%str(db_result))
                    self.testcase_sheet.put_cell(index,14,1,json.dumps(eval(checkDB_result)).encode().decode("unicode-escape"),0)
                    self.testcase_sheet.put_cell(index,15,1,json.dumps(db_result).encode().decode("unicode-escape"),0)
                else:
                    self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + '写入数据库后结果为%s，比对结果不正确\n'%db_result,('a'))
                    self.parent.show_log.update()
                    self.parent.show_log.see("end")
                    loggers.info("写入数据库后结果为%s，比对结果不正确"%str(db_result))
                    result = "Fail,写入数据库结果不正确"
                    self.report.Fail_num = self.report.Fail_num + 1
                    self.report.Fail_case_ID.append(self.testcase_sheet.row_values(index)[3])
                    # 把结果替换给catch_result字段
                    catch_result_single = self.rdict.body_replace(self.catch_result,response)
                    self.catch_result_total = dict(catch_result_single,**self.catch_result_total)
                    if self.db_result != '' and isinstance(self.db_result,dict):
                        catch_result_single = self.rdict.body_replace(self.catch_result,self.db_result)  # 抓取查询DB的结果字段
                        self.catch_result_total = self.rdict.body_replace(self.catch_result_total,catch_result_single)
                    self.testcase_sheet.put_cell(index,14,1,json.dumps(eval(checkDB_result)).encode().decode("unicode-escape"),0)
                    self.testcase_sheet.put_cell(index,15,1,json.dumps(db_result).encode().decode("unicode-escape"),0)
            else:
                self.testcase_sheet.put_cell(index,15,1,"",0)

        if result == "Pass":

            self.report.Success_num = self.report.Success_num + 1
            self.report.Success_case_ID.append(self.testcase_sheet.row_values(index)[3])
            # 把结果替换给catch_result字段
            catch_result_single = self.rdict.body_replace(self.catch_result,response)
            self.catch_result_total = dict(catch_result_single,**self.catch_result_total)
            if self.db_result != '' and isinstance(self.db_result,dict):
                catch_result_single = self.rdict.body_replace(self.catch_result,self.db_result)  # 抓取查询DB的结果字段
                self.catch_result_total = self.rdict.body_replace(self.catch_result_total,catch_result_single)

            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + '流程接口可获取字典：%s\n' %self.catch_result_total)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info('流程接口可获取字典：%s' %self.catch_result_total)
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "成功用例数集合--%s\n" %self.report.Success_case_ID,('b'))
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info("成功用例数集合--%s" %self.report.Success_case_ID)

        elif result in ['Fail','Fail,响应包体与期望包体字段不一致']:

            self.report.Fail_num = self.report.Fail_num + 1
            self.report.Fail_case_ID.append(self.testcase_sheet.row_values(index)[3])
            # 把结果替换给catch_result字段
            catch_result_single = self.rdict.body_replace(self.catch_result,response)
            self.catch_result_total = dict(catch_result_single,**self.catch_result_total)
            if self.db_result != '' and isinstance(self.db_result,dict):
                catch_result_single = self.rdict.body_replace(self.catch_result,self.db_result)  # 抓取查询DB的结果字段
                self.catch_result_total = self.rdict.body_replace(self.catch_result_total,catch_result_single)
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + '流程接口可获取字典：%s\n' %self.catch_result_total)
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info('流程接口可获取字段：%s' %self.catch_result_total)
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "失败用例数集合--%s\n" %self.report.Fail_case_ID,('a'))
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info("失败用例数集合--%s" %self.report.Fail_case_ID)
            self.testcase_sheet.put_cell(index,15,1,"",0)

        self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "测试结果为：%s\n"%result)
        self.parent.show_log.update()
        self.parent.show_log.see("end")
        loggers.info("测试结果为：%s"%result)
        self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "比对失败字段：%s\n"%fail_keys)
        self.parent.show_log.update()
        self.parent.show_log.see("end")
        loggers.info("比对失败字段：%s"%fail_keys)
        return result,result_response,fail_keys
