from collections import OrderedDict

class Ordered_Dict():
    def __init__(self):

        self.column = OrderedDict()

    def test_case_header(self):
        # 定义表格Treeview
        self.column['case_type'] = '用例类型'
        self.column['model'] = '所属模块'
        self.column['interfacename'] = '接口名称'
        self.column['ID'] = '用例ID'
        self.column['Method'] = '请求方式'
        self.column['IPPort'] = '请求IP:Port'
        self.column['Case_Name'] = '用例名称'
        self.column['Params'] = '请求包体'
        self.column['Expect_Type'] = '期望结果类型'
        self.column['Expectation'] = '期望结果'
        self.column['Check_field'] = '检查字段'
        self.column['Check_field_remove'] = '不需检查字段'
        self.column['catch_result'] = '结果获取字段'
        self.column['checkDB_SQL'] = '检查写入数据库SQL'
        self.column['checkDB_result'] = '检查写入数据库结果'
        self.column['status'] = '用例状态'
        self.column['HTTPHeader'] = 'HTTPHeader'

        return self.column

    def interface_conf_header(self):

        # 定义表格Treeview
        self.column['case_type'] = '用例类型'
        self.column['model'] = '所属模块'
        self.column['interfacename'] = '接口名称'
        self.column['interfaceename'] = '接口英文名称'
        self.column['parent_interface'] = '前置接口英文名称'
        self.column['parent_caseNo.'] = '前置接口执行用例序号'
        self.column['requestType'] = '请求类型'
        self.column['host'] = '请求地址'
        self.column['Method'] = '请求方式'

        return self.column

    def task_header(self):

        # 定义表格Treeview
        self.column['task_id'] = '任务ID'
        self.column['task_name'] = '任务名称'
        self.column['task_description'] = '任务描述'
        self.column['task_interface'] = '执行接口'
        self.column['task_caseID'] = '执行用例ID'
        self.column['setup_SQL'] = '初始化SQL'
        self.column['teardown_SQL'] = '清除数据SQL'
        self.column['task_email_option'] = '邮件发送'
        self.column['task_statues'] = '任务状态'
        self.column['task_result'] = '执行结果'
        self.column['task_time'] = '执行时间'

        return self.column

    def inter_conf_header(self):

        # 定义表格Treeview
        self.column['case_type'] = '用例类型'
        self.column['model'] = '所属模块'
        self.column['interfacename'] = '接口名称'
        self.column['interfaceEname'] = '接口英文名称'
        self.column['requestType'] = '请求类型'
        self.column['host'] = '请求地址'
        self.column['Method'] = '请求方式'
        self.column['status'] = '接口状态'

        return self.column