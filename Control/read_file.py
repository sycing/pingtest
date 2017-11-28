from Control.read_path import read_path
import xlrd
from xlutils.copy import copy

class read_file():

    def __init__(self):
        self.case_path_file = "\pingtest\Data\TestCase.xls"
        self.conf_path_file = "\pingtest\Data\project_conf.xls"
        self.task_path_file = "\pingtest\Data\Task.xls"
        self.accounts_file = "\pingtest\Data\Accounts.xls"
        self.save_task_file = ".\Data\Task.xls"
        self.save_case_file = ".\Data\TestCase.xls"
        self.save_conf_file = ".\Data\project_conf.xls"
        self.save_accounts_file = ".\Data\Accounts.xls"

    def read_case_file(self):

        case = read_path(self.case_path_file)
        case_execl_object = xlrd.open_workbook(case.read_superior_path())
        case_sheet = case_execl_object.sheet_by_index(0)

        return case_sheet

    def write_case_file(self):

        case = read_path(self.case_path_file)
        case_execl_object = xlrd.open_workbook(case.read_superior_path())
        case_sheet = case_execl_object.sheet_by_index(0)
        rows = case_sheet.nrows
        w_xls = copy(case_execl_object)
        sheet_write = w_xls.get_sheet(0)

        return w_xls,sheet_write,rows

    def write_task_file(self):

        task = read_path(self.task_path_file)
        task_execl_object = xlrd.open_workbook(task.read_superior_path())
        task_sheet = task_execl_object.sheet_by_index(0)
        rows = task_sheet.nrows
        w_xls = copy(task_execl_object)
        sheet_write = w_xls.get_sheet(0)

        return w_xls,sheet_write,rows

    def read_task(self):

        task = read_path(self.task_path_file)
        task_execl_object = xlrd.open_workbook(task.read_superior_path())
        task_sheet = task_execl_object.sheet_by_index(0)
        return task_sheet

    def read_conf_file(self):

        conf = read_path(self.conf_path_file)
        conf_execl_object = xlrd.open_workbook(conf.read_superior_path())
        conf_sheet = conf_execl_object.sheet_by_index(0)
        return conf_sheet

    def write_conf_file(self):
        conf = read_path(self.conf_path_file)
        conf_execl_object = xlrd.open_workbook(conf.read_superior_path())
        conf_sheet = conf_execl_object.sheet_by_index(0)
        rows = conf_sheet.nrows
        w_xls = copy(conf_execl_object)
        sheet_write = w_xls.get_sheet(0)
        return w_xls,sheet_write,rows

    def read_accounts_file(self):
        accounts = read_path(self.accounts_file)
        accounts_execl_object = xlrd.open_workbook(accounts.read_superior_path())
        accounts_sheet = accounts_execl_object.sheet_by_index(0)
        return accounts_sheet

    def write_accounts_file(self):
        accounts = read_path(self.accounts_file)
        accounts_execl_object = xlrd.open_workbook(accounts.read_superior_path())
        accounts_sheet = accounts_execl_object.sheet_by_index(0)
        rows = accounts_sheet.nrows
        w_xls = copy(accounts_execl_object)
        sheet_write = w_xls.get_sheet(0)
        return w_xls,sheet_write,rows
