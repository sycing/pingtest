from View.Inter_Conf_Frame import *
from tkinter import ttk
from Control.read_file import read_file
from Control.Ordered_Dict import Ordered_Dict
from View.View_Case import View_Case
from View.New_Case import New_Case
from View.EditCase_View import EditCase_View
from test_main import test_main
import time
import os
import win32com.client

class CaseFrame(Frame): # 继承Frame类
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.root = master #定义内部变量root
        self.interName = StringVar()
        self.case_type = StringVar()
        self.model = StringVar()
        self.caseID = StringVar()
        self.inter_Method = StringVar()
        self.Case_Status = StringVar()
        self.read_file = read_file()
        self.selected_item = ()
        self.case_count = 0
        self.createPage()

    def logout(self):
        CaseFrame.quit(self.root)

    def createPage(self):

        self.labelframe = LabelFrame(self.root, text="用例管理",font=('Tempus Sams ITC',20))
        self.labelframe.pack(fill="both", expand="yes")
        def search(event):
            self.search()
        ttk.Label(self.labelframe, text = '用例类型: ').grid(row=0, stick=W,column=0, pady=10,padx=10)
        case_type_Chosen = ttk.Combobox(self.labelframe,width=20,textvariable=self.case_type)
        case_type_Chosen['values'] = list(set(self.read_file.read_conf_file().col_values(0)[1:]))  # 设置下拉列表的值
        case_type_Chosen.grid(row=0,column=0, stick=E,padx=10)
        case_type_Chosen.bind("<Return>",search)

        ttk.Label(self.labelframe, text = '所属模块: ').grid(row=0, stick=W, column=1,pady=10,padx=30)
        model_Chosen = ttk.Combobox(self.labelframe,width=20,textvariable=self.model)
        model_Chosen['values'] = list(set(self.read_file.read_conf_file().col_values(1)[1:]))  # 设置下拉列表的值
        model_Chosen.grid(row=0,column=1, stick=E,padx=30)
        model_Chosen.bind("<Return>",search)

        ttk.Label(self.labelframe, text = '接口名称: ').grid(row=0, stick=W, column=2,pady=10,padx=10)
        interName_Chosen = ttk.Combobox(self.labelframe,width=20,textvariable=self.interName)
        interName_Chosen['values'] = list(set(self.read_file.read_conf_file().col_values(2)[1:]))  # 设置下拉列表的值
        interName_Chosen.grid(row=0,column=2, stick=E,padx=30)
        interName_Chosen.bind("<Return>",search)

        ttk.Label(self.labelframe, text = '用例ID:').grid(row=1, stick=W, column=0,pady=10,padx=10)
        caseID = ttk.Entry(self.labelframe, textvariable=self.caseID,width=20)
        caseID.grid(row=1, column=0, stick=E,padx=30)
        caseID.bind("<Return>",search)

        # 创建一个下拉框
        ttk.Label(self.labelframe, text = '请求方式: ').grid(row=1, stick=W, column=1,pady=10,padx=30)

        Method_Chosen = ttk.Combobox(self.labelframe,width=20,textvariable=self.inter_Method)
        Method_Chosen['values'] = ('','POST','GET','PUT','DELETE') # 设置下拉列表的值
        Method_Chosen.grid(row=1,column=1, stick=E,padx=30)
        Method_Chosen.current(0) # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值
        Method_Chosen.bind("<Return>",search)

        # 用例状态下拉框
        ttk.Label(self.labelframe, text = '用例状态: ').grid(row=1, stick=W, column=2,pady=10,padx=10)
        Case_Status = ttk.Combobox(self.labelframe,width=20,textvariable=self.Case_Status)
        Case_Status['values'] = ('','有效','无效') # 设置下拉列表的值
        Case_Status.grid(row=1,column=2, stick=E,padx=30)
        Case_Status.current(1) # 设置下拉列表默认显示的值
        Case_Status.bind("<Return>",search)

        ttk.Button(self.labelframe, text='搜索',command=self.search).grid(row=1,column=3,padx=30,stick=W)
        #Button(self.labelframe, text='导入用例', width=8, height=1).grid(row=1,column=12,padx=10,stick=W)
        #Button(self.labelframe, text='下载模板', width=8, height=1).grid(row=1,column=13,padx=10,stick=W)
        #Button(self.labelframe, text='导出用例', width=8, height=1).grid(row=1,column=14,padx=10,stick=W)

        ttk.Button(self.labelframe, text='登出',command=self.logout).grid(row=1,column=4,padx=10,stick=W)

        ttk.Button(self.labelframe, text='新增用例',command=self.new_case).grid(row=2,column=0,padx=10,stick=W)
        ttk.Button(self.labelframe, text='编辑用例',command=self.ButtonClick).grid(row=2,column=0,padx=10,stick=E)
        ttk.Button(self.labelframe, text='删除用例',command=self.del_case).grid(row=2,column=1,padx=60,stick=W)

        self.case_count_label = ttk.Label(self.labelframe)
        self.case_count_label.grid(row=3,column=0,padx=10,stick=W,pady=3)
        self.case_count_label['text'] = "共 "+str(self.case_count) + " 条用例"

        # 从公共模块获取测试用例列表字段
        self.column = Ordered_Dict()
        column = self.column.test_case_header()

        style = ttk.Style(self.labelframe)
        style.configure('Treeview',rowheight=40,relief='raised')

        self.tree = ttk.Treeview(self.labelframe,height=35,show="headings",columns=tuple(column.keys()))

        for i in column.keys():
            self.tree.column(i, width=110, anchor='w')

        for k,v in column.items():
            self.tree.heading(k,text=v, command=lambda c=k : self.sortby(c, 0))

        vbar = ttk.Scrollbar(self.labelframe,orient=VERTICAL,command=self.tree.yview)
        self.tree.configure(yscrollcommand=vbar.set)

        # 读取EXCEL用例数据到列表中
        self.search()
        self.tree.grid(row=4,column=0,columnspan=8,pady=5,sticky=NSEW)
        vbar.grid(row=4,column=13,pady=20,sticky=NS)

        def treeviewClick(event):
            item = self.tree.selection()[0]
            selected_item = self.tree.item(item, "values")
            popup_view_case = View_Case(selected_item)
            self.wait_window(popup_view_case)

        def selectAll(event):
            ori_items = self.tree.get_children()               # 获取Treeview的所有元素
            self.tree.selection_set(ori_items)                 # 选中多行数据， ori_items，即以I开头的行的id编码的元组。

        self.tree.bind('<Double-Button-1>', treeviewClick)
        self.tree.bind('<Control-a>', selectAll)               # 实现用Ctrl+a可选中Treeview中的所有元素

    def ButtonClick(self):

            item = self.tree.selection()
            if item =='':
                showerror(title='错误', message='请选中编辑内容！')
            else:
                self.selected_item=self.tree.item(item[0], 'values')
                #print(self.selected_item)
                selected_item=self.selected_item
                pop=EditCase_View(selected_item,self)
                self.wait_window(pop)

    def sortby(self, col, descending):
        """sort tree contents when a column header is clicked on"""
        # grab values to sort
        data = [(self.tree.set(child, col), child) for child in self.tree.get_children('')]
        # if the data to be sorted is numeric change to float
        # data =  change_numeric(data)
        # now sort the data in place
        data.sort(reverse=descending)
        for index, item in enumerate(data):
            self.tree.move(item[1], '', index)
        # switch the heading so it will sort in the opposite direction
        self.tree.heading(col, command=lambda col=col: self.sortby(col, int(not descending)))

    def search(self):

        ori_items = self.tree.get_children()

        [self.tree.delete(item) for item in ori_items]
        # 把查询条件的值放入列表中，用于遍历
        search_item = [self.case_type.get(),self.model.get(),self.interName.get(),self.caseID.get(),
                       self.inter_Method.get(),'','','','','','','','','','',self.Case_Status.get()]
        case_sheet = self.read_file.read_case_file()

        search_case_list = range(1,case_sheet.nrows)
        temp_list = []

        # 读取EXCEL用例数据到列表中
        for i,v in enumerate(search_item):
            case_list = case_sheet.col_values(i)[1:]

            for k,k1 in enumerate(case_list):
                if -1 != case_list[int(k)].find(search_item[i].strip()):
                    temp_list.append(k+1)

            search_case_list = list(set(search_case_list).intersection(set(temp_list)))
            temp_list = []

        search_case_list.sort()

        for index in search_case_list:
            insert_list = case_sheet.row_values(index)
            self.tree.insert('','end',values=insert_list)

        self.case_count_label['text'] = "共 " + str(len(self.tree.get_children())) + " 条用例"

    def new_case(self):

        popup_new_case = New_Case(self)
        self.wait_window(popup_new_case)

    def del_case(self):
        if self.tree.selection() in ['',()]:
            showerror(title='错误', message='请选中内容！')
        else:
            items = self.tree.selection()
            confirm = askyesno(title="确认", message="删除用例后，关联的测试任务会无法执行，确认删除所选用例？")
            #if selected_item[13] == '有效':

            if confirm is True:
                    w_xls,sheet_write,rows = self.read_file.write_case_file()
                    case_list = self.read_file.read_case_file().col_values(3)
                    for item in items:
                        selected_item = self.tree.item(item, "values")
                        if selected_item[15] != '无效':
                            for i in range(0,len(case_list)):
                                if selected_item[3] == case_list[i]:
                                    sheet_write.write(i,15,'无效')
                            w_xls.save(self.read_file.save_case_file)
                        else:
                                path = os.getcwd()
                                log_file = path + "\Data\TestCase.xls"
                                self.xlApp=win32com.client.Dispatch('Excel.Application')
                                self.xlApp.Visible=0
                                self.xlApp.DisplayAlerts=0    #后台运行，不显示，不警告
                                self.xlBook=self.xlApp.Workbooks.Open(log_file)
                                self.sht = self.xlBook.Worksheets('Sheet1')
                                del_case_list = self.read_file.read_case_file().col_values(3)
                                if items == []:
                                    showerror(title='错误', message='请选择删除的用例！')
                                else:
                                        for value in del_case_list:
                                            if selected_item[3] == value:
                                                row = del_case_list.index(selected_item[3])
                                                self.sht.Rows(row+1).Delete()
                                                self.xlBook.Save()
                                                break
                    os.system('taskkill /im EXCEL.EXE /f')
                    showinfo(title="信息", message="删除成功！")
                    self.search()
                    return