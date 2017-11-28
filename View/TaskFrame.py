from tkinter import *
import tkinter as tk
from tkinter.messagebox import *
from View.LoginPage import *
from tkinter import ttk
from Control.read_file import read_file
from View.New_task import New_task
from Control.Ordered_Dict import Ordered_Dict
from Control.Execute_Interface_Main import Execute_Interface_Main
from View.View_Task import View_Task
from View.View_report import FindFile
from View.Edit_task import Edit_task
from tkinter import scrolledtext        # 导入滚动文本框的模块
import os
import win32com.client

class TaskFrame(Frame): # 继承Frame类
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.root = master #定义内部变量root
        self.task_ID = StringVar()
        self.taskName = StringVar()
        self.task_statues = StringVar()
        self.task_result = StringVar()
        self.task_time = StringVar()
        self.read_file = read_file()
        self.task_list = []
        self.createPage()

    def logout(self):
        TaskFrame.quit(self.root)

    def createPage(self):
        def search(event):
            self.search()
        self.labelframe = LabelFrame(self.root, text="任务列表",font=('Tempus Sams ITC',20))
        self.labelframe.pack(fill="both", expand="yes")

        ttk.Label(self.labelframe, text = '任务ID: ').grid(row=1, stick=E,column=0, pady=10,padx=10)
        task_ID = ttk.Entry(self.labelframe, textvariable=self.task_ID)
        task_ID.grid(row=1, column=1, stick=W)
        task_ID.bind("<Return>",search)

        ttk.Label(self.labelframe, text = '任务名称: ').grid(row=1, stick=W,column=2, pady=10,padx=10)
        taskName = ttk.Entry(self.labelframe, textvariable=self.taskName)
        taskName.grid(row=1, column=3, stick=W)
        taskName.bind("<Return>",search)

        ttk.Label(self.labelframe, text = '执行状态: ').grid(row=1, stick=W, column=4,pady=10,padx=10)
        task_statues_Chosen = ttk.Combobox(self.labelframe,width=12,textvariable=self.task_statues)
        task_statues_Chosen['values'] = ['','待执行','已执行'] # 设置下拉列表的值
        task_statues_Chosen.grid(row=1,column=5, stick=W)
        task_statues_Chosen.bind("<Return>",search)

        ttk.Label(self.labelframe, text = '执行结果: ').grid(row=1, stick=E, column=6,pady=10,padx=10)
        task_result_Chosen = ttk.Combobox(self.labelframe,width=20,textvariable=self.task_result)
        task_result_Chosen['values'] = ['','Pass','Fail'] # 设置下拉列表的值
        task_result_Chosen.grid(row=1,column=7, stick=W)
        task_result_Chosen.bind("<Return>",search)

        ttk.Label(self.labelframe, text = '执行时间: ').grid(row=1, stick=E, column=8,pady=10,padx=10)
        task_time = ttk.Entry(self.labelframe, textvariable=self.task_time)
        task_time.grid(row=1, column=9, stick=W)
        task_time.bind("<Return>",search)

        search_button = ttk.Button(self.labelframe, text='搜索',command=self.search)
        search_button.grid(row=1,column=10,padx=30,stick=W)
        ttk.Button(self.labelframe, text='登出',command=self.logout).grid(row=1,column=11,padx=10,stick=W)

        ttk.Button(self.labelframe, text='新增测试任务',command=self.new_task).grid(row=2,column=0,padx=10,stick=W)
        ttk.Button(self.labelframe, text='编辑测试任务',command=self.edit_task).grid(row=2,column=1,padx=10,stick=W)
        execute = ttk.Button(self.labelframe, text='执行/再次执行',command=self.execute_interface)
        execute.grid(row=2,column=2,padx=10,stick=W)

        ttk.Button(self.labelframe, text='查看报告',command=self.view_report).grid(row=2,column=3,padx=10,stick=W)
        ttk.Button(self.labelframe, text='查看log').grid(row=2,column=4,padx=10,stick=W)
        ttk.Button(self.labelframe, text='删除任务',command=self.del_task).grid(row=2,column=5,padx=10,stick=W)

        self.column = Ordered_Dict()
        column = self.column.task_header()

        style = ttk.Style(self.labelframe)
        style.configure('Treeview',rowheight=20,relief='raised')

        self.tree = ttk.Treeview(self.labelframe,height=20,show="headings",columns=tuple(column.keys()))

        for i in column.keys():
            self.tree.column(i, width=150, anchor='w')

        for k,v in column.items():
            self.tree.heading(k,text=v)

        vbar = ttk.Scrollbar(self.labelframe,orient=VERTICAL,command=self.tree.yview)
        self.tree.configure(yscrollcommand=vbar.set)

        task_sheet = self.read_file.read_task()
        self.task_list = self.read_file.read_task().col_values(0)[1:]

        for index in range(1,(len(self.task_list)+1)):
            self.tree.insert('',-index,values=task_sheet.row_values(index))

        self.tree.grid(row=3,column=0,columnspan=12,pady=10,sticky=NSEW)
        vbar.grid(row=3,column=12,pady=20,sticky=NS)

        ttk.Label(self.labelframe, text='执行任务LOG: ').grid(row=4, stick=W, column=0)
        self.show_log = Text(self.labelframe,width=250,height=30,wrap=tk.WORD)
        self.show_log.tag_config('b',foreground = 'green')
        self.show_log.tag_config('a',foreground = 'red')
        self.show_log.tag_config('c',foreground = 'blue')
        self.show_log.grid(row=5, stick=W, column=0,pady=0,columnspan=12)
        vbar2 = ttk.Scrollbar(self.labelframe,orient=VERTICAL,command=self.show_log.yview)
        self.show_log.configure(yscrollcommand=vbar2.set)
        vbar2.grid(row=5,column=12,pady=0,sticky=NS)

        def treeviewClick(event):
            item = self.tree.selection()[0]
            selected_item = self.tree.item(item, "values")
            popup_view_case = View_Task(selected_item)
            self.wait_window(popup_view_case)
        self.tree.bind('<Double-Button-1>', treeviewClick)

    def search(self):  # 组合查询功能

        ori_items = self.tree.get_children()

        [self.tree.delete(item) for item in ori_items]
        # 把查询条件的值放入列表中，用于遍历
        search_item = [self.task_ID.get(),self.taskName.get(),'','','','','','',self.task_statues.get(),self.task_result.get(),self.task_time.get()]
        case_sheet = self.read_file.read_task()

        search_case_list = range(1,case_sheet.nrows)
        temp_list = []

        # 读取EXCEL用例数据到列表中
        for i in range(0,len(search_item)):
            case_list = self.read_file.read_task().col_values(i)[1:]
            for k in range(0,len(case_list)):
                if -1 != case_list[int(k)].find(search_item[i].strip()):
                    temp_list.append(k+1)

            search_case_list = list(set(search_case_list).intersection(set(temp_list)))
            temp_list = []

        search_case_list.sort()
        for index in search_case_list:
            self.tree.insert('',-index,values=case_sheet.row_values(index))

    def new_task(self):
        # 打开新增任务pop页面
        popup_new_task = New_task(self)
        self.wait_window(popup_new_task)

    def edit_task(self):
        if self.tree.selection() not in ['',()]:
            # 打开编辑任务pop页面
            item = self.tree.selection()[0]
            selected_task = self.tree.item(item, "values")
            popup_edit_task = Edit_task(self,selected_task)
            self.wait_window(popup_edit_task)
        else:
            showerror(title="错误",message="请选择任务")

    def execute_interface(self):
        self.show_log.focus()
        self.show_log.delete('1.0',END)

        if self.tree.selection() not in ['',()]:
            # 执行任务入口
            item = self.tree.selection()
            for i in range(0,len(item)):
                selected_task = self.tree.item(item[i], "values")
                #try:
                excution = Execute_Interface_Main(selected_task,self)
                w_xls,sheet_write,rows = self.read_file.write_task_file()
                self.task_list = self.read_file.read_task().col_values(0)[1:]
                row = self.task_list.index(selected_task[0])

                sheet_write.write(row+1,8,"已执行")
                sheet_write.write(row+1,9,excution.test_result)
                sheet_write.write(row+1,10,excution.test_time)
                w_xls.save(self.read_file.save_task_file)
            showinfo(title="信息",message="任务执行完毕")

            self.search()
            '''except BaseException as e:
                showinfo(title="信息",message="执行失败！\n%s"%e)'''
        else:
            showinfo(title="信息",message="请选择任务！")

    def view_report(self):
        item = self.tree.selection()
        if item =='':
            showerror(title='错误', message='请选中任务！')
        else:
            self.selected_item=self.tree.item(item[0], 'values')
            selected_item=self.selected_item
            pop=FindFile(selected_item[0])
            pop.MaxTime()
            #self.wait_window(pop)

    def del_task(self):        #任务列表物理删除
        path = os.getcwd()
        log_file = path + "\Data\Task.xls"
        self.xlApp=win32com.client.Dispatch('Excel.Application')
        self.xlApp.Visible=0
        self.xlApp.DisplayAlerts=0    #后台运行，不显示，不警告
        self.xlBook=self.xlApp.Workbooks.Open(log_file)
        self.sht = self.xlBook.Worksheets('Sheet1')

        item1 = list(self.tree.selection())
        self.task_list = self.read_file.read_task().col_values(0)

        if item1 is []:
            showerror(title='错误', message='请选择删除的任务！')
        else:
            confirm = askyesno(title="确认", message="确认删除所选任务？")
            if confirm is True:
                for k in item1:
                   # K = k.split()
                    selected_task = self.tree.item(k, "values")
                    for value in self.task_list:
                        if selected_task[0] == value:
                            row = self.task_list.index(selected_task[0])
                            self.sht.Rows(row+1).Delete()
                            self.xlBook.Save()
                            break
                            #[self.tree.delete(item) for item in K]

            showinfo(title="信息", message="删除成功！")
            self.search()
