from tkinter import ttk
import tkinter as tk
from View.TaskFrame import *
import json
from tkinter import scrolledtext        # 导入滚动文本框的模块

class View_Task(tk.Toplevel): # 继承Toplevel类
    def __init__(self,view_item):
        super().__init__()
        self.title("查看任务")
        self.geometry('800x700+400+100') #设置窗口大小
        self.grab_set()
        self.createPage(view_item)

    def close_popup(self):
        self.destroy()

    def createPage(self,view_item):

        self.labelframe = ttk.Frame(self)
        self.labelframe.pack(fill="both", expand="yes")

        # 从公共模块获取测试用例列表字段
        self.column = Ordered_Dict()
        column = self.column.task_header()
        column_list = list(column.values())

        for value in column_list:

            i = column_list.index(value)
            label_column = ttk.Label(self.labelframe, text = value + "：")
            label_column.grid(row=i, stick=E,column=0,pady=6,padx=40)

            #用Text控件布局
            text_param = tk.Text(self.labelframe,width=80,height=1,borderwidth=0)
            text_param.insert(INSERT,view_item[i])
            text_param['state'] = DISABLED
            text_param.grid(row=i, stick=W,column=2, columnspan=1, rowspan=1,pady=7,padx=0)

            #ttk.Label(self.labelframe, text = view_item[i]).grid(row=i, column=1,columnspan=2, rowspan=1,sticky=W+N, pady=10)
            #单独对执行用例ID进行布局
            text_param = scrolledtext.ScrolledText(self.labelframe,width=80,height=20,borderwidth=0)
            text_param.insert(INSERT,view_item[4])
            text_param['state'] = DISABLED
            text_param.grid(row=4, stick=W,column=2, columnspan=3, rowspan=1,pady=7,padx=0)

            setip_SQL = tk.Text(self.labelframe,width=80,height=3,borderwidth=0)
            setip_SQL.insert(INSERT,view_item[5])
            setip_SQL['state'] = DISABLED
            setip_SQL.grid(row=5, stick=W,column=2, columnspan=3, rowspan=1,pady=7,padx=0)

            teardown_SQL = tk.Text(self.labelframe,width=80,height=3,borderwidth=0)
            teardown_SQL.insert(INSERT,view_item[6])
            teardown_SQL['state'] = DISABLED
            teardown_SQL.grid(row=6, stick=W,column=2, columnspan=3, rowspan=1,pady=7,padx=0)

        ttk.Button(self.labelframe, text='关闭',command=self.close_popup).grid(row=999,column=1,columnspan=2,sticky=W, pady=20)
