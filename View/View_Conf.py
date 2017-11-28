from tkinter import ttk
import tkinter as tk
from View.Inter_Conf_Frame import *
from Control.Ordered_Dict import Ordered_Dict
import json

class View_Conf(tk.Toplevel): # 继承Toplevel类
    def __init__(self,view_item):
        super().__init__()
        self.title("查看接口配置")
        self.geometry('600x400+400+20') #设置窗口大小
        self.grab_set()    # 确保鼠标或者键盘事件不会被发送到错误的窗口。 使得只能打开一个弹框，且只能在此弹框关闭后，才能对root进行操作，相当于锁定
        self.createPage(view_item)

    def close_popup(self):
        self.destroy()

    def createPage(self,view_item):

        self.labelframe = ttk.Frame(self)
        self.labelframe.pack(fill="both", expand="yes")

        # 从公共模块获取测试用例列表字段
        self.column = Ordered_Dict()
        column = self.column.inter_conf_header()
        column_list = list(column.values())

        for value in column_list:

            i = column_list.index(value)
            label_column = ttk.Label(self.labelframe, text = value + "：")
            label_column.grid(row=i, stick=E,column=0, pady=6,padx=40)

            #用Text控件布局
            text_param = tk.Text(self.labelframe,width=50,height=1,borderwidth=0)
            text_param.insert(INSERT,view_item[i])
            text_param['state'] = DISABLED
            text_param.grid(row=i, stick=W,column=2, pady=7,padx=0)

        # 单独针对请求包体和期望包体字段布局
        # text_param = tk.Text(self.labelframe,width=50,height=8,borderwidth=0)
        # text_param.insert(INSERT,view_item[7])
        # text_param['state'] = DISABLED
        # text_param.grid(row=7, stick=W,column=2, pady=7,padx=0)
        #
        # text_param = tk.Text(self.labelframe,width=50,height=8,borderwidth=0)
        # text_param.insert(INSERT,view_item[9])
        # text_param['state'] = DISABLED
        # text_param.grid(row=9, stick=W,column=2, pady=7,padx=0)

        ttk.Button(self.labelframe, text='关闭',command=self.close_popup).grid(row=999,column=1,columnspan=2,sticky=W, pady=20)

