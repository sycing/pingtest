from tkinter import ttk
import tkinter as tk
from View.Inter_Conf_Frame import *

class New_Conf(tk.Toplevel): # 继承Toplevel类
    def __init__(self,parent):
        super().__init__()
        self.title("新增接口")
        self.geometry('500x400+400+20') #设置窗口大小
        self.parent = parent # 显式地保留父窗口
        self.read_file = read_file()
        self.case_type = StringVar()
        self.model = StringVar()
        self.interName = StringVar()
        self.requestType = StringVar()
        self.IPPort = StringVar()
        self.inter_Method = StringVar()
        self.inter_Status = StringVar()
        self.inter_Ename = StringVar()  # 测试用例英文名称
        self.case_header_list = []
        self.grab_set()
        self.createPage()

    def close_popup(self):
        # put focus back to the parent window
        self.parent.focus_set()
        self.destroy()

    def createPage(self):

        # self.labelframe = LabelFrame(self, text="新增接口配置",font=('Tempus Sams ITC',20))
        self.labelframe = Frame(self)
        self.labelframe.pack(fill="both", expand="yes")

        # 从公共模块获取测试用例列表字段
        self.column = Ordered_Dict()
        column = self.column.inter_conf_header()
        column_list = list(column.values())

        # 布局字段label
        for value in column_list:

            i = column_list.index(value)
            label_column = ttk.Label(self.labelframe, text = value + "：")
            label_column.grid(row=i, stick=E,column=0, pady=3,padx=30)

        # 用用例类型
        self.case_type_entry = ttk.Entry(self.labelframe, textvariable=self.case_type, width=32)
        self.case_type_entry.grid(row=0, stick=W, column=1, pady=10, columnspan=2)

        # 所属模块
        self.model_entry = ttk.Entry(self.labelframe, textvariable=self.model, width=32)
        self.model_entry.grid(row=1, stick=W, column=1, pady=10, columnspan=2)

        # 接口名称
        self.interName_entry = ttk.Entry(self.labelframe, textvariable=self.interName, width=32)
        self.interName_entry.grid(row=2, stick=W, column=1, pady=10, columnspan=2)

        # 接口英文名称
        self.inter_Ename_entry = ttk.Entry(self.labelframe, textvariable=self.inter_Ename, width=32)
        self.inter_Ename_entry.grid(row=3, stick=W, column=1, pady=10, columnspan=2)

        # 请求类型
        self.requestType_list = ttk.Combobox(self.labelframe, width=30, text=self.requestType)
        self.requestType_list['values'] = ['http','https']  # 设置下拉列表的值
        self.requestType_list.grid(row=4, column=1, stick=W, columnspan=2)
        self.requestType_list.current(0)

        # 请求IP端口
        self.label_IPPort = ttk.Entry(self.labelframe, textvariable=self.IPPort,width=32)
        self.label_IPPort.grid(row=5, stick=W, column=1,pady=10,columnspan=2)


        # 请求方式
        self.case_status_list = ttk.Combobox(self.labelframe,width=30,text=self.inter_Method)
        self.case_status_list['values'] = ['POST', 'GET', 'PUT', 'DELETE']  # 设置下拉列表的值
        self.case_status_list.grid(row=6,column=1, stick=W,columnspan=2)
        self.case_status_list.current(0)

        # 接口状态下拉框
        self.case_status_list = ttk.Combobox(self.labelframe,width=30,text=self.inter_Status)
        self.case_status_list['values'] = ['有效','无效']  # 设置下拉列表的值
        self.case_status_list.grid(row=7,column=1, stick=W,columnspan=2)
        self.case_status_list.current(0)

        self.summit_button = ttk.Button(self.labelframe, text='保存',command=self.summit_data)
        self.summit_button.grid(row=999,column=1,sticky=W, pady=10)
        ttk.Button(self.labelframe, text='关闭',command=self.close_popup).grid(row=999,column=2,sticky=NE, pady=10,padx=0)

    def summit_data(self):
        column = self.column.inter_conf_header()
        w_xls,sheet_write,rows = self.read_file.write_conf_file()

        # 读取EXCEL用例数据到列表中
        case_interName = self.read_file.read_conf_file().col_values(2)[1:]
        case_inter_Ename = self.read_file.read_conf_file().col_values(3)[1:]

        if self.case_type.get() == '': # 用例类型为空
            showerror(title='错误', message=column['case_type'] + '不可为空')
            return
        elif self.model.get() == '': # 所属模块为空
            showerror(title='错误', message=column['model'] + '不可为空')
            return
        elif self.interName.get() == '':  # 接口名称为空
            showerror(title='错误', message=column['interfacename'] + '不可为空')
            return
        elif self.interName.get() in case_interName:  # 接口名称已存在
            showerror(title='错误', message=column['interfacename'] + '不能重复，请输入新的接口名称！')
            return
        elif self.inter_Ename.get() == '':  # 接口英文名称为空
            showerror(title='错误', message=column['interfaceEname'] + '不可为空')
            return
        elif self.inter_Ename.get() in case_inter_Ename:  # 接口英文名称已存在
            showerror(title='错误', message=column['interfaceEname'] + '不能重复，请输入新的接口英文名称！')
            return
        elif self.IPPort.get() == '':  # 请求地址为空
            showerror(title='错误', message=column['host'] + '不可为空')
            return
        elif self.requestType.get() == '':  # 请求类型为空
            showerror(title='错误', message=column['requestType'] + '不可为空')
            return
        elif self.inter_Method.get() == '':  # 请求方式为空
            showerror(title='错误', message=column['Method'] + '不可为空')
            return
        elif self.inter_Status.get() == '':  # 状态为空
            showerror(title='错误', message=column['status'] + '不可为空')
            return

        # 逐列写入数据
        self.case_header_list = [self.case_type.get(),self.model.get(),self.interName.get(),self.inter_Ename.get(),self.requestType.get(),self.IPPort.get(),self.inter_Method.get(),self.inter_Status.get()]
        try:
            for i in range(0,len(self.case_header_list)):
                sheet_write.write(rows,i,self.case_header_list[i])
            w_xls.save(self.read_file.save_conf_file)
            showinfo(title="信息",message="接口新增成功！")
            self.destroy()
            self.parent.search()
        except:
            showerror(title='错误', message="用例保存失败！")

