import tkinter as tk
from tkinter import *
from tkinter import ttk
from Control.read_file import read_file
from View.Inter_Conf_Frame import *


class Edit_Conf(tk.Toplevel): # 继承Toplevel类
    def __init__(self,view_item,parent):
        super().__init__()
        self.title("修改接口配置")
        self.geometry('600x400+400+20') #设置窗口大小
        self.parent = parent # 显式地保留父窗口
        self.read_file = read_file()
        self.case_type = StringVar()
        self.model = StringVar()
        self.IP = StringVar()         # 请求地址
        self.interName = StringVar()
        self.request_type = StringVar()    # 请求类型
        self.inter_Method = StringVar()   # 请求方式
        self.inter_Status = StringVar()   # 用例状态
        self.inter_Ename = StringVar()  # 测试用例英文名称
        self.case_header_list = []
        self.grab_set()
        self.createPage(view_item)


    def close_popup(self):
        self.destroy()


    def createPage(self,view_item):

        self.labelframe = tk.Frame(self)
        #self.labelframe = LabelFrame(self, text="修改用例",font=('Tempus Sams ITC',20))
        self.labelframe.pack(fill="both", expand="yes")

        # 从公共模块获取测试用例列表字段
        self.column = Ordered_Dict()
        column = self.column.inter_conf_header()
        column_list = list(column.values())


        #布局label
        for value in column_list:

            i = column_list.index(value)
            label_column = Label(self.labelframe, text=value + "：" )
            label_column.grid(row=i, stick=E,column=0, pady=7,padx=40)

        '''    #布局字段内容,用Text控件布局
        for value1 in column_list[:6]:
            i = column_list.index(value1)
            text_param = tk.Text(self.labelframe,width=50,height=1)
            text_param.insert(INSERT,view_item[i])


            text_param.grid(row=i, stick=W,column=2, pady=7,padx=0)
            text_param.bind("<KeyPress>", lambda e: "break")  # 文本框只读'''



        # 单独针对请求包体和期望包体字段布局
            #用例类型
        self.case_type_list = ttk.Label(self.labelframe, textvariable=self.case_type)
        self.case_type.set(view_item[0])
        self.case_type_list.grid(row=0, stick=W, column=1,pady=7,columnspan=2)
            # 所属模块
        self.model_Chosen = ttk.Label(self.labelframe, textvariable=self.model)
        self.model.set(view_item[1])
        self.model_Chosen.grid(row=1, stick=W, column=1,pady=7,columnspan=2)
            # 接口名称
        self.interName_Chosen = ttk.Label(self.labelframe, textvariable=self.interName)
        self.interName.set(view_item[2])
        self.interName_Chosen.grid(row=2, stick=W, column=1,pady=7,columnspan=2)
            # 接口英文名称
        ttk.Label(self.labelframe,textvariable=self.inter_Ename).grid(row=3, stick=W, column=1,pady=7,columnspan=2)
        self.inter_Ename.set(view_item[3])

        # 请求类型
        request_type_select = ttk.Combobox(self.labelframe,width=30,textvariable=self.request_type)
        request_type_select['values'] = ('http','https') # 设置下拉列表的值
        request_type_select.grid(row=4,column=1, stick=W, pady=7,columnspan=2)
        request_type_select.current(request_type_select['values'].index(view_item[4])) # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值

            # 请求IP地址
        self.label_IP = ttk.Entry(self.labelframe, textvariable=self.IP,width=32)
        self.IP.set(view_item[5])
        self.label_IP.grid(row=5, stick=W, column=1,pady=7,columnspan=2)

        # self.label_Method = ttk.Entry(self.labelframe, textvariable=self.inter_Method)
        # self.inter_Method.set(view_item[6])
        # self.label_Method.grid(row=6, stick=W, column=1, pady=7, columnspan=2)
        # 请求方式
        inter_Method_select = ttk.Combobox(self.labelframe,width=30,textvariable=self.inter_Method)
        inter_Method_select['values'] = ('POST', 'GET', 'PUT', 'DELETE')  # 设置下拉列表的值
        inter_Method_select.grid(row=6,column=1, stick=W, pady=7,columnspan=2)
        inter_Method_select.current(inter_Method_select['values'].index(view_item[6])) # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值

            #用例状态
        ttk.Label(self.labelframe, text = view_item[7]).grid(row=7, column=1, stick=W, pady=10)
            #保存/关闭
        ttk.Button(self.labelframe, text='保存', width=8, command=self.summit_data).grid(row=999,column=0,columnspan=1,sticky=E)
        ttk.Button(self.labelframe, text='关闭', width=8, command=self.close_popup).grid(row=999,column=1,columnspan=1,sticky=W,padx=80)

    def Editrow(self):

        search_interName = self.interName.get()
        search_inter_Ename = self.inter_Ename.get()

            # 读取EXCEL用例数据到列表中
        case_interName = self.read_file.read_conf_file().col_values(2)[1:]
        case_inter_Ename = self.read_file.read_conf_file().col_values(3)[1:]
        for k in range(0,len(case_interName)):
            if case_interName[k] == search_interName and case_inter_Ename[k] == search_inter_Ename:
                return k+1

    def summit_data(self):
        column = self.column.inter_conf_header()
        w_xls,sheet_write,rows = self.read_file.write_conf_file()

        if self.IP.get() in ['','\n']:  #用例名称为空
            showerror(title='错误', message=column['host'] + '不可为空')
            return



                # 逐列写入数据
        self.case_header_list = [self.request_type.get(), self.IP.get(), self.inter_Method.get()]

        rows = self.Editrow()
        try:
            for i in range(0,len(self.case_header_list)):
                sheet_write.write(rows, i+4 ,self.case_header_list[i])
            w_xls.save(self.read_file.save_conf_file)
            showinfo(title="信息",message="用例接口配置成功！")
            self.destroy()
            self.parent.search()
        except:
            showerror(title='错误', message="接口配置保存失败！")