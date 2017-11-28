from tkinter import ttk
import tkinter as tk
import json
from View.CaseFrame import *
from tkinter import scrolledtext        # 导入滚动文本框的模块

class New_Case(tk.Toplevel): # 继承Toplevel类
    def __init__(self,parent):
        super().__init__()
        self.title("新增用例")
        self.geometry('800x910+400+10') #设置窗口大小
        self.parent = parent # 显式地保留父窗口
        self.read_file = read_file()
        self.case_type = StringVar()
        self.model = StringVar()
        self.interName = StringVar()
        self.requestType = StringVar()
        self.Method = StringVar()
        self.case_ID = StringVar()
        self.model_list = []
        self.interface_list = []
        self.IPPort = StringVar()
        self.Case_Name = StringVar()
        self.Params = StringVar()
        self.Expect_Type = StringVar()
        #self.Expectation = StringVar()
        self.Check_field = StringVar()
        self.Http_Header = StringVar()
        self.Check_field_remove = StringVar()
        self.catch_result = StringVar()
        self.case_status = StringVar()
        self.case_header_list = []
        self.inter_Ename = ''  # 测试用例英文名称
        self.grab_set()
        self.resizable(False, True)
        self.createPage()

    def close_popup(self):
        # put focus back to the parent window
        self.parent.focus_set()
        self.destroy()


    def createPage(self):
        # create a canvas object and a vertical scrollbar for scrolling it
        vscrollbar = Scrollbar(self, orient=VERTICAL)
        vscrollbar.pack(fill=Y, side=RIGHT, expand=True)
        canvas = Canvas(self, bd=0, highlightthickness=0,
                        yscrollcommand=vscrollbar.set)
        canvas.pack(side=LEFT, fill=BOTH, expand=TRUE)
        vscrollbar.config(command=canvas.yview)

        # reset the view
        canvas.xview_moveto(0)
        canvas.yview_moveto(0)

        #self.labelframe = LabelFrame(self, text="新增用例",font=('Tempus Sams ITC',20))
        self.labelframe =interior= LabelFrame(canvas)
        self.labelframe.pack(fill="both", expand="yes")
        interior_id = canvas.create_window(0, 0, window=interior,
                                           anchor=NW)

        # track changes to the canvas and frame width and sync them,
        # also updating the scrollbar
        def _configure_interior(event):
            # update the scrollbars to match the size of the inner frame
            size = (interior.winfo_reqwidth(), interior.winfo_reqheight())
            canvas.config(scrollregion="0 0 %s %s" % size)
            if interior.winfo_reqwidth() != canvas.winfo_width():
                # update the canvas's width to fit the inner frame
                canvas.config(width=interior.winfo_reqwidth())
        interior.bind('<Configure>', _configure_interior)

        def _configure_canvas(event):
            if interior.winfo_reqwidth() != canvas.winfo_width():
                # update the inner frame's width to fill the canvas
                canvas.itemconfigure(interior_id, width=canvas.winfo_width())
        canvas.bind('<Configure>', _configure_canvas)

        # 从公共模块获取测试用例列表字段
        self.column = Ordered_Dict()
        column = self.column.test_case_header()
        column_list = list(column.values())

        # 布局字段label
        for value in column_list:

            i = column_list.index(value)
            label_column = ttk.Label(self.labelframe, text = value + "：")
            label_column.grid(row=i, stick=E,column=0, pady=8,padx=30)


        # 布局字段对应控件
        # 用例类型下拉框触发事件
        def select_model(event):
            self.model_Chosen.delete(0,END)
            self.interName_Chosen.delete(0,END)
            self.Method.set('')
            self.IPPort.set('')
            self.model_Chosen['values'] = []
            self.interName_Chosen['values'] = []
            model_list = []
            case_type_list = list(self.read_file.read_conf_file().col_values(0)[1:])
            if self.case_type.get() != '':
                for i in range(0,len(case_type_list)):
                    if case_type_list[i] == self.case_type.get():
                        model_list.append(self.read_file.read_conf_file().col_values(1)[i+1])

                model_list = list(set(model_list))
                self.model_list = model_list
                self.model_Chosen['values'] = self.model_list  # 设置下拉列表的值
            else:
                self.model_list = model_list
            return self.model_list

        # 用例类型下拉框
        self.case_type_list = list(set(self.read_file.read_conf_file().col_values(0)[1:]))
        self.case_type_Chosen = ttk.Combobox(self.labelframe,width=12,textvariable=self.case_type)
        self.case_type_Chosen['values'] = self.case_type_list  # 设置下拉列表的值
        self.case_type_Chosen.grid(row=0,column=1, stick=W,columnspan=2)
        self.case_type_Chosen.bind("<<ComboboxSelected>>",select_model)

        # 所属模块下拉框触发事件
        def select_interface(event):
            self.interName_Chosen.delete(0,END)
            self.Method.set('')
            self.IPPort.set('')
            self.interName_Chosen['values'] = []
            interface_list_i = []
            interface_list_k = []
            case_type_list = list(self.read_file.read_conf_file().col_values(0)[1:])
            model_list = list(self.read_file.read_conf_file().col_values(1)[1:])

            if self.model.get() != '' and self.case_type.get() != '':
                for i in range(0,len(case_type_list)):
                    for k in range(0,len(model_list)):
                        if case_type_list[i] == self.case_type.get() and model_list[k] == self.model.get():
                            interface_list_i.append(self.read_file.read_conf_file().col_values(2)[i+1])
                            interface_list_k.append(self.read_file.read_conf_file().col_values(2)[k+1])

                interface_list = list(set(interface_list_i).intersection(set(interface_list_k)))
                self.interface_list = interface_list
                self.interName_Chosen['values'] = self.interface_list  # 设置下拉列表的值
            else:
                return []
            return self.interface_list

        # 所属模块下拉框
        self.model_Chosen = ttk.Combobox(self.labelframe,width=12,textvariable=self.model)
        self.model_Chosen['values'] = self.model_list  # 设置下拉列表的值
        self.model_Chosen.grid(row=1,column=1, stick=W,columnspan=2)
        self.model_Chosen.bind("<<ComboboxSelected>>",select_interface)

        def select_data(event):

            index = self.read_file.read_conf_file().col_values(2).index(self.interName.get())
            self.Method.set(self.read_file.read_conf_file().col_values(6)[index])
            self.IPPort.set(self.read_file.read_conf_file().col_values(5)[index])
            self.inter_Ename = self.read_file.read_conf_file().col_values(3)[index]

        # 接口名称下拉框
        self.interName_Chosen = ttk.Combobox(self.labelframe,width=20,text=self.interName)
        self.interName_Chosen['values'] = self.interface_list  # 设置下拉列表的值
        self.interName_Chosen.grid(row=2,column=1, stick=W,columnspan=2)
        self.interName_Chosen.bind("<<ComboboxSelected>>",select_data)

        # 用例ID
        ttk.Label(self.labelframe,textvariable=self.case_ID).grid(row=3, stick=W, column=1,pady=10,columnspan=2)
        # 请求方式
        self.label_Method = ttk.Label(self.labelframe, textvariable=self.Method)
        self.label_Method.grid(row=4, stick=W, column=1,pady=10,columnspan=2)
        # 请求IP端口
        self.label_IPPort = ttk.Entry(self.labelframe,width=80,textvariable=self.IPPort)
        self.label_IPPort.grid(row=5, stick=W, column=1,pady=10,columnspan=2)
        # 用例描述
        ttk.Entry(self.labelframe, textvariable=self.Case_Name,width=80).grid(row=6,stick=W,column=1,pady=10,columnspan=2)
        # 请求包体
        self.Params_text = Text(self.labelframe,width=80,height=10)
        self.Params_text.grid(row=7, stick=W,column=1, pady=10,columnspan=2)

        # 期望结果类型
        Expect_Type = ttk.Combobox(self.labelframe,width=15,textvariable=self.Expect_Type)
        Expect_Type['values'] = ("single","multiple") # 设置下拉列表的值
        Expect_Type.grid(row=8,column=1, stick=W, pady=10,columnspan=2)
        Expect_Type.current(0) # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值

        def Check_field_Update():

            self.Check_field_text.delete('0.0', END)
            if self.Expectation.get('0.0', END) != '\n'or self.Expectation.get('0.0', END) != '':
                try:
                    Expectation = eval(self.Expectation.get('0.0', END).replace("\n"," "))
                    keys = Expectation.keys()
                    self.Check_field_text.insert(INSERT,str(list(keys)))
                except:
                    showinfo(title='错误', message='您输入的期望结果为非json格式，无法计算检查字段！')

        # 期望结果
        self.Expectation = Text(self.labelframe,width=80,height=10)
        self.Expectation.grid(row=9, stick=W,column=1, pady=10,columnspan=2)

        # 检查字段
        self.Check_field_text = Text(self.labelframe,width=60,height=2)
        self.Check_field_text.grid(row=10, stick=W, column=1,pady=10,columnspan=1)
        # 计算检查字段按钮
        ttk.Button(self.labelframe,text='计算检查字段',command=Check_field_Update).grid(row=10,stick=E,column=2,pady=10)
        # 非检查字段
        ttk.Entry(self.labelframe, textvariable=self.Check_field_remove,width=80).grid(row=11, stick=W, column=1,pady=10,columnspan=2)
        # 结果获取字段
        ttk.Entry(self.labelframe, textvariable=self.catch_result,width=80).grid(row=12, stick=W, column=1,pady=10,columnspan=2)

        # 检查写入数据库SQL
        self.checkDB_SQL = Text(self.labelframe,width=80,height=3)
        self.checkDB_SQL.grid(row=13, stick=W, column=1,pady=10,columnspan=2)
        # 检查写入数据库结果
        self.checkDB_result = Text(self.labelframe,width=80,height=3)
        self.checkDB_result.grid(row=14, stick=W,column=1, pady=10,columnspan=2)

        # 接口名称下拉框
        self.case_status_list = ttk.Combobox(self.labelframe,width=20,text=self.case_status)
        self.case_status_list['values'] = ['有效','无效']  # 设置下拉列表的值
        self.case_status_list.grid(row=15,column=1, stick=W,columnspan=2)
        self.case_status_list.current(0)
        # httpheader字段
        ttk.Entry(self.labelframe, textvariable=self.Http_Header,width=80).grid(row=16, stick=W, column=1,pady=10,columnspan=2)

        self.summit_button = ttk.Button(self.labelframe, text='保存',command=self.summit_data)
        self.summit_button.grid(row=999,column=1,sticky=W, pady=10)
        ttk.Button(self.labelframe, text='关闭',command=self.close_popup).grid(row=999,column=1,sticky=NE, pady=10,padx=0)

    def summit_data(self):
        column = self.column.test_case_header()
        w_xls,sheet_write,rows = self.read_file.write_case_file()

        if self.case_type.get() == '': # 用例类型为空
            showerror(title='错误', message=column['case_type'] + '不可为空')
            return
        elif self.model.get() == '': # 所属模块为空
            showerror(title='错误', message=column['model'] + '不可为空')
            return
        elif self.interName.get() == '':  # 接口名称为空
            showerror(title='错误', message=column['interfacename'] + '不可为空')
            return
        elif self.Method.get() == '':  # 请求方式为空
            showerror(title='错误', message=column['Method'] + '不可为空')
            return
        elif self.Case_Name.get() == '':  # 用例名称为空
            showerror(title='错误', message=column['Case_Name'] + '不可为空')
            return
        elif self.Params_text.get('0.0', END) in['\n','']: # 请求包体为空
            showerror(title='错误', message=column['Params'] + '不可为空')
            return
        elif self.Expect_Type.get() == '':  # 期望结果类型为空
            showerror(title='错误', message=column['Expect_Type'] + '不可为空')
            return
        elif self.Expectation.get('0.0', END) in['\n','']:  # 期望结果为空
            showerror(title='错误', message=column['Expectation'] + '不可为空')
            return
        elif self.case_status.get() == '':  # 状态为空
            showerror(title='错误', message=column['status'] + '不可为空')
            return
        elif self.Check_field_text.get('1.0', END) in['\n',''] and self.Check_field_remove.get() in['\n','']:  # 检查字段和非检查字段同时为空
            showerror(title='错误', message=column['Check_field'] + '和' + column['Check_field_remove'] + '不可同时为空')
            return
        elif self.checkDB_SQL.get('0.0', END).replace("\n","") != "" and self.checkDB_result.get('0.0', END).replace("\n","") == "":  # 检查字段和非检查字段同时为空
            showerror(title='错误', message=column['checkDB_SQL'] + '不为空时' + column['checkDB_result'] + '不可为空')
            return

        if self.Method.get() in ['POST','PUT','DELETE','GET']:
            try:
                P = eval(self.Params_text.get('0.0', END).replace("\n"," "))
                if isinstance(P,dict) is False:
                    showerror(title='错误', message="请求方法为" + self.Method.get() + "，" + column['Params'] + '需为json格式')
                    return
            except:
                showerror(title='错误', message="请求方法为" + self.Method.get() + "，" + column['Params'] + '需为json格式')
                return

        if self.Expect_Type.get() in ["single","multiple"]:
            try:
                E = eval(self.Expectation.get('0.0', END).replace("\n"," "))
                if isinstance(E,dict) is False:
                    showerror(title='错误', message=column['Expectation'] + '需为json格式')
                    return
            except:
                showerror(title='错误', message=column['Expectation'] + '需为json格式')
                return

        if self.Check_field_text.get('1.0', END) not in ['','\n']:
            try:
                eval(self.Check_field_text.get('1.0', END).replace("\n"," "))
                if self.Check_field_text.get('1.0', END).replace("\n","") == "ALL":
                    pass
                elif isinstance(eval(self.Check_field_text.get('1.0', END)),list) is False:
                    showerror(title='错误', message=column['Check_field'] + "需为列表格式，如['column1','column2']")
                    return
            except:
                showerror(title='错误', message=column['Check_field'] + "需为列表格式，如['column1','column2']")
                return

        if self.Check_field_remove.get() != '':
            try:
                eval(self.Check_field_remove.get().replace("\n"," "))
                if isinstance(eval(self.Check_field_remove.get()),list) is False:
                    showerror(title='错误', message=column['Check_field_remove'] + "需为列表格式，如['column1','column2']")
                    return
            except:
                showerror(title='错误', message=column['Check_field_remove'] + "需为列表格式，如['column1','column2']")
                return

        if self.catch_result.get() != '':
            try:
                CA = eval(self.catch_result.get().replace("\n"," "))
                if isinstance(CA,dict) is False:
                    showerror(title='错误', message=column['catch_result'] + '''需为json格式，如{"key1":"value1","key2":"value2"}''')
                    return
            except:
                showerror(title='错误', message=column['catch_result'] + '''需为json格式，如{"key1":"value1","key2":"value2"}''')
                return

        try:
            hh = eval(self.Http_Header.get().replace("\n"," "))
            if isinstance(hh,dict) is False:
                showerror(title='错误', message=column['HTTPHeader'] + '''需为json格式，如{"key1":"value1","key2":"value2"}''')
                return
        except:
            showerror(title='错误', message=column['HTTPHeader'] + '''需为json格式，如{"key1":"value1","key2":"value2"}''')
            return

        test_case_id = self.add_case_id()
        # 逐列写入数据
        self.case_header_list = [self.case_type.get(),self.model.get(),self.interName.get(),test_case_id,self.Method.get(),self.IPPort.get(),self.Case_Name.get()
                                ,self.Params_text.get('0.0', END).replace('\n',' ').strip(),self.Expect_Type.get(),self.Expectation.get('0.0', END).replace('\n',' ').strip(),
                                self.Check_field_text.get('1.0', END).replace('\n',' ').strip(),self.Check_field_remove.get(),self.catch_result.get(),
                                 self.checkDB_SQL.get('0.0', END).replace('\n',' ').strip(),self.checkDB_result.get('0.0', END).replace('\n',' ').strip(),self.case_status.get(),self.Http_Header.get()]
        try:
            for i in range(0,len(self.case_header_list)):
                sheet_write.write(rows,i,self.case_header_list[i])
            w_xls.save(self.read_file.save_case_file)
            showinfo(title="信息",message="用例新增成功！")
            self.destroy()
            self.parent.search()

        except BaseException as e:
            showerror(title='错误', message="用例保存失败！%s"%e)

    # 计算新增测试用例递增ID
    def add_case_id(self):
        case_list = self.read_file.read_case_file().col_values(2)[1:]
        num_temp = []
        test_case_id_list = []
        for i in range(0,len(case_list)):
            if self.interName.get() == case_list[i]:
                num_temp.append(i+1)
        if num_temp != []:
            for k in num_temp:
                test_case_id_list.append(self.read_file.read_case_file().col_values(3)[k])

            test_case_id_ori = max(test_case_id_list)
            test_case_id_ori = test_case_id_ori.split('_')
            test_case_id_ori[1] = int(test_case_id_ori[1]) + 1
            test_case_id_ori[1] = "%03d" % int(test_case_id_ori[1])  #前面补0，保留3位
            test_case_id_new = self.inter_Ename + '_' + test_case_id_ori[1]
        else:
            test_case_id_new = self.inter_Ename + '_' + '001'
        return test_case_id_new