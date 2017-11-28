import tkinter as tk
from tkinter import *
from tkinter import ttk
from Control.read_file import read_file
from View.CaseFrame import *
import json

class EditCase_View(tk.Toplevel): # 继承Toplevel类
    def __init__(self,view_item,parent):
        super().__init__()
        self.title("修改用例")
        self.geometry('820x845+400+10') #设置窗口大小
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
        self.Params = StringVar()
        self.Expect_Type = StringVar()
        self.Check_field = StringVar()
        self.Check_field_remove = StringVar()
        self.case_status = StringVar()
        self.statues = StringVar()
        self.Http_Header = StringVar()
        self.case_header_list = []
        self.inter_Ename = ''     #测试用例英文名称
        self.grab_set()
        self.resizable(False, True)
        self.createPage(view_item)


    def close_popup(self):
        self.destroy()


    def createPage(self,view_item):

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


        #布局label
        for value in column_list:

            i = column_list.index(value)
            label_column = Label(self.labelframe, text=value + "：")
            label_column.grid(row=i, stick=E,column=0, pady=7,padx=40)

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
            # 用例ID
        ttk.Label(self.labelframe,textvariable=self.case_ID).grid(row=3, stick=W, column=1,pady=7,columnspan=2)
        self.case_ID.set(view_item[3])
            # 请求方式
        self.label_Method = ttk.Label(self.labelframe, textvariable=self.Method)
        self.Method.set(view_item[4])
        self.label_Method.grid(row=4, stick=W, column=1,pady=7,columnspan=2)
            # 请求IP端口
        self.label_IPPort = ttk.Entry(self.labelframe, textvariable=self.IPPort,width=80)
        self.IPPort.set(view_item[5])
        self.label_IPPort.grid(row=5, stick=W, column=1,pady=7,columnspan=2)
            #用例名称
        self.Case_Name = tk.Text(self.labelframe, width=80,height=1)
        self.Case_Name.insert(INSERT,view_item[6])
        self.Case_Name.grid(row=6, stick=W, column=1,pady=7,columnspan=2)
            #请求包体
        self.Params_text = Text(self.labelframe,width=80,height=10)
        self.Params_text.insert(INSERT,view_item[7])
        self.Params_text.grid(row=7, stick=W,column=1, pady=7, padx=0)
            #期望结果类型
        Expect_Type = ttk.Combobox(self.labelframe,width=30,textvariable=self.Expect_Type)
        Expect_Type['values'] = ('single','multiple') # 设置下拉列表的值
        Expect_Type.grid(row=8,column=1, stick=W, pady=7,columnspan=2)
        self.Expect_Type.set(view_item[8])
        #Expect_Type.current(0) # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值
            #期望结果
        self.Expectation = Text(self.labelframe,width=80,height=10)
        self.Expectation.insert(INSERT,view_item[9])
        self.Expectation.grid(row=9, stick=W,column=1, pady=7,padx=0)
            #检查字段
        #Entry(self.labelframe, width=50, textvariable=self.CkeckItem).grid(row=10, column=2, stick=W)
        self.CkeckItem = Text(self.labelframe,width=80,height=2)
        self.CkeckItem.insert(INSERT,view_item[10])
        self.CkeckItem.grid(row=10, stick=W, column=1,pady=10,columnspan=1)
            #非检查字段
        #Entry(self.labelframe, width=50, textvariable=self.NCkeckItem).grid(row=11, column=2, stick=W)
        self.NCkeckItem = Text(self.labelframe,width=80,height=2)
        self.NCkeckItem.insert(INSERT,view_item[11])
        self.NCkeckItem.grid(row=11, stick=W, column=1,pady=10,columnspan=1)
            #结果获取字段
        self.catch_result = tk.Text(self.labelframe, width=80,height=1)
        self.catch_result.insert(INSERT,view_item[12])
        self.catch_result.grid(row=12, column=1, stick=W)

        # 检查写入数据库SQL
        #self.checkDB_SQL_list = ttk.Entry(self.labelframe, textvariable=self.checkDB_SQL,width=80).grid(row=13, stick=W, column=1)
        self.checkDB_SQL = Text(self.labelframe, width=80,height=3)
        self.checkDB_SQL.insert(INSERT,view_item[13])
        self.checkDB_SQL.grid(row=13, column=1, stick=W)

        # 检查数据库结果
        self.checkDB_result = Text(self.labelframe, width=80,height=3)
        self.checkDB_result.insert(INSERT,view_item[14])
        self.checkDB_result.grid(row=14, column=1, stick=W,pady=7)
            #用例状态
        statues = ttk.Combobox(self.labelframe,width=30,textvariable=self.statues)
        statues['values'] = ('有效','无效') # 设置下拉列表的值
        statues.grid(row=15, column=1, stick=W, pady=10)
        self.statues.set(view_item[15])
        #statues.current(0) # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值

        # 鉴权字段
        Http_Header = ttk.Entry(self.labelframe, textvariable=self.Http_Header,width=80)
        self.Http_Header.set(view_item[16])
        Http_Header.grid(row=16, stick=W, column=1,pady=10,columnspan=2)
            #保存/关闭
        ttk.Button(self.labelframe, text='保存', width=8, command=self.summit_data).grid(row=999,column=0,columnspan=1,sticky=E)
        ttk.Button(self.labelframe, text='关闭', width=8, command=self.close_popup).grid(row=999,column=1,columnspan=1,sticky=W,padx=80)

    def Editrow(self):

        search_item = self.case_ID.get()

            # 读取EXCEL用例数据到列表中
        case_list = self.read_file.read_case_file().col_values(3)[1:]
        for k in range(0,len(case_list)):
            if case_list[k] == search_item:
                return k+1

    def summit_data(self):
        column = self.column.test_case_header()
        w_xls,sheet_write,rows = self.read_file.write_case_file()

        if self.Case_Name.get('1.0',END) in ['','\n']:  #用例名称为空
            showerror(title='错误', message=column['Case_Name'] + '不可为空')
            return
        elif self.Params_text.get('0.0',END) in ['\n','']: # 请求包体为空
            showerror(title='错误', message=column['Params'] + '不可为空')
            return
        elif self.Expect_Type.get() == '':  # 期望结果类型为空
            showerror(title='错误', message=column['Expect_Type'] + '不可为空')
            return
        elif self.Expectation.get('0.0', END) in ['\n','']:  # 期望结果为空
            showerror(title='错误', message=column['Expectation'] + '不可为空')
            return
        elif self.CkeckItem.get('1.0', END) in['\n',''] and self.NCkeckItem.get('1.0', END) in['\n','']:  # 检查字段和非检查字段同时为空
            showerror(title='错误', message=column['Check_field'] + '和' + column['Check_field_remove'] + '不可同时为空')
            return
        if self.Method.get() in ['POST','PUT','DELETE','GET']:
            try:
                if isinstance(eval(self.Params_text.get('0.0',END).replace("\n"," ")),dict) is False:
                    showerror(title='错误', message="请求方法为" + self.Method.get() + "，" + column['Params'] + '需为json格式')
                    return
            except:
                showerror(title='错误', message="请求方法为" + self.Method.get() + "，" + column['Params'] + '需为json格式')
                return
        if self.Expect_Type.get() in['single','multiple'] :
            try:
                if isinstance(eval(self.Expectation.get('0.0',END).replace("\n"," ")),dict) is False:
                    showerror(title='错误', message=column['Expectation'] + '需为json格式')
                    return
            except:
                showerror(title='错误', message=column['Expectation'] + '需为json格式')
                return

        if self.CkeckItem.get('1.0',END) not in ['','\n']:
            try:
                if self.CkeckItem.get('1.0', END).replace("\n","") == "ALL":
                    pass
                elif isinstance(eval(self.CkeckItem.get('1.0',END)),list) is False:
                    showerror(title='错误',message=column['Check_field'] + '需要为列表格式')
                    return
            except:
                showerror(title='错误',message=column['Check_field'] + '需要为列表格式')
                return

        if self.NCkeckItem.get('1.0',END) not in ['','\n']:
            try:
                if isinstance(eval(self.NCkeckItem.get('1.0',END)),list) is False:
                    showerror(title='错误',message=column['Check_field_remove'] + '需要为列表格式')
            except:
                showerror(title='错误',message=column['Check_field_remove'] + '需要为列表格式')

        if self.catch_result.get('1.0',END) not in ['','\n']:
            try:
                if isinstance(eval(self.catch_result.get('1.0',END).replace("\n"," ")),dict) is False:
                    showerror(title='错误',message=column['catch_result'] + '需为json格式')
                    return
            except:
                showerror(title='错误',message=column['catch_result'] + '需为json格式')
                return
        try:
            hh = eval(self.Http_Header.get())
            if isinstance(hh,dict) is False:
                showerror(title='错误', message=column['HTTPHeader'] + '''需为json格式，如{"key1":"value1","key2":"value2"}''')
                return
        except:
            showerror(title='错误', message=column['HTTPHeader'] + '''需为json格式，如{"key1":"value1","key2":"value2"}''')
            return

                # 逐列写入数据
        self.case_header_list = [self.IPPort.get(),self.Case_Name.get('1.0', END).replace('\n',''), self.Params_text.get('0.0', END).replace('\n',' ').strip(), self.Expect_Type.get(),
                                 self.Expectation.get('0.0', END).replace('\n',' ').strip(),self.CkeckItem.get('1.0', END).replace('\n','').strip(), self.NCkeckItem.get('1.0', END).replace('\n','').strip(),
                                 self.catch_result.get('1.0',END).replace('\n','').strip(),self.checkDB_SQL.get('0.0', END).replace('\n',' ').strip(),self.checkDB_result.get('0.0', END).replace('\n',' ').strip(),
                                 self.statues.get(),self.Http_Header.get()]

        rows = self.Editrow()
        try:
            for i in range(0,len(self.case_header_list)):
                sheet_write.write(rows, i+5 ,self.case_header_list[i])
            w_xls.save(self.read_file.save_case_file)
            showinfo(title="信息",message="用例修改成功！")
            self.destroy()
            self.parent.search()
        except BaseException as e:
            showerror(title='错误', message="用例保存失败！%s"%e)

