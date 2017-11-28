from tkinter import ttk
import tkinter as tk
from View.CaseFrame import *
from tkinter import scrolledtext        # 导入滚动文本框的模块
import time

class Edit_task(tk.Toplevel): # 继承Toplevel类
    def __init__(self,parent,selected_task):
        super().__init__()
        self.title("编辑任务")
        self.geometry('1200x900+200+20') #设置窗口大小
        self.parent = parent # 显式地保留父窗口
        self.read_file = read_file()
        self.task_name = StringVar()
        self.email = StringVar()
        self.case_type = StringVar()
        self.model = StringVar()
        self.interface = StringVar()
        self.case_id = StringVar()
        self.exe_case_id = StringVar()
        self.case_type_list = []
        self.model_list = []
        self.interface_list = []
        self.caseID_list = []
        self.select_item_intername = []
        self.select_item_caseID = []
        self.grab_set()
        self.selected_task = selected_task
        self.resizable(False, True)
        self.createPage()

    def close_popup(self):
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

        # 任务名称
        ttk.Label(self.labelframe,text='任务名称：').grid(row=0, stick=E, column=0,padx=30,pady=20)
        ttk.Entry(self.labelframe, textvariable=self.task_name,width=130).grid(row=0, stick=W, column=1,pady=20,columnspan=7)
        self.task_name.set(self.selected_task[1])

        # 任务描述
        ttk.Label(self.labelframe,text='任务描述：').grid(row=1, stick=E, column=0,pady=5,padx=30)
        self.Task_description = tk.Text(self.labelframe,width=130,height=2)
        self.Task_description.grid(row=1, stick=W, column=1,pady=5,columnspan=7)
        self.Task_description.insert("end",self.selected_task[2])

        # 邮件发送
        ttk.Label(self.labelframe,text='邮件发送：').grid(row=2, stick=E, column=0,pady=5,padx=30)
        self.email_option = ttk.Combobox(self.labelframe,width=15,textvariable=self.email)
        self.email_option['values'] = ('是','否') # 设置下拉列表的值
        self.email_option.grid(row=2,column=1, stick=W, pady=5,columnspan=2)
        self.email.set(self.selected_task[7])

        # 任务初始化数据操作
        ttk.Label(self.labelframe,text='任务初始化数据SQL：').grid(row=3, stick=E, column=0,pady=5,padx=30)
        self.setup_SQL = Text(self.labelframe,width=130,height=4)
        self.setup_SQL.grid(row=3, stick=W, column=1,pady=5,columnspan=7)
        self.setup_SQL.insert("end",self.selected_task[5])

        # 任务结束清除数据操作
        ttk.Label(self.labelframe,text='任务结束清除数据SQL：').grid(row=4, stick=E, column=0,pady=5,padx=30)
        self.teardown_SQL = Text(self.labelframe,width=130,height=4)
        self.teardown_SQL.grid(row=4, stick=W, column=1,pady=5,columnspan=7)
        self.teardown_SQL.insert("end",self.selected_task[6])

        def select_model_list(event):
            self.case_id.set([])
            self.interface.set([])
            self.model.set([])
            self.select_item_caseID = []
            select_item_sum = eval(self.case_type.get()) # listbox所有的值
            select_index = self.Case_Type.curselection() # listbox所选的ITEM序号
            select_item = []
            temp = []
            model_list = []
            for index in select_index:
                select_item.append(select_item_sum[index])
            for i in select_item:
                case_type_list = self.read_file.read_conf_file().col_values(0)[1:]
                for k in range(0,len(case_type_list)):
                    if case_type_list[k] == i:
                        temp.append(k+1)
            for k in temp:
                model_list.append(self.read_file.read_conf_file().col_values(1)[k])
            self.model_list = list(set(model_list))
            self.model_list.sort()
            self.Model.delete(0,END)
            for item in self.model_list:
                self.Model.insert(END,item)

        # 用例类型列表
        ttk.Label(self.labelframe,text='用例类型').grid(row=5, stick=W, column=0,pady=5,padx=30)
        self.case_type_list = list(set(self.read_file.read_conf_file().col_values(0)[1:]))
        self.Case_Type = Listbox(self.labelframe,listvariable = self.case_type,selectmode = SINGLE,height=30,borderwidth=0, selectborderwidth=0)
        for item in self.case_type_list:
            self.Case_Type.insert(END,item)
        self.Case_Type.grid(row=6,column=0, stick=N, pady=2,padx=30)
        self.Case_Type.bind("<<ListboxSelect>>",select_model_list)

        def select_interface_list(event):
            self.case_id.set([])
            self.interface.set([])
            self.select_item_caseID = []
            select_item_sum = eval(self.model.get()) # listbox所有的值
            select_index = self.Model.curselection() # listbox所选的ITEM序号
            select_item = []
            temp = []
            interface_list = []
            for index in select_index:
                select_item.append(select_item_sum[index])
            for i in select_item:
                model_list = self.read_file.read_conf_file().col_values(1)[1:]
                for k in range(0,len(model_list)):
                    if model_list[k] == i:
                        temp.append(k+1)
            for k in temp:
                interface_list.append(self.read_file.read_conf_file().col_values(2)[k])
            self.interface_list = list(set(interface_list))
            self.interface_list.sort()
            self.inter_name.delete(0,END)
            for item in self.interface_list:
                self.inter_name.insert(END,item)
        # 所属模块列表
        ttk.Label(self.labelframe,text='所属模块').grid(row=5, stick=W, column=1,pady=5,padx=30)
        self.Model = Listbox(self.labelframe,selectmode = EXTENDED,listvariable = self.model,height=30,borderwidth=0, selectborderwidth=0)
        for item in self.model_list:
            self.Model.insert(END,item)
        self.Model.grid(row=6,column=1, stick=N, pady=2,padx=30)
        self.Model.bind("<<ListboxSelect>>",select_interface_list)

        interface_list = self.read_file.read_case_file().col_values(2)[1:]
        status_list = self.read_file.read_case_file().col_values(15)[1:]
        case_list1 = self.read_file.read_case_file().col_values(3)
        casename_list = self.read_file.read_case_file().col_values(6)
        def select_case_list(event):
            self.case_id.set([])
            self.select_item_caseID = []
            if self.interface.get() == '':
                select_item_sum = ()
            else:
                select_item_sum = eval(self.interface.get()) # listbox所有的值
            select_index = self.inter_name.curselection() # listbox所选的ITEM序号
            self.select_item_intername = []
            temp = []
            case_list = []
            status = []

            for index in select_index:
                self.select_item_intername.append(select_item_sum[index])
            for i in self.select_item_intername:

                for k,interface in enumerate(interface_list):
                    if interface == i:
                        temp.append(k+1)

            for q in range(0,len(status_list)):  # 过滤有效的case_ID
                if status_list[q] == '有效':
                    status.append(q+1)
                status.sort()
            temp = list(set(temp).intersection(status))

            for k in temp:
                case_list.append(case_list1[k]+"***"+casename_list[k])

            self.caseID_list = list(set(case_list))
            self.caseID_list.sort()
            self.caseID.delete(0,END)

            for item in self.caseID_list:
                self.caseID.insert(END,item)

        # 接口名称列表
        ttk.Label(self.labelframe,text='接口名称').grid(row=5, stick=W, column=2,pady=5,padx=30)
        self.inter_name = Listbox(self.labelframe,listvariable = self.interface,selectmode = EXTENDED,height=30,borderwidth=0, selectborderwidth=0)
        for item in self.interface_list:
            self.inter_name.insert(END,item)
        self.inter_name.grid(row=6,column=2, stick=N, pady=2,padx=30)
        self.select_item_intername = eval(self.selected_task[3])
        self.inter_name.bind("<<ListboxSelect>>",select_case_list)

        # 用例ID列表
        ttk.Label(self.labelframe,text='用例列表').grid(row=5, stick=W, column=3,pady=5,padx=30)
        self.caseID = Listbox(self.labelframe,listvariable = self.case_id,selectmode = EXTENDED,height=30,width=30,borderwidth=0, selectborderwidth=0)
        for item in self.caseID_list:
            self.caseID.insert(END,item)
        self.caseID.grid(row=6,column=3, stick=N, pady=2,padx=0)
        def add_case(event):
            self.add_case()
        self.caseID.bind('<Double-Button-1>', add_case)

        vbar = ttk.Scrollbar(self.labelframe,orient=VERTICAL,command=self.caseID.yview)
        self.caseID.configure(yscrollcommand=vbar.set)
        vbar.grid(row=6,column=4,pady=2,sticky=NS)

        self.add_case_button = tk.Button(self.labelframe, text='  →  ',width=7,height=1,command=self.add_case)
        self.add_case_button.grid(row=6,column=5,sticky=N, pady=230)
        self.delete_case_button = tk.Button(self.labelframe, text='  ←  ',width=7,height=1,command=self.delete_case)
        self.delete_case_button.grid(row=6,column=5,sticky=S, pady=230,padx=3)

        # 待执行用例ID列表
        ttk.Label(self.labelframe,text='待执行用例ID列表').grid(row=5, stick=W, column=6,pady=5,padx=30)
        self.exe_caseID = Listbox(self.labelframe,listvariable = self.exe_case_id,selectmode = EXTENDED,height=30,borderwidth=0, selectborderwidth=0)
        self.exe_caseID.grid(row=6,column=6, stick=N, pady=2,padx=0)
        for item in eval(self.selected_task[4]):
            self.exe_caseID.insert("end",item)
        def delete_case(event):
            self.delete_case()
        self.exe_caseID.bind('<Double-Button-1>', delete_case)

        def move_case_up(event):
            self.move_case_up()
        def move_case_down(event):
            self.move_case_down()
        self.move_case_up_button = tk.Button(self.labelframe, text=' ↑ ',command=self.move_case_up)
        self.move_case_up_button.grid(row=6,column=8,sticky=N, pady=230,padx=3)
        self.delete_case_button = tk.Button(self.labelframe, text=' ↓ ',command=self.move_case_down)
        self.delete_case_button.grid(row=6,column=8,sticky=S, pady=230,padx=3)

        self.exe_caseID.bind('<Left>',move_case_up)
        self.exe_caseID.bind('<Right>',move_case_down)
        self.move_case_up_button.bind('<Left>',move_case_up)
        self.move_case_up_button.bind('<Right>',move_case_down)
        self.delete_case_button.bind('<Left>',move_case_up)
        self.delete_case_button.bind('<Right>',move_case_down)

        vbar = ttk.Scrollbar(self.labelframe,orient=VERTICAL,command=self.exe_caseID.yview)
        self.exe_caseID.configure(yscrollcommand=vbar.set)
        vbar.grid(row=6,column=7,pady=2,sticky=NS)

        self.summit_button = ttk.Button(self.labelframe, text='保存',command=self.summit_data)
        self.summit_button.grid(row=999,column=2,sticky=W, pady=10)
        ttk.Button(self.labelframe, text='关闭',command=self.close_popup).grid(row=999,column=2,sticky=NE, pady=10,padx=0)

    def summit_data(self):
        self.column = Ordered_Dict()
        column = self.column.task_header()
        w_xls,sheet_write,rows = self.read_file.write_task_file()

        if self.task_name.get() == '':  # 任务名称为空
            showerror(title='错误', message=column['task_name'] + '不可为空')
            return
        elif self.Task_description.get('1.0',END) in ['','\n']:  # 任务描述为空
            showerror(title='错误', message=column['task_description'] + '不可为空')
            return
        elif self.email.get() == '':  # 邮件发送为空
            showerror(title='错误', message=column['task_email_option'] + '不可为空')
            return
        elif len(self.select_item_intername) == 0:  # 执行用例ID为空
            showerror(title='错误', message=column['task_interface'] + '未选择')
            return
        elif self.exe_caseID.size() == 0:  # 执行用例ID为空
            showerror(title='错误', message=column['task_caseID'] + '未选择')
            return

        self.new_task_list = [self.selected_task[0],self.task_name.get(),self.Task_description.get('1.0',END),self.select_item_intername,
                              list(self.exe_caseID.get(0,END)),self.setup_SQL.get('1.0',END).replace("\n"," ").strip(),
                              self.teardown_SQL.get('1.0',END).replace("\n"," ").strip(),self.email.get(),self.selected_task[8]]

        taskid_list = self.read_file.read_task().col_values(0)
        edit_row = taskid_list.index(self.selected_task[0])
        try:
            for i in range(0,len(self.new_task_list)):
                sheet_write.write(edit_row,i,str(self.new_task_list[i]))
            w_xls.save(self.read_file.save_task_file)
            showinfo(title="信息",message="保存任务成功！")
            self.destroy()

            # 重新加载主页面treeview数据
            ori_items = self.parent.tree.get_children()
            [self.parent.tree.delete(item) for item in ori_items]

            case_sheet = self.read_file.read_task()
            search_task_list = range(1,case_sheet.nrows)
            for index in search_task_list:
                self.parent.tree.insert('',-index,values=case_sheet.row_values(index))
        except:
            showerror(title='错误', message="任务保存失败！")

    def add_case(self):

        select_item_sum = eval(self.case_id.get()) # listbox所有的值
        select_index = self.caseID.curselection() # listbox所选的ITEM序号
        self.select_item_caseID = []
        for index in select_index:
            select_item = select_item_sum[index].split("***")
            self.select_item_caseID.append(select_item[0])
        for item in self.select_item_caseID:
            #if item not in self.exe_caseID.get(0,END):
            self.exe_caseID.insert(END,item)

    def delete_case(self):

        select_index = self.exe_caseID.curselection() # listbox所选的ITEM序号
        exe_caseID_list = list(self.exe_caseID.get(0,END))
        remove_list = []
        [remove_list.append(exe_caseID_list[i]) for i in select_index]
        [exe_caseID_list.remove(value) for value in remove_list]
        self.exe_caseID.delete(0,END)
        [self.exe_caseID.insert("end",item) for item in exe_caseID_list]

    def move_case_up(self):

        select_index = self.exe_caseID.curselection()  # listbox所选的ITEM序号
        if select_index != () and select_index[0] != 0:
            exe_caseID_list = list(self.exe_caseID.get(0,END))
            item = exe_caseID_list[select_index[0]]
            exe_caseID_list.remove(item)
            exe_caseID_list.insert(select_index[0]-1,item)
            self.exe_caseID.delete(0,END)
            [self.exe_caseID.insert("end",item) for item in exe_caseID_list]
            self.exe_caseID.selection_set(select_index[0]-1,select_index[0]-1)
            self.exe_caseID.see(select_index[0]-1)

    def move_case_down(self):

        select_index = self.exe_caseID.curselection()  # listbox所选的ITEM序号
        if select_index != ():
            exe_caseID_list = list(self.exe_caseID.get(0,END))
            item = exe_caseID_list[select_index[0]]
            exe_caseID_list.remove(item)
            exe_caseID_list.insert(select_index[0]+1,item)
            self.exe_caseID.delete(0,END)
            [self.exe_caseID.insert("end",item) for item in exe_caseID_list]
            if select_index[0]+1 < len(exe_caseID_list):
                self.exe_caseID.selection_set(select_index[0]+1,select_index[0]+1)
            else:
                self.exe_caseID.selection_set(END,END)
            self.exe_caseID.see(select_index[0]+1)
