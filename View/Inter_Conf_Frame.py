from View.LoginPage import *
from tkinter import ttk
from Control.read_file import read_file
from Control.Ordered_Dict import Ordered_Dict
from View.View_Conf import View_Conf
from View.New_Conf import New_Conf
from View.Edit_Conf import Edit_Conf

class Inter_Conf_Frame(Frame): # 继承Frame类
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.root = master #定义内部变量root
        self.case_type = StringVar()
        self.model = StringVar()
        self.interName = StringVar()
        self.inter_Method = StringVar()
        self.inter_Status = StringVar()
        self.selected_item = ()
        self.read_file = read_file()
        self.createPage()

    def logout(self):
        Inter_Conf_Frame.quit(self)

    def createPage(self):

        self.labelframe = LabelFrame(self.root, text="接口配置",font=('Tempus Sams ITC',20))
        self.labelframe.pack(fill="both", expand="yes")
        def search(event):
            self.search()

        ttk.Label(self.labelframe, text = '用例类型: ').grid(row=1, stick=E,column=1, pady=10,padx=10)
        case_type_Chosen = ttk.Combobox(self.labelframe,width=12,textvariable=self.case_type)
        case_type_Chosen['values'] = list(set(self.read_file.read_conf_file().col_values(0)[1:])) # 设置下拉列表的值
        case_type_Chosen.grid(row=1,column=2, stick=W)
        case_type_Chosen.bind("<Return>",search)

        ttk.Label(self.labelframe, text = '所属模块: ').grid(row=1, stick=E, column=3,pady=10,padx=10)
        model_Chosen = ttk.Combobox(self.labelframe,width=20,textvariable=self.model)
        model_Chosen['values'] = list(set(self.read_file.read_conf_file().col_values(1)[1:])) # 设置下拉列表的值
        model_Chosen.grid(row=1,column=4, stick=W)
        model_Chosen.bind("<Return>",search)

        ttk.Label(self.labelframe, text = '接口名称: ').grid(row=1, stick=E, column=5,pady=10,padx=10)
        interName_Chosen = ttk.Combobox(self.labelframe,width=20,textvariable=self.interName)
        interName_Chosen['values'] = list(set(self.read_file.read_conf_file().col_values(2)[1:])) # 设置下拉列表的值
        interName_Chosen.grid(row=1,column=6, stick=W)
        interName_Chosen.bind("<Return>",search)

        # 配置状态下拉框
        ttk.Label(self.labelframe, text='接口状态: ').grid(row=1, stick=E, column=7, pady=10, padx=10)
        inter_Status = ttk.Combobox(self.labelframe, width=12, textvariable=self.inter_Status)
        inter_Status['values'] = ('', '有效', '无效')  # 设置下拉列表的值
        inter_Status.grid(row=1, column=8, stick=W)
        inter_Status.current(1)  # 设置下拉列表默认显示的值
        inter_Status.bind("<Return>",search)

        ttk.Button(self.labelframe, text='搜索',command=self.search).grid(row=1,column=9,padx=30,stick=W)

        ttk.Button(self.labelframe, text='登出',command=self.logout).grid(row=1,column=10,padx=10,stick=E)
        ttk.Button(self.labelframe, text='新增配置', command=self.new_conf).grid(row=2, stick=W,column=1,padx=10)
        ttk.Button(self.labelframe, text='编辑配置',command=self.edit_conf).grid(row=2,column=2, padx=10,stick=W)
        ttk.Button(self.labelframe, text='删除配置',command=self.del_conf).grid(row=2,column=3, padx=10, stick=W)

        self.column = Ordered_Dict()
        column = self.column.inter_conf_header()

        style = ttk.Style(self.labelframe)
        style.configure('Treeview',rowheight=20,relief='raised')

        self.tree = ttk.Treeview(self.labelframe,height=40,show="headings",columns=tuple(column.keys()))

        for i in column.keys():
            self.tree.column(i, width=200, anchor='w')
        self.tree.column('host', width=300, anchor='w')

        for k,v in column.items():
            self.tree.heading(k,text=v,command=lambda c=k : self.sortby(c, 0))

        vbar = ttk.Scrollbar(self.labelframe,orient=VERTICAL,command=self.tree.yview)
        self.tree.configure(yscrollcommand=vbar.set)

        # 读取EXCEL用例数据到列表中
        # case_sheet = self.read_file.read_conf_file()
        # case_list = case_sheet.col_values(0)[1:]
        # for index in range(1,len(case_list)+1):
        #     self.tree.insert('',index,values=case_sheet.row_values(index))
        self.search()
        self.tree.grid(row=3,column=1,columnspan=10,pady=20,sticky=NSEW)
        vbar.grid(row=3,column=11,pady=20,sticky=NS)

        def treeviewClick(event):
            item = self.tree.selection()[0]
            selected_item = self.tree.item(item, "values")
            popup_view_case = View_Conf(selected_item)
            self.wait_window(popup_view_case)

        def selectAll(event):
            ori_items = self.tree.get_children()  # 获取Treeview的所有元素
            self.tree.selection_set(ori_items)  # 选中多行数据， ori_items，即以I开头的行的id编码的元组。

        self.tree.bind('<Double-Button-1>', treeviewClick)
        self.tree.bind('<Control-a>', selectAll)  # 实现用Ctrl+a可选中Treeview中的所有元素

    def edit_conf(self):
        item = self.tree.selection()
        # print(item)
        if item in ['',()]:
            showerror(title='错误', message='请选中编辑内容！')
        else:
            self.selected_item = self.tree.item(item[0], 'values')
            # print(self.selected_item)
            selected_item = self.selected_item
            pop = Edit_Conf(selected_item, self)
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
        search_item = [self.case_type.get(),self.model.get(),self.interName.get(),'','','','',self.inter_Status.get()]
        conf_sheet = self.read_file.read_conf_file()

        search_conf_list = range(1,self.read_file.read_conf_file().nrows)
        temp_list = []

        # 读取EXCEL用例数据到列表中
        for i in range(0,len(search_item)):
            conf_list = self.read_file.read_conf_file().col_values(i)[1:]
            for k in range(0,len(conf_list)):
                if -1 != conf_list[int(k)].find(search_item[i].strip()):
                    temp_list.append(k+1)

            search_conf_list = list(set(search_conf_list).intersection(set(temp_list)))
            temp_list = []

        search_conf_list.sort()

        for index in search_conf_list:
            self.tree.insert('',index,values=conf_sheet.row_values(index))

    def new_conf(self):

        popup_new_conf = New_Conf(self)
        self.wait_window(popup_new_conf)

    def del_conf(self):
        if self.tree.selection() in ['',()]:
            showerror(title='错误', message='请选中内容！')
        else:
            items = self.tree.selection()
            confirm = askyesno(title="确认", message="确认删除所选接口配置？")
            if confirm is True:
                w_xls,sheet_write,rows = self.read_file.write_conf_file()
                conf_list = self.read_file.read_conf_file().col_values(3)[1:]
                for item in items:
                    selected_item = self.tree.item(item, "values")
                    for i in range(0,len(conf_list)):
                        if selected_item[3] == conf_list[i]:
                            sheet_write.write(i+1,7,'无效')
                w_xls.save(self.read_file.save_conf_file)
                showinfo(title="信息", message="用例配置成功！")
                self.search()
            # else:
            #     showinfo(title="信息",message="用例为无效状态")
