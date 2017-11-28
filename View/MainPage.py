from tkinter import *
from View.TaskFrame import *  #菜单栏对应的各个子页面
from View.Inter_Conf_Frame import *
from View.CaseFrame import *
from View.AboutFrame import *

class MainPage(object):
    def __init__(self, master=None):
        self.root = master #定义内部变量root
        w, h = self.root.maxsize()
        self.root.geometry("{}x{}".format(w, h))
        #self.root.geometry('1500x900+400+30') #设置窗口大小
        self.createPage()

    def createPage(self):

        self.casePage = CaseFrame(self.root) # 创建不同Frame
        self.taskPage = TaskFrame(self.root)
        self.interPage = Inter_Conf_Frame(self.root)
        self.aboutPage = AboutFrame(self.root)

        self.casePage.labelframe.pack() #默认显示数据录入界面
        self.taskPage.labelframe.pack_forget()
        self.interPage.labelframe.pack_forget()
        self.aboutPage.pack_forget()

        menubar = Menu(self.root)
        inter_menu = Menu(menubar, tearoff=0)

        menubar.add_cascade(label="接口测试", menu=inter_menu)
        inter_menu.add_command(label='用例列表', command = self.TestCaseManagment)
        inter_menu.add_command(label='任务列表', command = self.TaskManagement)
        inter_menu.add_command(label='接口配置', command = self.Interface_Managment)

        self.root['menu'] = menubar  # 设置菜单栏
        menubar.add_command(label='关于', command = self.aboutDisp)

        menubar.add_command(label='WEB测试')
        menubar.add_command(label='APP测试')
        menubar.add_command(label='性能测试')

    def TestCaseManagment(self):
        self.casePage.labelframe.pack()
        self.taskPage.labelframe.pack_forget()
        self.interPage.labelframe.pack_forget()
        self.aboutPage.pack_forget()

    def TaskManagement(self):
        self.casePage.labelframe.pack_forget()
        self.taskPage.labelframe.pack()
        self.interPage.labelframe.pack_forget()
        self.aboutPage.pack_forget()

    def Interface_Managment(self):
        self.casePage.labelframe.pack_forget()
        self.taskPage.labelframe.pack_forget()
        self.interPage.labelframe.pack()
        self.aboutPage.pack_forget()

    def aboutDisp(self):
        self.casePage.labelframe.pack_forget()
        self.taskPage.labelframe.pack_forget()
        self.interPage.labelframe.pack_forget()
        self.aboutPage.pack()