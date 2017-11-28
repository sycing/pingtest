from tkinter import *
from tkinter.messagebox import *
from tkinter import ttk

class AboutFrame(Frame): # 继承Frame类
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.root = master #定义内部变量root
        self.createPage()

    def createPage(self):
        Label(self, text='关于界面').pack()