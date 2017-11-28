from tkinter import *
from tkinter import ttk
from tkinter.messagebox import *  
from View.MainPage import *
  
class LoginPage(object):

    def __init__(self, master=None):  
        self.root = master #定义内部变量root
        #w, h = self.root.maxsize()
        #self.root.geometry("{}x{}".format(w, h))
        self.root.geometry('850x610') #设置窗口大小
        #self.root.resizable(False, False)    # 固定窗口大小，最大化按钮置灰

        # self.root.attributes("-alpha",0.6)    # 使窗口透明，透明度从0-1，1是不透明，0是全透明
        # self.root.attributes("-toolwindow", 1)  # 工具样式窗口，左上角的图标消失
        self.read_file = read_file()
        self.username = StringVar()  
        self.password = StringVar()
        self.signup_username = StringVar()
        self.signup_password = StringVar()
        self.signup_repassword = StringVar()
        self.case_header_list = []
        self.createPage()

    def createPage(self):

        self.page = ttk.Frame(self.root) #创建Frame
        self.page.pack()

        ttk.Label(self.page).grid(row=0, stick=W)
        ttk.Label(self.page).grid(row=1, stick=W)
        ttk.Label(self.page).grid(row=2, stick=W)
        ttk.Label(self.page).grid(row=3, stick=W)
        ttk.Label(self.page).grid(row=4, stick=W)
        ttk.Label(self.page, text='欢迎使用PingTest！', font=("Times", 22, "bold"), foreground='#4682B4').grid(row=5,columnspan=3, stick=E, rowspan=5)
        ttk.Label(self.page).grid(row=10, stick=W)
        ttk.Label(self.page).grid(row=11, stick=W)
        ttk.Label(self.page).grid(row=12, stick=W)
        # ttk.Label(self.page).grid(row=13, stick=W)
        ttk.Label(self.page, text = '账户: ', font=("Times", 12)).grid(row=14, stick=W, pady=10)
        # self.user = ttk.Entry(self.page, textvariable=self.username,validate='focusout',validatecommand=self.validate_user)
        self.user = ttk.Entry(self.page, textvariable=self.username, width=25)
        self.user.grid(row=14, column=1, stick=W, pady=10, columnspan=2)
        self.user.focus()
        self.user.bind("<Return>",self.loginCheck)

        ttk.Label(self.page, text = '密码: ', font=("Times", 12)).grid(row=15, stick=W, pady=10)
        # self.pw = ttk.Entry(self.page, textvariable=self.password, show='*',validate='focusout',validatecommand=self.validate_pw)
        self.pw = ttk.Entry(self.page, textvariable=self.password, show='*', width=25)
        self.pw.grid(row=15, column=1, stick=W, columnspan=2)
        self.pw.bind("<Return>",self.loginCheck)
        ttk.Label(self.page).grid(row=16, stick=W)
        ttk.Label(self.page).grid(row=17, stick=W)
        # ttk.Label(self.page).grid(row=18, stick=W)
        login = ttk.Button(self.page, text='登录', command=self.loginCheck, )
        login.grid(row=19, pady=20,stick=E)
        login.bind("<Return>", self.loginCheck)
        # ttk.Button(self.page, text='退出', command=self.page.quit, ).grid(row=19,  column=1,pady=20,stick=E)
        ttk.Button(self.page, text='注册', command=self.signup).grid(row=19, column=2, pady=20, stick=E)

    def loginCheck(self, event=None):
        name = self.username.get()  
        secret = self.password.get()

        # 读取EXCEL用例数据到列表中
        user_name = self.read_file.read_accounts_file().col_values(0)[1:]
        pwd = self.read_file.read_accounts_file().col_values(1)[1:]
        if name in user_name and pwd[user_name.index(name)] == secret:
            self.page.destroy()
            MainPage(self.root)
        else:  
            showinfo(title='错误', message='账号或密码错误！')
            self.user.delete(0,END)
            self.pw.delete(0,END)

    def signup(self):
        self.page.destroy()
        self.page = ttk.Frame(self.root)  # 创建Frame
        self.page.pack()
        ttk.Label(self.page).grid(row=0, stick=W)
        ttk.Label(self.page, text='账户: ').grid(row=1, stick=W, pady=10)
        # self.signup_user = ttk.Entry(self.page, textvariable=self.signup_username, validate='focusout',validatecommand=self.validate_user)
        self.signup_user = ttk.Entry(self.page, textvariable=self.signup_username)
        self.signup_user.grid(row=1, column=1, stick=W, pady=10, columnspan=2)
        self.signup_user.focus()

        ttk.Label(self.page, text='密码: ').grid(row=2, stick=W, pady=10)
        self.signup_pw = ttk.Entry(self.page, textvariable=self.signup_password, show='*')
        self.signup_pw.grid(row=2, column=1, stick=W, columnspan=2)

        ttk.Label(self.page, text='确认密码: ').grid(row=3, stick=W, pady=10)
        self.signup_repw = ttk.Entry(self.page, textvariable=self.signup_repassword, show='*')
        self.signup_repw.grid(row=3, column=1, stick=W, columnspan=2)

        ttk.Button(self.page, text='确定', command=self.signupCheck, ).grid(row=4, pady=20, stick=E)
        ttk.Button(self.page, text='取消', command=self.cancel).grid(row=4, column=1, pady=20, stick=E)
        self.root.bind("<Return>", self.signupCheck)


    def signupCheck(self ,event=None):
        name = self.signup_username.get().strip()
        pw = self.signup_password.get().strip()
        repw = self.signup_repw.get().strip()
        w_xls, sheet_write, rows = self.read_file.write_accounts_file()
        user_name = self.read_file.read_accounts_file().col_values(0)[1:]
        if len(name) > 20 or len(name) < 1:
            showinfo(title='错误', message='账号输入长度限制在1~20个字符！')
            self.signup_user.delete(0,END)
            return
        elif name in user_name:
            showinfo(title='错误', message='账号已存在！')
            self.signup_user.delete(0, END)
            return
        elif len(pw) > 20 or len(pw) < 1:
            showinfo(title='错误', message='密码输入长度限制在1~20个字符！')
            self.signup_pw.delete(0,END)
            return False
        # elif name in pw:
        #     showinfo(title='错误', message='密码中不能包含用户名！')
        #     self.signup_pw.delete(0, END)
        #     return
        elif repw == '':
            showinfo(title='错误', message='确认密码不能为空！')
            return
        elif pw != repw:
            showinfo(title='错误', message='两次密码输入不一致，请重新输入！')
            self.signup_repw.delete(0, END)
            return
            # 逐列写入数据
        self.case_header_list = [name, pw]
        try:
            for i in range(0, len(self.case_header_list)):
                sheet_write.write(rows, i, self.case_header_list[i])
            w_xls.save(self.read_file.save_accounts_file)
            showinfo(title="信息", message="注册成功！")
            self.cancel()
        except:
            showerror(title='错误', message="注册失败！")

    def cancel(self):
        self.signup_user.delete(0, END)
        self.signup_pw.delete(0, END)
        self.signup_repw.delete(0, END)
        self.page.destroy()
        self.createPage()

    # def validate_user(self):
    #     name = self.signup_user.get()
    #     # 读取EXCEL用例数据到列表中
    #     user_name = self.read_file.read_accounts_file().col_values(0)[1:]
    #     if len(name) > 20:
    #         showinfo(title='错误', message='账号输入长度过长，限制为20个字符！')
    #         self.signup_user.delete(0,END)
    #         return False
    #     elif name.strip() in user_name:
    #         showinfo(title='错误', message='账号已存在！')
    #         self.signup_user.delete(0,END)
    #         return False
    #     else:
    #         return True
    #
    # def validate_pw(self):
    #     pw = self.signup_pw.get()
    #     name = self.signup_user.get()
    #     if len(pw) > 20:
    #         showinfo(title='错误', message='密码输入长度过长，限制为20个字符！')
    #         self.signup_pw.delete(0,END)
    #         return False
    #     elif name.strip() in pw.strip() and pw.strip() != '':
    #         showinfo(title='错误', message='密码中不能包含用户名！')
    #         self.signup_pw.delete(0, END)
    #         return False
    #     else:
    #         return True
    #
    # def validate_repw(self):
    #     pw = self.signup_pw.get()
    #     repw = self.signup_repw.get()
    #     print(pw)
    #     print(repw.strip())
    #     if pw.strip() != repw.strip() and repw.strip() != '' and pw.strip() != '':
    #         showinfo(title='错误', message='两次密码输入不一致，请重新输入！')
    #         self.signup_repw.delete(0, END)
    #         return False
    #     else:
    #         return True
