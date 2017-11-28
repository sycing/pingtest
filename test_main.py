#!/usr/bin/env python
#coding=utf-8
from tkinter import *
from View.LoginPage import *
class test_main():
    def __init__(self):
        self.root = Tk()

    def main(self):

        self.root.title('自动化测试平台pingtest')
        LoginPage(self.root)
        self.root.mainloop()

if __name__ =='__main__':
    PingTest = test_main()
    PingTest.main()