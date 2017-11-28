import glob
from Control.read_path import *
from View.TaskFrame import *

class FindFile(object):
    def __init__(self,TaskNo):
        PATH = os.getcwd()
        self.superior_path = PATH + '\Report\*.xlsx'
        self.files = ""
        self.files1=[]
        self.Tlist0=[]
        self.Tlist1=[]
        self.Listnewfile=[]
        self.List0=[]
        self.List1=[]
        self.TaskNo =TaskNo
        self.Maxindex=''

    def MaxTime(self):

        self.files = glob.glob(self.superior_path)

        for item in self.files:
            path, file = os.path.split(item)
            self.Listnewfile.append(file)

        for i in range(len(self.Listnewfile)):
            Tpath=self.Listnewfile[i].split('.')
            self.List0.append(Tpath[0])        #Task0009_20170425133943

        for i in range(len(self.List0)):
            Tnum=self.List0[i].split('_')
            self.List1.append(Tnum[0])      #Task0009
            self.Tlist0.append(Tnum[1])  #所有报告文件尾部时间字符串的列表

        #self.TaskNo='Task0009'
        if self.TaskNo in self.List1:
            for i in range(0,len(self.List1)):
                if self.List1[i] == self.TaskNo:
                    self.files1.append(self.files[i])
                    self.Tlist1.append(self.Tlist0[i]) #查询到的同一任务的所有报告的列表
                    self.Maxindex = self.Tlist1.index(max(self.Tlist1))
                else:
                    continue
        try:
            os.system(self.files1[self.Maxindex])
        except:
            showerror(title="错误",message='报告不存在')
