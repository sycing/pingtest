# coding:utf-8
import pymysql.cursors
from log_config.read_logconfig import loggers
from View.TaskFrame import *
import configparser
from Control.read_path import read_path
import time
import datetime
import decimal

# ======== MySql base operating ===================
class DB():
    def __init__(self,parent):
        self.parent = parent
        #读取配置文件
        '''conf_path = read_path('\pingtest\Control\DataBase_conf.ini')
        config = configparser.ConfigParser()
        # 从配置文件中读取数据
        config.read(conf_path.read_superior_path(), encoding='utf-8')
        self.host = config['DataBase_Connection']['host']
        self.user = config['DataBase_Connection']['user']
        self.password = config['DataBase_Connection']['password']'''

        try:
            # Connect to the database
            self.connection = pymysql.Connect(host='192.168.50.215',
                                              user='kairuTest',
                                              password='l4atnjTA',
                                              charset='utf8',
                                              cursorclass=pymysql.cursors.DictCursor)
        except pymysql.err.OperationalError as e:
            loggers.error("Mysql Error %d: %s" % (e.args[0], e.args[1]))

    # close database
    def close(self):
        self.connection.close()

    def dict_datetime(self,datetime_json):  # 转换QB结果中类型为datetime类型的字段，转化为YYYY-MM-DD HH:MM:SS格式
        if isinstance(datetime_json,dict):
            for key,value in datetime_json.items():
                if isinstance(value,datetime.datetime):
                    datetime_json[key] = datetime_json[key].strftime("%Y-%m-%d %H:%M:%S")
                    continue
                if isinstance(value,datetime.date):
                    datetime_json[key] = datetime_json[key].strftime("%Y-%m-%d")
                if isinstance(value,decimal.Decimal):
                    datetime_json[key] = float(datetime_json[key])
                if isinstance(value,dict):
                    self.dict_datetime(value)
                if isinstance(value,list):
                    for i in range(0,len(value)):
                        if isinstance(value[i],dict):
                            self.dict_datetime(value[i])
                        else:
                            value[i] = value[i].strftime("%Y-%m-%d %H:%M:%S")
        elif isinstance(datetime_json,list):
            for i in range(0,len(datetime_json)):
                self.dict_datetime(datetime_json[i])
        return datetime_json

    def search(self, sql,Expect_Type):
        real_sql = sql.replace("\n"," ")
        result = {}
        self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "期望结果为SQL语句：\n%s\n" %real_sql,('b'))
        self.parent.show_log.update()
        self.parent.show_log.see("end")
        loggers.info("期望结果为SQL语句：\n%s\n" %real_sql)
        with self.connection.cursor() as cursor:
            cursor.execute(real_sql)
        self.connection.commit()
        if Expect_Type == "single":
            result = cursor.fetchone()
            if result is None:
                result = {}
        elif Expect_Type == "multiple":
            result = cursor.fetchall()
            if result == ():
                result = list(result)
        else:
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "测试用例期望类型填写有误，请选择填写json或SQL\n",('b'))
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.error("测试用例期望类型填写有误，请选择填写json或SQL")

        result = self.dict_datetime(result)
        #loggers.info("mysql查询结果：%s" %result)
        #self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "mysql查询结果：%s\n" %result,('b'))
        #self.parent.show_log.update()
        #self.parent.show_log.see("end")
        return result

    def run(self, sql):
        result = []
        try:
            real_sql = eval(sql.replace("\n"," "))
            self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "执行的SQL语句：\n%s\n" %real_sql,('b'))
            self.parent.show_log.update()
            self.parent.show_log.see("end")
            loggers.info("执行的SQL语句：\n%s\n" %real_sql)
            for i in range(0,len(real_sql)):
                with self.connection.cursor() as cursor:
                    try:
                        cursor.execute(real_sql[i])
                    except BaseException as e:
                        self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "执行的SQL语句错误：\n%s\n" %e,('a'))
                        self.parent.show_log.update()
                        self.parent.show_log.see("end")
                        return e

                self.connection.commit()
                f = cursor.fetchall()
                if f == ():
                    f = dict(f)
                if len(f) == 1:
                    f = f[0]
                result.append(f)
        except:
            real_sql = sql.replace("\n"," ")
            with self.connection.cursor() as cursor:
                try:
                    cursor.execute(real_sql)
                except BaseException as e:
                    self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "执行的SQL语句错误：\n%s\n" %e,('a'))
                    self.parent.show_log.update()
                    self.parent.show_log.see("end")
                    return e
            self.connection.commit()
            f = cursor.fetchall()
            if f == ():
                f = dict(f)
            if len(f) == 1:
                f = f[0]
            result.append(f)

        if len(result) == 1:
            result = result[0]

        result = self.dict_datetime(result)
        #loggers.info("mysql执行结果：%s" %result)
        #self.parent.show_log.insert(END,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ":" + "mysql执行结果：%s\n" %result,('b'))
        #self.parent.show_log.update()
        #self.parent.show_log.see("end")
        return result
