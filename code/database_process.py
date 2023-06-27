#数据库处理
import sqlite3
from sqlite3 import Error
import os

class db_process():
    def __init__(self):
        self.db="./cost_compare.sqlite3"#数据库名称

    def db_get(self):#获取数据库
        self.conn = None
        try:
            self.conn = sqlite3.connect(self.db)
        except Error as result:
            print(result)
        if self.conn is not None:
            return self.conn

    def db_close(self):#关闭数据库
        if self.cur is not None:
            self.cur.close()
        if self.conn is not None:
            self.conn.close()

    def db_load(self,tablename):#读数据库
        sql = 'select * from '+ tablename #表格名称str
        self.cur = self.conn.cursor()
        self.cur.execute(sql)
        self.data_loaded = self.cur.fetchall()
        self.db_close()
        
    def db_load_selected(self,sql):#自定义读数据库
        self.cur = self.conn.cursor()
        self.cur.execute(sql)
        self.data_loaded = self.cur.fetchall()
        self.db_close()

    def db_write(self,tablename,inputtype,inputdata):#新增数据
        inputdata_len = len(inputdata[1])
        temp = 'values(?'
        for i in range(inputdata_len-1):
            temp=temp+',?'
        temp=temp+')'
        self.cur = self.conn.cursor()
        sql = 'insert into ' + tablename +  ' ' + inputtype + temp
        self.cur.executemany(sql,inputdata)#多个数据
        self.conn.commit()
        self.db_close()
        
    def db_write_one(self,tablename,inputtype,inputdata):#新增单条数据
        inputdata_len = len(inputdata)
        temp = 'values(?'
        for i in range(inputdata_len-1):
            temp=temp+',?'
        temp=temp+')'
        self.cur = self.conn.cursor()
        sql = 'insert into ' + tablename +  ' ' + inputtype + temp
        self.cur.execute(sql,inputdata)#单个数据
        self.conn.commit()
        self.db_close()
        
    def db_del(self,tablename,deldata_name,deldata):#删数据
        self.cur = self.conn.cursor()
        sql = 'delete from ' + tablename +  ' where '+deldata_name+'=?'
        self.cur.execute(sql,deldata)
        self.conn.commit()
        self.db_close()

    def empty_table(self,tablename):#清空表数据
        self.cur = self.conn.cursor()
        sql = 'delete from ' + tablename
        self.cur.execute(sql)
        self.conn.commit()
        self.db_close()

    def db_upload(self,tablename,inputdata_tag,inputdata_add,inputdata):#改多个数据
        self.cur = self.conn.cursor()
        sql = 'update ' + tablename +  ' set '+inputdata_tag+' where '+inputdata_add
        self.cur.execute(sql,inputdata)
        self.conn.commit()
        self.db_close()
        
    def db_upload_list(self,tablename,inputdata_tag,inputdata_add,inputdata,inputdata_add_tag):#改多行多个数据
        self.cur = self.conn.cursor()
        a = 0
        for i in inputdata:
            sql = 'update ' + tablename +  ' set '+inputdata_tag+' where '+inputdata_add+"="+str(inputdata_add_tag[a])
            self.cur.execute(sql,i)
            self.conn.commit()
            a += 1
        self.db_close()
        