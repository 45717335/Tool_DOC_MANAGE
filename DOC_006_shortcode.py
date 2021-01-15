from tkinter import*
from fnmatch import fnmatch
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
from tkinter.simpledialog import askstring, askinteger, askfloat
from tkinter import ttk
import os
import sqlite3 as sqlite
import pandas as pd
import numpy as np
import sys
import re
from pandas import DataFrame
from xlwt import *
import xlrd
import shutil
import datetime
import time
DB_FILE_PATH = ''
TABLE_NAME = ''
SHOW_SQL = True
DC1 = {
            'dbpath':'Z:\\31_PTS\\01_Projects\\ASY\\CN.305899_.BBAC_M254 Engine Assy Line\\04_Documentation\\STATION_DATA',
            'dbname':'python_doc.db',
            'docroot':'Z:\\31_PTS\\01_Projects\ASY\\CN.305899_.BBAC_M254 Engine Assy Line\\04_Documentation\\048_Documentatio_CN\\TKSECN Scope\\TOCUST',
            'pjn':'CN.XXXXXXXX',
            'info':'2021-01-15',
      }
DC2 ={
        'wdn':"",
        'tbn':"",
        'dbn':"",
        'xln':"",
        'addone':"",
        'delone':"",
      }
def get_conn(path):
    conn = sqlite.connect(path)
    if os.path.exists(path) and os.path.isfile(path):
        print('硬盘上面:[{}]'.format(path))
        return conn
    else:
        conn = None
        print('内存上面:[:memory:]')
        return sqlite.connect(':memory:')
def get_cursor(conn):
    '''该方法是获取数据库的游标对象，参数为数据库的连接对象
    如果数据库的连接对象不为None，则返回数据库连接对象所创
    建的游标对象；否则返回一个游标对象，该对象是内存中数据
    库连接对象所创建的游标对象'''
    if conn is not None:
        return conn.cursor()
    else:
        return get_conn('').cursor()
###############################################################
####            创建|删除表操作     START
###############################################################
def drop_table(conn, table):
    '''如果表存在,则删除表，如果表中存在数据的时候，使用该
    方法的时候要慎用！'''
    if table is not None and table != '':
        sql = 'DROP TABLE IF EXISTS ' + table
        if SHOW_SQL:
            print('执行sql:[{}]'.format(sql))
        cu = get_cursor(conn)
        cu.execute(sql)
        conn.commit()
        print('删除数据库表[{}]成功!'.format(table))
        close_all(conn, cu)
    else:
        print('the [{}] is empty or equal None!'.format(sql))
def create_table(conn, sql):
    '''创建数据库表：student'''
    if sql is not None and sql != '':
        cu = get_cursor(conn)
        if SHOW_SQL:
            print('执行sql:[{}]'.format(sql))
        cu.execute(sql)
        conn.commit()
        print('创建数据库表成功!')
        close_all(conn, cu)
    else:
        print('the [{}] is empty or equal None!'.format(sql))
###############################################################
####            创建|删除表操作     END
###############################################################
def close_all(conn, cu):
    '''关闭数据库游标对象和数据库连接对象'''
    try:
        if cu is not None:
            cu.close()
    finally:
        if cu is not None:
            cu.close()
###############################################################
####            数据库操作CRUD     START
###############################################################
def save(conn, sql, data):
    '''插入数据'''
    if sql is not None and sql != '':
        if data is not None:
            cu = get_cursor(conn)
            for d in data:
                if SHOW_SQL:
                    print('执行sql:[{}],参数:[{}]'.format(sql, d))
                #有时候会插入一些可能会重复的项目，这个时候会发生错误，那么跳过它就行
                try:
                    cu.execute(sql, d)
                except:
                    print("ERROR1")
                finally:
                    conn.commit()
            close_all(conn, cu)
    else:
        print('the [{}] is empty or equal None!'.format(sql))
def rec_exist(conn,sql):
    if sql is not None and sql != '':
        cu = get_cursor(conn)
        if SHOW_SQL:
            print('执行sql:[{}]'.format(sql))
        cu.execute(sql)
        r = cu.fetchall()
        if len(r) > 0:
            return True
        else:
            return False
    else:
        return False
        print('the [{}] is empty or equal None!'.format(sql))
def fetchall_list(conn,sql):
    '''查询所有数据'''
    datax = []
    if sql is not None and sql != '':
        cu = get_cursor(conn)
        if SHOW_SQL:
            print('执行sql:[{}]'.format(sql))
        cu.execute(sql)
        r = cu.fetchall()
        if len(r) > 0:
            for e in range(len(r)):
                datax.append(r[e])
    else:
        print('the [{}] is empty or equal None!'.format(sql))
    return datax
def fetchall_st(conn,sql,text):
    '''查询所有数据'''
    if sql is not None and sql != '':
        cu = get_cursor(conn)
        if SHOW_SQL:
            print('执行sql:[{}]'.format(sql))
        cu.execute(sql)
        r = cu.fetchall()
        if len(r) > 0:
            for e in range(len(r)):
                text.insert(END,r[e])
                text.insert(END,"\n")
                #print(r[e])
        text.insert(END,"records found:{0}\n".format(len(r)))
        text.see(END)
    else:
        print('the [{}] is empty or equal None!'.format(sql))
def fetchall(conn, sql):
    if sql is not None and sql != '':
        cu = get_cursor(conn)
        if SHOW_SQL:
            print('执行sql:[{}]'.format(sql))
        cu.execute(sql)
        r = cu.fetchall()
        if len(r) > 0:
            for e in range(len(r)):
                print(r[e])
    else:
        print('the [{}] is empty or equal None!'.format(sql))
def fetchone(conn, sql, data):
    '''查询一条数据'''
    if sql is not None and sql != '':
        if data is not None:
            #Do this instead
            d = (data,)
            cu = get_cursor(conn)
            if SHOW_SQL:
                print('执行sql:[{}],参数:[{}]'.format(sql, data))
            cu.execute(sql, d)
            r = cu.fetchall()
            if len(r) > 0:
                for e in range(len(r)):
                    print(r[e])
        else:
            print('the [{}] equal None!'.format(data))
    else:
        print('the [{}] is empty or equal None!'.format(sql))
def update(conn, sql, data):
    '''更新数据'''
    if sql is not None and sql != '':
        if data is not None:
            cu = get_cursor(conn)
            for d in data:
                d=tuple(d)
                if SHOW_SQL:
                    print('执行sql:[{}],参数:[{}]'.format(sql, d))
                cu.execute(sql, d)
                conn.commit()
            close_all(conn, cu)
    else:
        print('the [{}] is empty or equal None!'.format(sql))
def delete(conn, sql, data):
    '''删除数据'''
    if sql is not None and sql != '':
        if data is not None:
            cu = get_cursor(conn)
            for d in data:
                if SHOW_SQL:
                    print('执行sql:[{}],参数:[{}]'.format(sql, d))
                cu.execute(sql, d)
                conn.commit()
            close_all(conn, cu)
    else:
        print('the [{}] is empty or equal None!'.format(sql))
###############################################################
####            数据库操作CRUD     END
###############################################################
def mytab_exist(conn,table):
    '''检查表单是否存在 返回 True，False'''
    sql = "SELECT name FROM sqlite_master WHERE type='table' AND name='{0}';".format(table)
    if SHOW_SQL:
        print('执行sql:[{}]'.format(sql))
    cu = get_cursor(conn)
    cu.execute(sql)
    r = cu.fetchall()
    if len(r) > 0:
        print("Table exist:" + table )
        return True
    else:
        print("Table Not exist:" + table )
        return False
###############################################################
####            XLS TO DB,DB TO XLS     START
###############################################################
def dbtoxls(dbName,excelName,tableName):
    #指定文件名称表名称
    xlspath = excelName
    dbpath = dbName
    print("<%s> --> <%s>" % (dbpath, xlspath))
    db = sqlite.connect(dbpath)
    cur = db.cursor()
    w = Workbook()
    for tbl_name in [row[0] for row in query_by_sql(cur, "select tbl_name FROM sqlite_master where type = 'table'")]:
        if tbl_name==tableName:
            select_sql = "select * from '%s'" % tbl_name
            sqlite_to_workbook_with_head(cur, tbl_name, select_sql, w)
    cur.close()
    db.close()
    w.save(xlspath)
def query_by_sql(cur, select_sql):
    cur.execute(select_sql)
    return cur.fetchall()
def sqlite_to_workbook_with_head(cur, table, select_sql, workbook):
    ws = workbook.add_sheet(table)
    print('create table %s.' % table)
    #enumerate针对一个可迭代对象，生成的是序号加上内容
    for colx, heading in enumerate(sqlite_get_col_names(cur, select_sql)):
        ws.write(0, colx, heading)    #在第1行的colx列写上头部信息
    for rowy, row in enumerate(query_by_sql(cur, select_sql)):
        for colx, text in enumerate(row):    #row是一行的内容
            ws.write(rowy + 1, colx, text)    #在rowy+1行，colx写入数据库内容text
def sqlite_get_col_names(cur, select_sql):
    cur.execute(select_sql)
    return [tuple[0] for tuple in cur.description]
def runcate_del(conn,table):
    sql="DELETE FROM "+table
    if sql is not None and sql != '':
        cu = get_cursor(conn)
        if SHOW_SQL:
            print('执行sql:[{}]'.format(sql))
        cu.execute(sql)
        conn.commit()
        close_all(conn, cu)
    else:
        print('the [{}] is empty or equal None!'.format(sql))
class ExcelToSqlite(object):
    exe = "     执行: "
    output = "     输出: "
    sheetDataStartIndex = 1  # 数据开始计算的行数，如第0行是表头，第1行及之后是数据
    def __init__(self, dbName):
        print("初始化数据库实例")
        super(ExcelToSqlite, self).__init__()
        self.conn = sqlite.connect(dbName)
        self.cursor = self.conn.cursor()
    def __del__(self):
        print("释放数据库实例")
        self.cursor.close()
        self.conn.close()
    def ExcelToDb(self, excelName, sheetIndex, tableName):
        """
        excel转化为sqlite数据库表
        :param excelName:excel名
        :param sheetIndex:excel中sheet位置
        :param tableName:数据库表名
        """
        print("Excel文件 转 db")
        self.tableName = tableName
        excel = xlrd.open_workbook(excelName)
        sheet = excel.sheets()[sheetIndex]  # sheets 索引
        self.sheetRows = sheet.nrows  # excel 行数
        self.sheetCols = sheet.ncols  # excle 列数
        fieldNames = sheet.row_values(0)  # 得到表头字段名
        # 创建表
        fieldTypes = ""
        for index in range(fieldNames.__len__()):
            if (index != fieldNames.__len__() - 1):
                fieldTypes += fieldNames[index] + " text,"
            else:
                fieldTypes += fieldNames[index] + " text"
        self.__CreateTable(tableName, fieldTypes)
        # 插入数据
        for rowId in range(self.sheetDataStartIndex, self.sheetRows):
            fieldValues = sheet.row_values(rowId)
            self.__Insert(fieldNames, fieldValues)
    def __CreateTable(self, tableName, field):
        """
        创建表
        :param tableName: 表名
        :param field: 字段名及类型
        :return:
        """
        print("创建表 " + tableName)
        sql = 'create table if not exists %s(%s)' % (self.tableName, field)  # primary key not null
        print(self.exe + sql)
        self.cursor.execute(sql)
        self.conn.commit()
    def __Insert(self, fieldNames, fieldValues):
        """
        插入数据
        :param fieldNames: 字段list
        :param fieldValues: 值list
        """
        # 通过fieldNames解析出字段名
        names = ""  # 字段名，用于插入数据
        nameTypes = ""  # 字段名及字段类型，用于创建表
        for index in range(fieldNames.__len__()):
            if (index != fieldNames.__len__() - 1):
                names += fieldNames[index] + ","
                nameTypes += fieldNames[index] + " text,"
            else:
                names += fieldNames[index]
                nameTypes += fieldNames[index] + " text"
        # 通过fieldValues解析出字段对应的值
        values = ""
        for index in range(fieldValues.__len__()):
            cell_value = str((fieldValues[index]))
            if (isinstance(fieldValues[index], float)):
                cell_value = str((int)(fieldValues[index]))  # 读取的excel数据会自动变为浮点型，这里转化为文本
            if (index != fieldValues.__len__() - 1):
                values += "\'" + cell_value + "\',"
            else:
                values += "\'" + cell_value + "\'"
        # 将fieldValues解析出的值插入数据库
        sql = 'insert into %s(%s) values(%s)' % (self.tableName, names, values)
        print(self.exe + sql)
        self.cursor.execute(sql)
        self.conn.commit()
    def Query(self, tableName):
        """
        查询数据库表中的数据
        :param tableName:表名
        """
        print("查询表 " + tableName)
        sql = 'select * from %s' % (tableName)
        print(self.exe + sql)
        self.cursor.execute(sql)
        results = self.cursor.fetchall()  # 获取所有记录列表
        index = 0
        for row in results:
            print(self.output + "index=" + index.__str__() + " detail=" + str(row))  # 打印结果
            index += 1
        print(self.output + "共计" + results.__len__().__str__() + "条数据")
    def executeSqlCommand(self, sqlCommand):
        """
        执行输入的sql命令
        :param sqlCommand: sql命令
        """
        print("执行自定义sql " + tableName)
        print(self.exe + sqlCommand)
        self.cursor.execute(sqlCommand)
        results = self.cursor.fetchall()
        print(self.output + str(results))
        for index in range(0, results.__len__()):
            print(self.output + str(results[index]))
        self.conn.commit()
def xlstodb( excelName,dbName,tableName):
    es = ExcelToSqlite(dbName)
    es.ExcelToDb(excelName, 0, tableName)
###############################################################
####            XLS TO DB,DB TO XLS     END
###############################################################
def myinit():
    DC1['dbpath']=os.path.dirname(os.path.realpath(sys.executable))
    print (DC1['dbpath'])
    global DB_FILE_PATH
    DB_FILE_PATH =os.path.join(DC1['dbpath'],DC1['dbname'])
    #数据库表名称
    global TABLE_NAME
    TABLE_NAME = 'student'
    #是否打印sql
    global SHOW_SQL
    SHOW_SQL = True
    print('show_sql : {}'.format(SHOW_SQL))
    conn = get_conn(DB_FILE_PATH)
    if mytab_exist(conn, 'doc') == False:
        create_table_sql = '''CREATE TABLE `doc` (
                                  `id_flntime` varchar(100) NOT NULL,
                                  `tkid_custid` varchar(100) NOT NULL,
                                  `tkid_stn` varchar(50) DEFAULT NULL,
                                  `custid_stn` varchar(50) DEFAULT NULL,
                                  `id_doctype` varchar(50) DEFAULT NULL,
                                  `status` varchar(50) DEFAULT NULL,
                                  `file_fullpath` varchar(200) DEFAULT NULL,
                                  `to_fullpath` varchar(200) DEFAULT NULL,
                                  `fldate` varchar(50) DEFAULT NULL,
                                  `docdate` varchar(50) DEFAULT NULL,
                                  `temp_fullpath` varchar(200) DEFAULT NULL,
                                   PRIMARY KEY (`to_fullpath`)
                                )'''
        create_table(conn, create_table_sql)
    if mytab_exist(conn, 'station') == False:
        create_table_sql = '''CREATE TABLE `station` (
                                  `tkid_custid` varchar(100) NOT NULL,
                                  `tkid_stn` varchar(50) DEFAULT NULL,
                                  `custid_stn` varchar(50) DEFAULT NULL,
                                  `status` varchar(50) DEFAULT NULL,
                                   PRIMARY KEY (`tkid_custid`)
                                )'''
        create_table(conn, create_table_sql)
    if mytab_exist(conn, 'doc_type') == False:
        create_table_sql = '''CREATE TABLE `doc_type` (
                                  `id_doctype` varchar(50) NOT NULL,
                                  `folder` varchar(100) DEFAULT NULL,
                                  `ower` varchar(50) NOT NULL,
                                  `desc` varchar(200) NOT NULL,
                                   PRIMARY KEY (`id_doctype`)
                                )'''
        create_table(conn, create_table_sql)
    if mytab_exist(conn, 'milestone') == False:
        create_table_sql = '''CREATE TABLE `milestone` (
                                  `id_date` varchar(50) NOT NULL,
                                  `desc1` varchar(200) DEFAULT NULL,
                                  `desc2` varchar(200) NOT NULL,
                                  `desc3` varchar(200) NOT NULL,
                                   PRIMARY KEY (`id_date`)
                                )'''
        create_table(conn, create_table_sql)

    if mytab_exist(conn, 'tobe_doc') == False:
        create_table_sql = '''CREATE TABLE `tobe_doc` (
                                     `id_flntime` varchar(100) NOT NULL,
                                     `tkid_custid` varchar(100) NOT NULL,
                                     `tkid_stn` varchar(50) DEFAULT NULL,
                                     `custid_stn` varchar(50) DEFAULT NULL,
                                     `id_doctype` varchar(50) DEFAULT NULL,
                                     `status` varchar(50) DEFAULT NULL,
                                     `file_fullpath` varchar(200) DEFAULT NULL,
                                     `to_fullpath` varchar(200) DEFAULT NULL,
                                     `fldate` varchar(50) DEFAULT NULL,
                                     `temp_fullpath` varchar(200) DEFAULT NULL,
                                      PRIMARY KEY (`id_flntime`)
                                   )'''
        create_table(conn, create_table_sql)
    if mytab_exist(conn, 'dbinit') == False:
        create_table_sql = '''CREATE TABLE `dbinit` (
                                     `init_key` varchar(100) NOT NULL,
                                     `init_val` varchar(2000) NOT NULL,
                                      PRIMARY KEY (`init_key`)
                                   )'''
        create_table(conn, create_table_sql)
    conn = get_conn(DB_FILE_PATH)
    cu = get_cursor(conn)
    li1= fetchall_list(conn,"SELECT init_val from dbinit where init_key='docroot'")
    if len(li1)==0:
        x1=""
        while os.path.exists(x1)==False:
            x1 = askstring("init", "Please input the root path for the project document")
        cu.execute('''INSERT INTO dbinit values (?, ?)''',["docroot",x1])
        conn.commit()
        DC1["docroot"] = x1
    else:
        DC1["docroot"] = str(li1[0][0])
    li1 = fetchall_list(conn, "SELECT init_val from dbinit where init_key='pjn'")
    if len(li1) == 0:
        x1 = askstring("init", "Please input the project Number")
        cu.execute('''INSERT INTO dbinit values (?, ?)''', ["pjn", x1])
        conn.commit()
        DC1["pjn"] = x1
    else:
        DC1["pjn"] =str(li1[0][0])

    li1 = fetchall_list(conn, "SELECT init_val from dbinit where init_key='info'")
    if len(li1) == 0:
        x1 = '''
2021-01-15

STN...      Station information, adding and deleting stations

ADD_DOC...  Add the documents that need submite to the customer to database

DOC...      Displays documents that have been added to the database

DOC_TYP...  Define the type and their store path of the documents

INI...      Initialization infomation , such as project number, ect.

MST...      Milestone, record the major events in the project documentation work
        '''
        cu.execute('''INSERT INTO dbinit values (?, ?)''', ["info", x1])
        conn.commit()
        DC1["info"] = x1
    else:
        DC1["info"] = str(li1[0][0])
s1=[["Manuf",""],["ManufRef",""],["tkid",""]]
def bt3():
    b1=False    # 必须先导出，然后才能导入，当 btn3_btn4 被执行后，本值被设定为True，bt3_btn5 被执行后再次设定为false
    def bt3_bt4():
        text3.delete(1.0,END)
        text3.insert(END,"EXPORT STATION:")
        text3.insert(END, "\n")
        text3.insert(END, os.path.join(DC1['dbpath'], "station.xls"))
        text3.insert(END, "\n")
        #DC1['dbpath']
        dbtoxls(DB_FILE_PATH, os.path.join(DC1['dbpath'],"station.xls"), "station")
        os.startfile(os.path.join(DC1['dbpath'],"station.xls"))
        nonlocal b1
        b1=True
        sql="SELECT * FROM station where status='OK'"
        conn = get_conn(DB_FILE_PATH)
        fetchall_st(conn, sql, text3)
    def bt3_bt5():#xlstodb
        if b1==False:
            text3.insert(END, "CAN NOT IMPORT WITHOUT EXPORT:")
            text3.insert(END, "\n")
            return
        text3.insert(END,"IMPORT STATION FROM:")
        text3.insert(END, "\n")
        text3.insert(END,os.path.join(DC1['dbpath'],"station.xls"))
        text3.insert(END, "\n")
        if os.path.exists( os.path.join(DC1['dbpath'],"station.xls") )==False:
            return
        conn = get_conn(DB_FILE_PATH)
        runcate_del(conn, "station")
        xlstodb(os.path.join(DC1['dbpath'], "station.xls"), DB_FILE_PATH, "station")
        sql="SELECT * FROM station where status='OK'"
        conn = get_conn(DB_FILE_PATH)
        fetchall_st(conn, sql, text3)
    def bt3_bt6():
        fdf=DC1['docroot']
        data=[]
        for root, dirs, files in os.walk(fdf):
            for dir1 in dirs:
                print('folder:' + os.path.join(root, dir1))
                data.append([dir1,"","","TOBECHECK"])
            break
        save_sql = '''INSERT INTO station values (?,?,?,?)'''
        conn = get_conn(DB_FILE_PATH)
        save(conn, save_sql, data)
    w_bt3 = Toplevel(myWindow)      # 按钮3 创建的 窗口，用于导入导出 STATION表
    w_bt3.title('Exp_Imp_Station')
    frame=Frame(w_bt3)              # 按钮3创建的表还有一排按钮，bt3_bt1 bt3_bt1 bt3_bt1 bt3_bt1
    b4 = Button(frame, command=bt3_bt4, text='EXP_XLS', font=('Helvetica 10 bold'), width=8, height=2)
    b4.grid(row=0, column=3, sticky=W, padx=5, pady=5)
    b5 = Button(frame, command=bt3_bt5, text='IMP_XLS', font=('Helvetica 10 bold'), width=8, height=2)
    b5.grid(row=0, column=4, sticky=W, padx=5, pady=5)
    b6 = Button(frame, command=bt3_bt6, text='UPDATE', font=('Helvetica 10 bold'), width=8, height=2)
    b6.grid(row=0, column=5, sticky=W, padx=5, pady=5)
    frame.pack()
    text3 = ScrolledText(w_bt3)
    text3.pack(expand=YES, fill=BOTH)
def bt4():
    def bt4_bt1():
        bt4_bt1_s1 = text4.get(1.0,END)
        text4.insert(END,"*"*50+"\n")
        text4.insert(END, bt4_bt1_s1)
        text4.insert(END, "*" * 50 + "\n")
        fetchone_sql=bt4_bt1_s1
        conn = get_conn(DB_FILE_PATH)
        fetchall_st(conn,fetchone_sql,text4)
    w_bt4 = Toplevel(myWindow)
    w_bt4.title('Input sql')
    w_bt4.geometry('300x200')
    frame=Frame(w_bt4)
    b1 = Button(frame, command=bt4_bt1, text='Run SQL', relief='raised', font=('Helvetica 10 bold'),width=8, height=2)
    b1.grid(row=0, column=0, sticky=W, padx=5, pady=5)
    frame.pack()
    text4 = ScrolledText(w_bt4)
    text4.pack(expand=YES, fill=BOTH)
lam1 = lambda x: "\n" + "*" * 50 + "\n" + x + "\n" + "*" * 50 + "\n"
def my_cli(lt1):
    li1 = []
    for x in lt1:
        li1.append(x[0])
    return li1
def bt5():#管理 tobe_doc表格
    b1=False
    def find_in(s1, li1):
        for x in li1:
            if len(x) > 1:
                if x.rfind(s1) >= 0:
                    return x
        return ""
    def find_match(s1, li1):
        for x in li1:
            if len(x) > 1:
                if s1.rfind(x) > 0:
                    return x
        return ""
    def remo(l1):
        '去除特殊字符'
        ll1 = []
        for x in l1:
            # ll1.append( x.rplace("(","").rplace(")",""))
            x1 = x
            if len(x1) > 2:
                ll1.append(x1.replace("(", "").replace(")", ""))
        return ll1
    def bt5_bt1():
        '''
        输入一个通配符 寻找指定文件夹下面的 文件，将其填入 tobe_doc 表格，
        1.清空tobe_doc
        2.要求输入文件全路径和，通配符
        3.遍历所有满足的，填写到 数据表中，
        4.补全其余信息
        5.填写到数据库。
        接着调用 导出 EXCEL，
        另外一个按钮是 excel导入到数据库，然后 检查是否合法，并复制
        :return:
        '''
        nonlocal fdpath,fltpf,text5,doctypeChosen
        if doctypeChosen.get()=="":
            text5.insert(END,"EMPTY DOC TYPE PLEASE SELECT")
            return
        myWindow.withdraw()
        fdpath.set( filedialog.askdirectory() )
        fltpf.set(askstring("Input", "TongPeiFu of the filename(contain * / ? )"))
        myWindow.deiconify()
        conn = get_conn(DB_FILE_PATH)
        runcate_del(conn,'tobe_doc')
        sql = "select tkid_stn from station"
        x1 = my_cli(fetchall_list(conn, sql))
        sql = "select custid_stn from station"
        x2 = my_cli(fetchall_list(conn, sql))
        sql = "select tkid_custid from station"
        x3 =  my_cli(fetchall_list(conn, sql))
        xx2=remo(x2)
        x6=doctypeChosen.get()
        sql = "select folder from doc_type where id_doctype='{}'".format(x6)
        xx6=my_cli(fetchall_list(conn, sql))[0]
        xx=[]
        for root, dirs, files in os.walk(fdpath.get()):
            for fl1 in files:
                if fnmatch(fl1,fltpf.get()):
                    #获取文件的最后修改时间?
                    x4=find_match(fl1,x1)
                    x5=find_match(fl1,xx2)
                    x8 = ""
                    x7 = ""
                    x9 = "TOBECHECK"
                    x10=""
                    t1=os.path.getmtime(os.path.join(root,fl1))
                    x11= datetime.datetime.fromtimestamp(t1).strftime("%Y-%m-%d %H%M%S")
                    x12="" #防止复制出问题，镜像一个TOCUST
                    #x6=find_match()
                    if len(x4)>0 and len(x5)>0:
                        x7=find_in(x5,x2)
                        if len(x7)>0:
                            if x3.count(x4+x7)>0:
                                x8=x4+x7
                            else:
                                x8=""
                            if len(x8)>0:
                                x9="OK"
                    #如果 不是 tkid custid 同时匹配的话，tkid匹配也行，这个时候，status肯定还是tobecheck 但是，文件夹和目标路径可以填写
                    if x8=="":
                        if len(x4)>0:
                            x8=find_in(x4,x3)
                    if len(x8)>0:
                        x10=os.path.join(DC1['docroot'],x8,xx6,fl1)
                    x12=os.path.join(DC1['dbpath'],"TOCUST",x8,xx6,fl1)
                    xx.append([fl1,x8,x4,x5,x6,x9,os.path.join(root,fl1),x10,x11,x12]) # fln_time,
        save_sql='''INSERT INTO tobe_doc values (?,?,?,?,?,?,?,?,?,?)'''
        conn = get_conn(DB_FILE_PATH)
        save(conn, save_sql,xx)
        sql ="SELECT * FROM tobe_doc"
        xx=my_cli( fetchall_list(conn,"SELECT status FROM tobe_doc"))
        text5.insert(END,lam1("Doctype:"+x6 + "\nDoc folder: "+ fdpath.get() +"\nTongPeiFu:" + fltpf.get() + "\nExport: tobe_doc\ntotalitem:"+ str(len(xx)) + "\nNumber of OK:"+ str(xx.count("OK"))))
        #保存到文件，并打开，以备修改
        dbtoxls(DB_FILE_PATH,os.path.join(DC1['dbpath'],"tobe_doc.xls"),'tobe_doc')
        os.startfile(os.path.join(DC1['dbpath'],"tobe_doc.xls"))
        b2['state'] = 'normal'
        b1['state'] = 'disabled'
        b3['state'] = 'normal'
        doctypeChosen['state']='disabled'
    def bt5_bt2():
        '''Export.xls'''
        dbtoxls(DB_FILE_PATH, os.path.join(DC1['dbpath'], "tobe_doc.xls"), 'tobe_doc')
        os.startfile(os.path.join(DC1['dbpath'], "tobe_doc.xls"))
        conn = get_conn(DB_FILE_PATH)
        xx = my_cli(fetchall_list(conn, "SELECT status FROM tobe_doc"))
        text5.insert(END, lam1("Import: tobe_doc\ntotalitem:" + str(len(xx)) + "\nNumber of OK:" + str(xx.count("OK"))))
        if  len(xx) == xx.count("OK"):
            b4['state'] = 'normal'
        else:
            b4['state'] = 'disabled'
    def bt5_bt3():
        '''IMPORT .xls'''
        conn = get_conn(DB_FILE_PATH)
        runcate_del(conn,"tobe_doc")
        xlstodb(os.path.join(DC1['dbpath'], "tobe_doc.xls"), DB_FILE_PATH ,"tobe_doc")
        # 检查是不是全部都是ok，并且FROM 文件存在，是的话 使能 COPY
        xx = my_cli(fetchall_list(conn, "SELECT status FROM tobe_doc"))
        text5.insert(END, lam1("Import: tobe_doc\ntotalitem:" + str(len(xx)) + "\nNumber of OK:" + str(xx.count("OK"))))
        os.startfile(os.path.join(DC1['dbpath'], "tobe_doc.xls"))
        if  len(xx) == xx.count("OK"):
            b4['state'] = 'normal'
        else:
            b4['state'] = 'disabled'
    def bt5_bt4():
        '''
        copy file:from tobe_doc
        1.先检查工位文件夹是否存在,
        2.检查from文件是否存在
        3.创建文件夹,并复制
        :return:
        '''
        conn = get_conn(DB_FILE_PATH)
        # 检查是不是全部都是ok，并且FROM 文件存在，是的话 使能 COPY
        xx = my_cli(fetchall_list(conn, "SELECT status FROM tobe_doc"))
        bc1=True
        desc1=""
        data=[]
        x10=datetime.datetime.fromtimestamp(time.time()).strftime("%Y-%m-%d")
        if len(xx) == xx.count("OK"):
            lt1=fetchall_list(conn,"SELECT id_flntime,tkid_custid,tkid_stn,custid_stn,id_doctype,file_fullpath,to_fullpath,fldate,temp_fullpath from tobe_doc where status='OK'")
            for x1,x2,x3,x4,x5,x6,x7,x8,x12 in lt1:
                if os.path.exists(x6)==False:
                    bc1=False
                    desc1 = "NOT EXIST:"+x6
                if os.path.exists(os.path.join(DC1['docroot'],x2))==False:
                    bc1=False
                    desc1 = "NOT EXIST:" + os.path.join(DC1['docroot'],x2)
        if  bc1 :
            for x1,x2,x3,x4,x5,x6,x7,x8,x12 in lt1:
                mydir = os.path.split(x12)[0]
                if os.path.exists(mydir) == False:
                    os.makedirs(mydir)
                    shutil.copy2(x6, x12)

                x9=os.path.join(DC1['dbpath'],"BK",x1,x8,x1)
                mydir = os.path.split(x9)[0]
                if os.path.exists(mydir) == False:
                    os.makedirs(mydir)
                shutil.copy2(x6, x9)
                data.append([x1,x2,x3,x4,x5,"OK",x6,x7,x8,x10,x12])
            # 登入数据库
            save_sql = '''INSERT INTO doc values (?,?,?,?,?,?,?,?,?,?,?)'''
            conn = get_conn(DB_FILE_PATH)
            save(conn, save_sql, data)
            text5.insert(END, lam1("COPY DONE " +"\nafter check please copy manually" ))
            b2['state'] = 'disabled'
            b3['state'] = 'disabled'
            b4['state'] = 'disabled'
            b5['state'] = 'disabled'
        else    :
            text5.insert(END,lam1(desc1))
        os.startfile(os.path.join(DC1['dbpath'],"TOCUST"))
        os.startfile(DC1['docroot'])
    w_bt5 = Toplevel(myWindow)      # 按钮3 创建的 窗口，用于导入导出 STATION表
    w_bt5.title('DOC THE FILE')
    frame=Frame(w_bt5)              # 按钮3创建的表还有一排按钮，bt3_bt1 bt3_bt1 bt3_bt1 bt3_bt1
    doctype = StringVar()
    fdpath= StringVar()
    fltpf= StringVar()
    doctypeChosen =  ttk.Combobox(frame, width=12, textvariable=doctype,font=('Helvetica 10 bold'),height=3)
    conn = get_conn(DB_FILE_PATH)
    datax=fetchall_list(conn,"SELECT id_doctype FROM doc_type ")
    doctypeChosen['values'] = datax  # 设置下拉列表的值
    doctypeChosen.grid( row=0 ,column=0 , sticky=W, padx=5, pady=5 )  # 设置其在界面中出现的位置  column代表列   row 代表行
    doctypeChosen.current(0)  # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值
    b1 = Button(frame, command=bt5_bt1, text='Read Doc', relief='raised', font=('Helvetica 10 bold'),width=8, height=2)
    b1.grid(row=0, column=1, sticky=W, padx=5, pady=5)
    b2 = Button(frame, command=bt5_bt2, text='EXP.xls', font=('Helvetica 10 bold'), width=8, height=2)
    b2.grid(row=0, column=2, sticky=W, padx=5, pady=5)
    b3 = Button(frame, command=bt5_bt3, text='IMP.xls', font=('Helvetica 10 bold'), width=8, height=2)
    b3.grid(row=0, column=3, sticky=W, padx=5, pady=5)
    b4 = Button(frame, command=bt5_bt4, text='COPY', font=('Helvetica 10 bold'), width=8, height=2)
    b4.grid(row=0, column=4, sticky=W, padx=5, pady=5)
    #b2['state'] = 'normal'
    b2['state'] = 'disabled'
    b3['state'] = 'disabled'
    b4['state'] = 'disabled'
    frame.pack()
    text5 = ScrolledText(w_bt5)
    text5.pack(expand=YES, fill=BOTH)
    conn = get_conn(DB_FILE_PATH)
    fetchall_st(conn, "SELECT * FROM doc_type ",text5)
def update_docstatus():
    '''
    因为记录在数据库种的文档随着项目的进展会发生一些改变,例如 原来有的工位,现在没有了,原来路径的文档,想在不存在了,所以需要用
    函数来跑一遍状态,以方便后面操作
    :return:
    '''
    sql = "SELECT status FROM doc"
    conn = get_conn(DB_FILE_PATH)
    # 输出有多少个工位,有多少种文档,有多少种状态
    li1 = my_cli(fetchall_list(conn, "SELECT tkid_custid from station"))
    dict1 = {}
    for i in li1:
        if i not in dict1:
            if os.path.exists(os.path.join(DC1['docroot'],i))==True:
                x1="EXIST"
            else:
                x1="NOT_EXIST"
            dict1[str(i)] = x1
    print(dict1)
    data=[]
    li2 = fetchall_list(conn, "SELECT tkid_custid,to_fullpath,status from doc")
    for x1,x2,x4 in li2:
        x3 = ""
        if x1 in dict1:
            if dict1[x1]=="EXIST":
                if os.path.exists(x2):
                    x3="OK"
                else:
                    x3="FileMiss"
            else:
                x3="StationMiss"
        else:
            x3="StationUnknown"
        if x4!=x3:#状态需要变化才进行更新
            data.append((x3,x1,x2))

    update_sql = 'UPDATE doc SET status = ? WHERE tkid_custid = ? AND to_fullpath = ? '
    conn = get_conn(DB_FILE_PATH)
    update(conn, update_sql, data)
def bt6():
    b1=False
    def bt6_bt5():
        update_docstatus()
    def bt6_bt4():
        """
        导出doc到 应用程序所在目录， doc.xls 文件
        :return:
        """
        text6.delete(1.0,END)
        text6.insert(END,"EXPORT STATION:")
        text6.insert(END, "\n")
        text6.insert(END, os.path.join(DC1['dbpath'], "doc.xls"))
        text6.insert(END, "\n")
        #DC1['dbpath']
        dbtoxls(DB_FILE_PATH, os.path.join(DC1['dbpath'],"doc.xls"), "doc")
        os.startfile(os.path.join(DC1['dbpath'],"doc.xls"))
        nonlocal b1
        b1=True
        sql="SELECT status FROM doc"
        conn = get_conn(DB_FILE_PATH)
        # 输出有多少个工位,有多少种文档,有多少种状态
        li1=my_cli(fetchall_list(conn,"SELECT status from doc"))
        dict1 = {}
        for i in li1:
            if i not in dict1:
                dict1[str(i)] = 0
            dict1[str(i)] += 1
        li1=my_cli(fetchall_list(conn,"SELECT id_doctype from doc"))
        dict2 = {}
        for i in li1:
            if i not in dict2:
                dict2[str(i)] = 0
            dict2[str(i)] += 1
        text6.insert(END,lam1( "count: status\n"+ str(dict1) +"\ncount: doc_type\n"+ str(dict2)))
    w_bt6 = Toplevel(myWindow)      # 按钮3 创建的 窗口，用于导入导出 STATION表
    w_bt6.title('project doc in database')
    frame=Frame(w_bt6)              # 按钮3创建的表还有一排按钮，bt3_bt1 bt3_bt1 bt3_bt1 bt3_bt1
    b4 = Button(frame, command=bt6_bt4, text='EXP_XLS', font=('Helvetica 10 bold'), width=8, height=2)
    b4.grid(row=0, column=3, sticky=W, padx=5, pady=5)
    b5 = Button(frame, command=bt6_bt5, text='UPDATE', font=('Helvetica 10 bold'), width=8, height=2)
    b5.grid(row=0, column=4, sticky=W, padx=5, pady=5)
    frame.pack()
    text6 = ScrolledText(w_bt6)
    text6.insert(END,'''
Updates include:

1. Check whether the station document folder exists on the hard disk.

2. Check whether the station in doc table exists in the station table

3. Check whether the document information exists on the hard disk

4. Deletion is not supported at present
    ''')
    text6.pack(expand=YES, fill=BOTH)

def btn_x():
    wdn = DC2['wdn']
    tbn = DC2['tbn']
    dbn = DC2['dbn']
    xln = DC2['xln']
    addone = DC2['addone']
    delone = DC2['delone']
    b1 = False  # 必须先导出，然后才能导入，当 btn3_btn4 被执行后，本值被设定为True，bt3_btn5 被执行后再次设定为false
    def bt3_bt4():
        text3.insert(END,lam1("EXPORT TABLE:"+tbn+'\n'+xln))
        try:
            dbtoxls(dbn,xln,tbn)
        except:
            text3.insert(END,lam1("ERROR:"+xln+"\nWAS OPEN CANNOT WRITE."))
            return
        os.startfile(xln)
        nonlocal b1
        b1 = True
        sql = "SELECT * FROM "+tbn
        text3.insert(END, lam1(sql))
        conn = get_conn(dbn)
        fetchall_st(conn, sql, text3)
    def bt3_bt5():
        if b1 == False:
            text3.insert(END, lam1("CAN NOT IMPORT WITHOUT EXPORT:"+tbn))
            return
        text3.insert(END, lam1("IMPORT " + tbn +" FROM:"+xln))
        if os.path.exists(xln) == False:
            text3.insert(END, lam1("FILE DOES NOT EXIST:" + xln))
            return
        conn = get_conn(dbn)
        runcate_del(conn, tbn)
        xlstodb(xln, dbn,tbn)
        sql = "SELECT * FROM "+tbn
        conn = get_conn(dbn)
        fetchall_st(conn, sql, text3)
    def bt3_bt6():
        nonlocal tbn ,dbn
        # 用空格隔开的字符串
        myWindow.withdraw()
        res = askstring("insert into " + tbn, "Space split without braces")
        myWindow.deiconify()
        data = [["", "", "", ""]]
        x1 = []
        x = re.split(r" (?![^{]*\})", res)
        x2=len(x)
        #save_sql='INSERT INTO '+tbn+" values ("+" ?,"*(x2-1)+" ?)"
        save_sql = 'INSERT INTO ' + tbn + " values (" + " ?," * (x2 - 1) + " ?)"
        print(x2)
        for xx in x:
            if xx[0] == "{" and xx[-1] == "}":
                x1.append(xx[1:len(xx[0]) - 2])
            else:
                x1.append(xx)
        data[0] = x1
        # data=[["D.05899.310(T119 {40})","D.05899.310","(T119 {40})","OK"]]
        conn = get_conn(DB_FILE_PATH)
        cu = conn.cursor()
        save(conn, save_sql, data)
        fetchall_st(conn,"Select * from "+ tbn,text3)
    def bt3_bt7():
        nonlocal tbn
        myWindow.withdraw()
        res = askstring("delete from " + tbn, "Space split without braces")
        myWindow.deiconify()
        data = [["", "", "", ""]]
        x1 = []
        x = re.split(r" (?![^{]*\})", res)
        x2=len(x)
        #save_sql='INSERT INTO '+tbn+" values ("+" ?,"*(x2-1)+" ?)"
        save_sql = 'DELETE FROM ' + tbn + " WHERE values (" + " ?," * (x2 - 1) + " ?)"
        print(x2)
        for xx in x:
            if xx[0] == "{" and xx[-1] == "}":
                x1.append(xx[1:len(xx[0]) - 2])
            else:
                x1.append(xx)
        data[0] = x1
        # data=[["D.05899.310(T119 {40})","D.05899.310","(T119 {40})","OK"]]
        conn = get_conn(DB_FILE_PATH)
        cu = conn.cursor()
        # 获取表名，保存在tab_name列表
        cu.execute("select name from sqlite_master where type='table'")
        tab_name = cu.fetchall()
        tab_name = [line[0] for line in tab_name]
        # 获取表的列名（字段名），保存在col_names列表,每个表的字段名集为一个元组
        col_names = []
        cu.execute('pragma table_info({})'.format(tbn))
        col_name = cu.fetchall()
        col_name = [x[1] for x in col_name]
        col_names.append(col_name)
        col_name = tuple(col_name)
        sql='DELETE FROM ' + tbn + " WHERE "
        for xxx in col_name:
            sql += (   xxx + ' = ? AND ')
        sql=sql[:-4]
        save(conn, sql, data)
        fetchall_st(conn, "Select * from " + tbn, text3)
    w_bt3 = Toplevel(myWindow)  # 按钮3 创建的 窗口，用于导入导出 STATION表
    w_bt3.title('Exp_Imp_'+wdn)
    frame = Frame(w_bt3)  # 按钮3创建的表还有一排按钮，bt3_bt1 bt3_bt1 bt3_bt1 bt3_bt1
    b4 = Button(frame, command=bt3_bt4, text='EXP_XLS', font=('Helvetica 10 bold'), width=8, height=2)
    b4.grid(row=0, column=3, sticky=W, padx=5, pady=5)
    b5 = Button(frame, command=bt3_bt5, text='IMP_XLS', font=('Helvetica 10 bold'), width=8, height=2)
    b5.grid(row=0, column=4, sticky=W, padx=5, pady=5)
    if addone==True:
        b6 = Button(frame, command=bt3_bt6, text='ADD_ONE' , font=('Helvetica 10 bold'), width=8, height=2)
        b6.grid(row=0, column=5, sticky=W, padx=5, pady=5)
    if delone==True:
        b7 = Button(frame, command=bt3_bt7, text='DEL_ONE' , font=('Helvetica 10 bold'), width=8, height=2)
        b7.grid(row=0, column=6, sticky=W, padx=5, pady=5)
    frame.pack()
    text3 = ScrolledText(w_bt3)
    text3.pack(expand=YES, fill=BOTH)
    sql = "SELECT * FROM " + tbn
    conn = get_conn(dbn)
    fetchall_st(conn, sql, text3)
    text3.see(END)
def bt7():
    DC2["wdn"],DC2["tbn"],DC2["dbn"],DC2["xln"],DC2['addone'],DC2['delone']="Document type","doc_type",DB_FILE_PATH,os.path.join(DC1['dbpath'],'doc_type.xls'),True,True
    btn_x()
def bt8():
    DC2["wdn"],DC2["tbn"],DC2["dbn"],DC2["xln"],DC2['addone'],DC2['delone']="init setting","dbinit",DB_FILE_PATH,os.path.join(DC1['dbpath'],'dbinit.xls'),False,False
    btn_x()
def bt9():
    DC2["wdn"],DC2["tbn"],DC2["dbn"],DC2["xln"],DC2['addone'],DC2['delone']="memo_milestone","milestone",DB_FILE_PATH,os.path.join(DC1['dbpath'],'milestone.xls'),True,True
    btn_x()
if __name__ == '__main__':
    myWindow = Tk()
    myWindow.title('Project Doc manage')
    frame = Frame()  # 定义容器
    b3=Button(frame, command=bt3 ,text='STN...', font=('Helvetica 10 bold'),width=8, height=2)
    b3.grid(row=0, column=2, sticky=E, padx=5,pady=5)
    b5=Button(frame, command=bt5 ,text='ADD_DOC...', font=('Helvetica 10 bold'),width=8, height=2)
    b5.grid(row=0, column=4, sticky=W, padx=5,pady=5)
    b6=Button(frame, command=bt6 ,text='DOC...', font=('Helvetica 10 bold'),width=8, height=2)
    b6.grid(row=0, column=5, sticky=W, padx=5,pady=5)
    b7 = Button(frame, command=bt7, text='DOC_TYP...', font=('Helvetica 10 bold'), width=8, height=2)
    b7.grid(row=0, column=6, sticky=W, padx=5, pady=5)
    b8 = Button(frame, command=bt8, text='INI...', font=('Helvetica 10 bold'), width=8, height=2)
    b8.grid(row=0, column=7, sticky=W, padx=5, pady=5)
    b9 = Button(frame, command=bt9, text='MST...', font=('Helvetica 10 bold'), width=8, height=2)
    b9.grid(row=0, column=8, sticky=W, padx=5, pady=5)
    frame.pack()
    myinit()
    text = ScrolledText(myWindow)
    text.insert(END,DC1['info'])
    text.pack(expand=YES, fill=BOTH)
    myWindow.title('Project Doc manage-'+DC1['pjn'])
    myWindow.mainloop()

