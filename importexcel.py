# """ 配置数据库 """
# HOST='127.0.0.1'
# PORT='3306'
# USERNAME='root'
# PASSWORD='root'
# DATABASE='uploadsearch'

# select * from character

import pymysql
import xlrd
import sqlite3
import pprint

import pandas as pd
# 用面向对象的方式编写，更加熟悉面向对象代码风格


class Mysql_csv(object):
    # 定义一个init方法，用于读取数据库
    def __init__(self):
        # 读取数据库和建立游标对象
        # self.connect = pymysql.connect(host="127.0.0.1",port=3306,user="root",password="root",database="weather",charset="utf8")
        self.connect = pymysql.connect(host="127.0.0.1", port=3306,
                                       user="root", password="root", database="uploadsearch", charset="utf8")
        self.cursor = self.connect.cursor()
    # 定义一个del类，用于运行完所有程序的时候关闭数据库和游标对象

    def __del__(self):
        self.connect.close()
        self.cursor.close()

    def read_csv_columns(self):
        # 读取csv文件的列索引，用于建立数据表时的字段
        # csv_name='D:\\code\\python project\\pachong\\weather_project\\weather_year.csv'
        csv_name = 'C:\\Users\\320200255\\Desktop\\1BRILLIANCE-5.csv'
        data = pd.read_csv(csv_name, encoding="utf-8")
        return data

    def read_csv_values(self):
        # 读取csv文件数据
        # csv_name='D:\\code\\python project\\pachong\\weather_project\\weather_year.csv'
        csv_name = 'C:\\Users\\320200255\\Desktop\\1BRILLIANCE-5.csv'
        data = pd.read_csv(csv_name, encoding="utf-8")
        data_3 = list(data.values)
        return data_3

    def write_mysql(self):

        # 在数据表中写入数据，因为数据是列表类型，把他转化为元组更符合sql语句
        for i in self.read_csv_values():  # 因为数据是迭代列表，所以用循环把数据提取出来            
            data_6 = tuple(i)
            sql = """insert into auto values{}""".format(data_6)
            self.cursor.execute(sql)
            self.commit()
        print("\n数据植入完成")

    def commit(self):
        # 定义一个确认事务运行
        self.connect.commit()

    def create(self):
        # 若已有数据表weather_year_db，则删除
        # table_name = "auto"
        query = "drop table if exists auto;"
        self.cursor.execute(query)
        # --+--
        # 创建数据库，table_items 为 table_name 中列名称，即表头信息        
        # book = xlrd.open_workbook('C:\\Users\\320200255\\Desktop\\1BRILLIANCE-5.csv')
        # sheets = book.sheets()
        for index,i in enumerate(self.read_csv_values()) :  # 因为数据是迭代列表，所以用循环把数据提取出来
            if index ==0:
                data_c = tuple(i)
                # sql = """insert into auto values{}""".format(data_c)
                sql = "create table auto ("
                for j in data_c:
                    sql += i + " text,"
                sql_line = sql[:-1] + ")"
                self.cursor.execute(sql_line)
                self.commit()
        # --+--
        

        # 创建数据表，用刚才提取的列索引作为字段
        # data_2 = self.read_csv_columns()
        # sql = "create table if not exists auto(date_time DATETIME not null,high varchar(50) not null,low varchar(50) not null,weather varchar(50) not null,primary key(date_time))default charset=utf8;"
        # sql = "create table if not exists auto(id varchar(50) not null,Headline varchar(50),ceshi1 varchar(50),primary key(id))default charset=utf8;"

        # sql = "create table if not exists auto default charset=utf8;"
        # self.cursor.execute(sql)
        # self.commit()

    # 运行程序，记得要先调用创建数据的类，在创建写入数据的类
    def run(self):
        self.create()
        self.write_mysql()

# 最后用一个main()函数来封装


def main():
    sql = Mysql_csv()
    sql.run()


if __name__ == '__main__':
    main()
