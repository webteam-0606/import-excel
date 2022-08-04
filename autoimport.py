# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import xlrd
import sqlite3
import pprint


# 连接数据库
def connect_db(file_path):
    conn = sqlite3.connect(file_path)
    return conn


# 获取数据库中所有表的名字
def get_tables(conn):
    # sql = "SELECT * FROM sys.Tables"
    sql = "SELECT * FROM uploadsearch"
    cursor = conn.cursor()
    # 获取表名
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = [tuple[0] for tuple in cursor.fetchall()]
    print(tables)
    return tables


# 获取数据库中，表table_name 的表头信息，列名称
def get_desc(conn, table_name):
    cursor = conn.cursor()
    sql1 = "select * from {}".format(table_name)
    cursor.execute(sql1)
    col_name_list = [tuple[0] for tuple in cursor.description]
    sql = "("
    for index in col_name_list:
        sql += index + ","
    ret = sql[:-1] + ")"
    return ret


# 显示数据库中表table_name 的所有元素
def show_table(conn, table_name):
    cursor = conn.cursor()
    sql = "select * from {}".format(table_name)
    cursor.execute(sql)
    pprint.pprint(cursor.fetchall())


# 创建数据库，table_items 为 table_name 中列名称，即表头信息
def create_table(conn, table_name, table_items):
    sqlline = "create table {} (".format(table_name)
    for i in table_items:
        sqlline += i + " text,"
    sql_line = sqlline[:-1] + ")"
    cursor = conn.cursor()
    cursor.execute(sql_line)
    conn.commit()


# 数据库文件插入，content_items 为需要插入表 table_name 的数据信息
def insert_data(conn, table_name, content_items):
    sql = ''' insert into {} 
    {}
    values ('''.format(table_name, get_desc(conn, table_name))
    for index in content_items:
        sql += str(index) + ","
    ret = sql[:-1] + ")"
    cursor = conn.cursor()
    cursor.execute(ret)
    conn.commit()


#数据库中table_name表中查找 table_head = table_content 的项
def find_data(conn, table_name, table_head, table_content):
    sql = "select {table_head} from {table_name} where {table_head} = {table_content}".format(table_head=table_head,
                                                                                              table_name=table_name,
                                                                                              table_content=table_content)
    cursor = conn.cursor()
    cursor.execute(sql)
    pprint.pprint(cursor.fetchone())


# 读取exel表格，并在数据库中创建该表
def read_exel(file_path, conn):
    if not file_path.endswith("xlsx"):
        print("path_wrong")
    # 获取一个Book对象
    book = xlrd.open_workbook(file_path)
    # 获取一个sheet对象的列表
    sheets = book.sheets()
    for sheet in sheets:
        sheet_name = sheet.name
        # 获取表行数
        rows = sheet.get_rows()
        for index, row in enumerate(rows):
            table_items = [tuple.value for tuple in row]
            print(table_items)
            if index == 0:
                # 默认第一行为表头信息，在数据库中创建该表
                create_table(conn, sheet_name, list(table_items))
            else:
                # 将次sheet中的每一行都插入数据库中
                insert_data(conn, sheet_name, table_items)
        show_table(conn, sheet_name)


def main():
    # Use a breakpoint in the code line below to debug your script.
    conn = connect_db("uploadsearch.db")
    table = get_tables(conn)
    # find_data(conn,table[0],"测试",2.0)
    find_data(conn,table[0],"auto",2.0)
    # def find_data(conn, table_name, table_head, table_content):
    show_table(conn,table[0])
    file_path = "C:\\Users\\320200255\\Desktop\\1BRILLIANCE-5.xlsx"
    read_exel(file_path, conn[0])
    conn.close()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

