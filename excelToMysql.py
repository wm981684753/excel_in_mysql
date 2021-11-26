# coding:utf-8
# 基于python3的 简单的单表excel入库脚本
# author weiming 2021.11
import math

import xlrd
import pymysql
import time
import os

# 数据库默认配置
config = {
    "host": "localhost",
    "database": "excel",
    "username": "root",
    "password": "root",
    "fieldlen": "330",
    "maxfield": "50",
}


# 主程序
def main():
    # 提示信息
    is_config = input("使用说明："
                      "\n请提前配置好数据库，详情参考 README.txt,字段默认使用varchar"
                      "\n只能导入.xls文件，并且 有且仅有一行表头(表头将作为字段备注)，且文件名为英文（将作为表名），参考demo.xls"
                      "\nMysql默认配置：\n\t数据库地址：localhost\n\t数据库名称：excel\n\t用户名：root\n\t密码：root\n\t单个字段长度：330（vachar）\n\t单表字段数量：50（mysql有单行字段数据限制65535，所以一般不要太大，建议不超过50）"
                      "\n如果使用默认配置 请按回车键继续，若要自定义配置请输入 [Y]：")

    # 数据库连接测试
    db_connect(is_config)

    # 执行入库程序
    run()


# 数据库连接测试
def db_connect(is_config):
    # 自定义配置
    if (is_config.lower() == "y"):
        set_config()

    # ------数据库配置
    host = config['host']  # 数据库地址
    database = config['database']  # 数据库名称
    username = config['username']  # 数据库用户名
    password = config['password']  # 数据库密码

    try:
        pymysql.connect(host=host, port=3306, user=username, passwd=password, db=database, charset='utf8')
    except pymysql.Error as e:
        print("数据库链接失败，请重新配置")
        print(e.args[0], e.args[1])
        return db_connect(is_config)
    return


# 数据库自定义配置
def set_config():
    input_host = input("数据库地址：")
    input_database = input("数据库名称：")
    input_username = input("用户名：")
    input_password = input("密码：")
    input_fieldlen = input("单个字段长度（使用默认直接回车）：")
    input_maxfield = input("单表字段数量（使用默认直接回车）：")
    if (len(input_fieldlen) == 0):
        input_fieldlen = config['fieldlen']
    if (len(input_maxfield) == 0):
        input_maxfield = config['maxfield']
    config['host'] = input_host
    config['database'] = input_database
    config['username'] = input_username
    config['password'] = input_password
    config['fieldlen'] = str(input_fieldlen)
    config['maxfield'] = str(input_maxfield)


# 选择文件
def choose_file():
    # 需要导入的文件地址
    file = input('请输入你要导入的excel文件路径(.xls)：')

    # 文件校验
    ext = os.path.splitext(file)[1]
    if (ext not in [".xls"]):
        print("只能选择.xls文件")
        return choose_file()

    if (os.path.isfile(file) != True):
        print("选择的文件不存在")
        return choose_file()

    return file

# 入库执行程序
def run():
    # ------数据库配置
    host = config['host']  # 数据库地址
    database = config['database']  # 数据库名称
    username = config['username']  # 数据库用户名
    password = config['password']  # 数据库密码

    # ------选择文件
    file = choose_file()

    # 程序执行
    t1 = time.time()
    print("开始入库")
    db = pymysql.connect(host=host, port=3306, user=username, passwd=password, db=database, charset='utf8')
    cur = db.cursor()
    ex = xlrd.open_workbook(file)
    sheet = ex.sheet_by_index(0)
    rows = sheet.nrows
    cols = sheet.ncols

    # 获取表头
    excelTitle = []
    for col_num in range(cols):
        title = db.escape_string(sheet.row_values(0, col_num)[0])
        excelTitle.append(title)

    # 需要写入的数据表名称(使用文件名)
    table = os.path.splitext(os.path.basename(file))[0]

    # 生成excel表头数组
    alpArr = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U",
              "V", "W", "X", "Y", "Z"]
    excelAlpTitle = alpArr
    for x in range(26):
        for y in range(26):
            if (len(excelTitle) < 1000):
                excelAlpTitle.append(alpArr[x] + alpArr[y])

    # 根据字段长度，需要做拆分表处理
    maxFieldNum = int(config['maxfield'])  # 单表最大字段数
    tableNum = math.ceil(len(excelTitle) / maxFieldNum)  # 需要拆分成的表的数量
    tableStr = ""  # 本次写入的表
    dataCount = 0  # 插入的总数据量
    for num in range(0, tableNum):
        # 先判断数据库是否有表，若无当前表，则自动创建
        if (num == 0):
            tableName = table
        else:
            tableName = table + str(num + 1)
        if (len(tableStr) > 0):
            tableStr = tableStr + "," + tableName
        else:
            tableStr = tableName
        drop = "DROP TABLE IF EXISTS `" + tableName + "`;"
        try:
            cur.execute(drop)
            db.commit()
        except pymysql.Error as e:
            print("数据库建表异常，停止运行")
            print(e.args[0], e.args[1])
            return run()

        create = "CREATE TABLE `" + tableName + "` ("
        create += "`id` int(11) NOT NULL AUTO_INCREMENT,"
        for index, item in enumerate(excelTitle[num * maxFieldNum:(num + 1) * maxFieldNum]):
            index = index + num * maxFieldNum
            create += "`e_" + excelAlpTitle[
                index] + "` varchar(" + config[
                          'fieldlen'] + ") COLLATE utf8_unicode_ci DEFAULT NULL COMMENT '" + item + "',"
        create += "PRIMARY KEY (`id`)"
        create += ") ENGINE=MyISAM AUTO_INCREMENT=70001 DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;"
        try:
            cur.execute(create)
            cur.execute("truncate table " + tableName)
            db.commit()
        except pymysql.Error as e:
            print("数据库写入异常，请重新选择文件")
            print(e.args[0], e.args[1])
            return run()

        # 跳过第一行的标题，直接从第二行真实数据开始
        for row_num in range(1, rows):
            # 循环取出每一列的值并写入数据表
            fields = values = ""
            for index, item in enumerate(excelTitle[num * maxFieldNum:(num + 1) * maxFieldNum]):
                index = index + num * maxFieldNum
                data = db.escape_string(str(sheet.row_values(row_num, 0)[index]))
                fields = fields + ",e_" + excelAlpTitle[index]
                values = values + ',"' + data + '"'
            # 截取字符串，丢弃第一个字符
            fields = fields[1:]
            values = values[1:]
            sql = "insert into " + tableName + "(" + fields + ") VALUES (" + values + ")"
            try:
                cur.execute(sql)
                db.commit()
            except pymysql.Error as e:
                print("数据库写入异常，请重新选择文件")
                print(e.args[0], e.args[1])
                return run()
            dataCount = dataCount + 1

    cur.close()
    db.close()
    t2 = time.time()
    print("一共用时：" + str(t2 - t1) + " 秒")
    print("共写入 " + str(tableNum) + " 张表：" + tableStr)
    print("单表字段：" + str(maxFieldNum) + " 总字段：" + str(len(excelTitle)))
    print("单表写入：" + str(int(dataCount / tableNum)) + " 条")
    print("总共写入：" + str(dataCount) + " 条")
    print("入库完成")
    input("入库完成，按回车键退出...")

# 主程序自动执行
main()
