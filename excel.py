# coding:utf-8
# 基于python3的 简单的单表excel入库脚本
# author weiming 2021.11
import math

import xlrd
import pymysql
import time
import os
import config

# ------数据库配置
host = config.host  # 数据库地址
database = config.database  # 数据库名称
username = config.username  # 数据库用户名
password = config.password  # 数据库密码

# ------输入参数
# 需要导入的文件地址
# file = "demo.xls"
file = input('请输入你要导入的excel文件：')

# 文件校验
ext = os.path.splitext(file)[1]
if (ext not in [".xls", ".xlsx"]):
    print("只能选择excel文件")
    exit()

if (os.path.isfile(file) != True):
    print("选择的文件不存在")
    exit()

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
table = os.path.splitext(file)[0]

# 生成excel表头数组
alpArr = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V",
          "W", "X", "Y", "Z"]
excelAlpTitle = alpArr
for x in range(26):
    for y in range(26):
        if (len(excelTitle) < 100):
            excelAlpTitle.append(alpArr[x] + alpArr[y])

# 根据字段长度，需要做拆分表处理
maxFieldNum = int(config.maxfield)  # 单表最大字段数
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
    cur.execute(drop)
    db.commit()
    create = "CREATE TABLE `" + tableName + "` ("
    create += "`id` int(11) NOT NULL AUTO_INCREMENT,"
    for index, item in enumerate(excelTitle[num * maxFieldNum:(num + 1) * maxFieldNum]):
        index = index + num * maxFieldNum
        create += "`" + excelAlpTitle[
            index] + "` varchar("+config.fieldlen+") COLLATE utf8_unicode_ci DEFAULT NULL COMMENT '" + item + "',"
    create += "PRIMARY KEY (`id`)"
    create += ") ENGINE=MyISAM AUTO_INCREMENT=70001 DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;"
    cur.execute(create)
    cur.execute("truncate table " + tableName)
    db.commit()

    # 跳过第一行的标题，直接从第二行真实数据开始
    for row_num in range(1, rows):
        # 循环取出每一列的值并写入数据表
        fields = values = ""
        for index, item in enumerate(excelTitle[num * maxFieldNum:(num + 1) * maxFieldNum]):
            index = index + num * maxFieldNum
            data = db.escape_string(sheet.row_values(row_num, 0)[index])
            fields = fields + "," + excelAlpTitle[index]
            values = values + ',"' + data + '"'
        # 截取字符串，丢弃第一个字符
        fields = fields[1:]
        values = values[1:]
        sql = "insert into " + tableName + "(" + fields + ") VALUES (" + values + ")"
        cur.execute(sql)
        db.commit()
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
