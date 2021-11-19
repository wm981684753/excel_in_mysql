# coding:utf-8
# 基于python2的 简单的单表excel入库脚本
# author weiming 2021.11
import xlrd
import pymysql
import time
import os
# import config
#
# # ------数据库配置
# host = config.host # 数据库地址
# database = config.database # 数据库名称
# username = config.username # 数据库用户名
# password = config.password # 数据库密码

host = "localhost" # 数据库地址
database = "excel" # 数据库名称
username = "root" # 数据库用户名
password = "root" # 数据库密码

# ------参数定义
# 需要导入的文件地址
# file = "taoke0425.xls"
file = input('请输入你要导入的excel文件：')

# 文件校验
ext = os.path.splitext(file)[1]
if(ext not in [".xls",".xlsx"]):
    print("只能选择excel文件")
    exit()

if(os.path.isfile(file)!=True):
    print("选择的文件不存在")
    exit()

# 程序执行
t1 = time.time()
print("开始入库")
db = pymysql.connect(host=host,port=3306, user=username, passwd=password, db=database, charset='utf8')
cur = db.cursor()
ex = xlrd.open_workbook(file)
sheet = ex.sheet_by_index(0)
rows = sheet.nrows
cols = sheet.ncols

# 获取表头，作为字段名称
excelTitle = []
for col_num in range(cols):
    title = db.escape_string(sheet.row_values(0, col_num)[0])
    excelTitle.append(title)

# 需要写入的数据表名称(使用文件名)
table = os.path.splitext(file)[0]

# 先判断数据库是否有表，若无当前表，则自动创建
drop = "DROP TABLE IF EXISTS `"+table+"`;"
cur.execute(drop)
db.commit()
create = "CREATE TABLE `"+table+"` ("
create += "`id` int(11) NOT NULL AUTO_INCREMENT,"
for index,item in enumerate(excelTitle):
    create += "`"+item+"` varchar(500) COLLATE utf8_unicode_ci DEFAULT NULL,"
create += "PRIMARY KEY (`id`)"
create += ") ENGINE=MyISAM AUTO_INCREMENT=70001 DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;"
cur.execute(create)
db.commit()

# 跳过第一行的标题，直接从第二行真实数据开始
count = 0
for row_num in range(1, rows):
    # 循环取出每一列的值并写入数据表
    fields = values = ""
    for index, item in enumerate(excelTitle):
        data = db.escape_string(sheet.row_values(row_num, 0)[index])
        fields = fields+","+item
        values = values+',"'+data+'"'
    # 截取字符串，丢弃第一个字符
    fields = fields[1:]
    values = values[1:]
    sql = "insert into "+table+"("+fields+") VALUES ("+values+")"
    cur.execute(sql)
    db.commit()
    count = count+1
cur.close()
db.close()
t2 = time.time()
print("一共用时："+str(t2 - t1)+" 秒")
print("插入数据："+str(count)+" 条")
print("入库完成")