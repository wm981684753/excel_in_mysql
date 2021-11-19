# coding:utf-8

import ConfigParser
import os

# 用os模块来读取
curpath = os.path.dirname(os.path.realpath(__file__))
cfgpath = os.path.join(curpath, "config.ini")  # 读取到本机的配置文件

# 调用读取配置模块中的类
conf = ConfigParser.ConfigParser()
conf.read(cfgpath)

host=conf.get("databases","host")
database=conf.get("databases","database")
username=conf.get("databases","username")
password=conf.get("databases","password")
