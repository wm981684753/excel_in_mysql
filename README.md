# excel_in_mysql

基于Python3 一键导入excel到mysql的脚本

只需预先创建一个空数据库
数据表自动生成（字段过多会做拆分表处理）
在config.ini中修改数据库配置
excel.py为运行主程序

dist文件夹包含可直接运行的exe文件，可单独拿出来使用（需按README.txt配置数据库）

tip：

1.下载项目后需要拉取公共依赖包

2.目前只支持.xls文件

3.excel表只支持一行表头


配置文件说明config.ini

host：数据库地址

database：数据库名称

username：用户名

password：密码

fieldlen：单个字段长度（默认类型为varchar）

maxfield：单表最大字段数（建议不超过50）



