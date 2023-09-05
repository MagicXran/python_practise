import os
import sys

from scrapy.cmdline import execute

"""
   启动项目脚本
"""
# print(__file__)  # 获取当前文件的绝对路径
# print(os.path.dirname(os.path.abspath(__file__)))  # 获取当前文件所在目录的绝对路径
# 建一个目录放进python搜索目录中去.
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
#   使用脚本启动spider, scrapy crawl 为命令, 详情:scrapy crawl -h
execute(["scrapy", "crawl", "jobbole"])
