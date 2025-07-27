from openpyxl import load_workbook
import time
from openpyxl.utils.exceptions import InvalidFileException
import os
'''
用于多日统计的工具
'''

# 定义常量
COMPLETE = 100
PASS = 60

#定义变量
OUTPUT_DIR = ""
INPUT_DIR = ""
unpassName = []
incompleteName = []
#设定班级人数，确定表格范围
NOS = int(input("班级人数,按表格中人名总数填写"))+3#NumberOfStudents,班级人数
NOS = str(NOS)
