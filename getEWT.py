from openpyxl import load_workbook
import time
from openpyxl.utils.exceptions import InvalidFileException
from os import listdir

'''
用于多日统计的工具
'''

# 定义常量
COMPLETE = 100
PASS = 60

# 定义变量
OUTPUT_DIR = ""
INPUT_DIR = "/mnt/sharedisk/Task/"
# 设定班级人数，确定表格范围
NOS = int(input("班级人数,按表格中人名总数填写")) + 3  # NumberOfStudents,班级人数
NOS = str(NOS)


def getData(cellRange, sheet):
    Data = []
    for location in cellRange:
        cellLocation = str(location)[20:][0:-3]
        Data.append(sheet[cellLocation].value)
    return Data

def dirXlsx(_DIR):
    files=[]
    for file in listdir(_DIR):
        files.append(_DIR+file)
    print(files)
    return files

def statistic(_DIR, NOS):
    inputWorkbook = load_workbook(_DIR)
    sheet = inputWorkbook.active
    time.sleep(0.5)
    sheet.title = "MAIN SHEET" #确保字符串长度一致
    cellRange_Name = sheet["C3":"C" + NOS]
    cellRange_completePercentage = sheet["H3":"H" + NOS]
    # 获取原始数据
    name = getData(cellRange_Name,sheet)
    completePercentage = getData(cellRange_completePercentage,sheet)
    # 接下来检查完成率，并将不合格的标记出来
    step = 0
    unpassName = []
    incompleteName = []
    incompletePer = []
    unpassPer = []
    for per in completePercentage:
        per = float(per)
        if per < COMPLETE and per >= PASS:
            incompleteName.append(name[step])
            incompletePer.append(str(per))

        if per < PASS:
            unpassName.append(name[step])
            unpassPer.append(str(per))
        step += 1
    print("未完成名单：", incompleteName, "\n", incompletePer)
    print("完成率小于60：", unpassName, "\n", unpassPer)

files = dirXlsx(INPUT_DIR)#扫描目录文件
for file in files:
    statistic(file,NOS)