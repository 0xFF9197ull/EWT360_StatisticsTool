from openpyxl import load_workbook
import time

from openpyxl.descriptors.excel import Percentage
from openpyxl.utils.exceptions import InvalidFileException
from os import listdir

'''
用于多日统计的工具
'''

# 定义常量
COMPLETE = 100
PASS = 60

# 定义变量

INPUT_DIR = input("文件路径:")
OUTPUT_DIR = INPUT_DIR+"检测结果.md"

InCompleteName = []  # 与下面函数内的变量不同，这里记录总人数
Names = []
frequency = []
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
    files = []
    for file in listdir(_DIR):
        if file[-5:] == ".xlsx":#注意文件后缀
            files.append(_DIR + file)
    print(files)
    return files


def statistic(_DIR, NOS):
    inputWorkbook = load_workbook(_DIR)
    sheet = inputWorkbook.active
    # time.sleep(0.5)
    sheet.title = "MAIN SHEET"  # 确保字符串长度一致
    cellRange_Name = sheet["C3":"C" + NOS]
    cellRange_completePercentage = sheet["H3":"H" + NOS]
    # 获取原始数据
    name = getData(cellRange_Name, sheet)
    completePercentage = getData(cellRange_completePercentage, sheet)
    # 接下来检查完成率，并将不合格的标记出来
    step = 0
    unpassName = []
    incompleteName = []
    incompletePer = []
    unpassPer = []
    for per in completePercentage:
        per = float(per)
        if per < COMPLETE:
            incompleteName.append(name[step])
            incompletePer.append(str(per))
        step += 1
    print("未完成名单：", incompleteName, "\n", incompletePer)
    return incompleteName


files = dirXlsx(INPUT_DIR)  # 扫描目录文件
for file in files:
    InCompleteName.extend(statistic(file, NOS))
    #print(InCompleteName)

while InCompleteName != []:
    name = InCompleteName[0]
    Names.append(name)
    frequency_ = InCompleteName.count(name)  # 每次从第一个开始查找，然后删掉
    frequency.append(frequency_)
    for k in range(frequency_):
        InCompleteName.remove(name)
print("#"*25)
print(Names, frequency)
#保存结果
formatted_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
with open(OUTPUT_DIR, "w") as file:
    index_ = 0
    file.write("# " + "生成时间：" + formatted_time + "\n")
    file.write("-------" + "\n" + "\n")
    file.write("|姓名|未完成次数|" + "\n")
    file.write("|:-----:|:-----:|" + "\n")
    for a in Names:
        file.write("|" + Names[index_])
        file.write("|" + str(frequency[index_]) + "|" + "\n")
        index_ += 1
print("统计完成，结果已输出至：" + OUTPUT_DIR)