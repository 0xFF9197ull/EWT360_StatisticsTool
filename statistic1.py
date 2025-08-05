from openpyxl import load_workbook
import time
from openpyxl.utils.exceptions import InvalidFileException


# 定义常量
INPUT_DIR = ""
OUTPUT_DIR = ""
COMPLETE = 100
PASS = 60
unpassName = []
incompleteName = []
unpassPer = []
incompletePer = []
# 准备工作表
input_OK = False
#尝试导入，出错后重新获取路径
while input_OK == False:
    try:
        INPUT_DIR= input("输入表格路径")
        inputWorkbook = load_workbook(INPUT_DIR)
        input_OK=True

    except InvalidFileException as error:
        print("Error:", error)
sheet = inputWorkbook.active
sheet.title = "MAIN SHEET"
# 获取数据,名称和完成率
time.sleep(1)#防止input与报错冲突
NOS = int(input("班级人数,按表格中人名总数填写"))+3#NumberOfStudents,班级人数
NOS = str(NOS)
cellRange_Name = sheet["C3":"C"+NOS]
cellRange_completePercentage = sheet["H3":"H"+NOS]

xlsxTime = sheet["A1"].value
print(xlsxTime)

def getData(cellRange):
    Data = []
    for location in cellRange:
        cellLocation = str(location)[20:][0:-3]
        Data.append(sheet[cellLocation].value)
    return Data


name = getData(cellRange_Name)
completePercentage = getData(cellRange_completePercentage)
# 接下来检查完成率，并将不合格的标记出来
step = 0
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

formatted_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
# 以markdown形式输出结果
input_OK = False
while input_OK == False:
    try:
        OUTPUT_DIR= input("输出结果路径")
        with open(OUTPUT_DIR, "w") as file:
            index_ = 0
            file.write("# " + xlsxTime + "\n")
            file.write("> " + "生成时间：" + formatted_time + "\n")
            file.write("-------" + "\n" + "\n")
            file.write("|姓名|完成率|" + "\n")
            file.write("|:-----:|:-----:|" + "\n")
            for a in incompleteName:
                if incompleteName == []:
                    break
                file.write("|" + incompleteName[index_])
                file.write("|" + incompletePer[index_] + "|" + "\n")
                index_ += 1
            index_ = 0
            for a in unpassName:
                if unpassName == []:
                    break
                file.write("| ** " + unpassName[index_])
                file.write(" ** | ** " + unpassPer[index_] + " ** |" + "\n")
                index_ += 1
        input_OK=True
    except Exception as error:
        print("Error:", error)
print("统计完成，结果已输出至：" + OUTPUT_DIR)