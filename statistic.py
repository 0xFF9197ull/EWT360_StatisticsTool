from openpyxl import load_workbook
import time
from openpyxl.utils.exceptions import InvalidFileException

COMPLETE = 100
PASS = 60


def getData(cellRange, sheet):  # 接受参数：表格范围，所操作的工作表
    Data = []
    for location in cellRange:
        cellLocation = str(location)[20:][0:-3]
        Data.append(sheet[cellLocation].value)
    return Data


def checkOneDay():  # 单日
    unpassName = []
    incompleteName = []
    unpassPer = []
    incompletePer = []
    # 准备工作表
    input_OK = False
    # 尝试导入，出错后重新获取路径
    while input_OK == False:
        try:
            INPUT_DIR = input("输入表格路径")
            inputWorkbook = load_workbook(INPUT_DIR)
            input_OK = True

        except Exception as error:
            print("Error:", error)
    sheet = inputWorkbook.active  # 默认第一个工作表
    sheet.title = "MAIN SHEET"  # 改名，使后续的单元格地址长度固定
    #######以上为准备部分##############
    NOS = int(input("班级人数,按表格中人名总数填写")) + 3  # NumberOfStudents,班级人数
    NOS = str(NOS)
    cellRange_Name = sheet["C3":"C" + NOS]
    cellRange_completePercentage = sheet["H3":"H" + NOS]
    #######以上为获取单元格位置#########
    xlsxTime = sheet["A1"].value  # 获取表格内的截止时间
    print(xlsxTime)
    name = getData(cellRange_Name,sheet)
    completePercentage = getData(cellRange_completePercentage,sheet)
    #######以上为获取数据部分###########
    step = 0
    for per in completePercentage :
        if per == None:
            break
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
    #######以上为统计部分##############
    formatted_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())#获取当前系统时间
    input_OK = False
    while input_OK == False:
        try:
            OUTPUT_DIR = input("输出结果路径")
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
            input_OK = True
        except Exception as error:
            print("Error:", error)
    print("统计完成，结果已输出至：" + OUTPUT_DIR)
def checkXlsx():  # 多日
    pass


def main():  # 主函数
    pass
if __name__ == '__main__':
    checkOneDay()