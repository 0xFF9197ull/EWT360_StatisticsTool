from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from os import getcwd
import time

# 填写关键内容
# 账号，密码，你要统计的地址

userName = ""
userPassword = ""
url = ""
limit = 60  # 设置阈值，若完成度小于该值视为非今日任务，不参与统计
'''
##############################
##########初始化部分############
##############################
'''
# 创建 Chrome WebDriver 实例
service = webdriver.ChromeService(executable_path="/usr/bin/chromedriver")
browser = webdriver.Chrome(service=service)
# 打开ewt
browser.get(url)
browser.implicitly_wait(0.5)

OUTPUT_DIR = getcwd()


# 自动登录ewt
def login(userName, userPassword):
    login_userName = browser.find_element(by=By.ID, value="login__password_userName")
    login_Password = browser.find_element(by=By.ID, value="login__password_password")
    loginButton = browser.find_element(by=By.CLASS_NAME, value="ant-btn")
    login_userName.send_keys(userName)
    login_Password.send_keys(userPassword)
    loginButton.click()
    try:
        closeButton = WebDriverWait(browser, timeout=10).until(EC.presence_of_element_located(
            (By.XPATH, """//*[@id="driver-popover-item"]/div[4]/button""")))
        closeButton.click()
    except Exception:
        pass


# 按下"向前"按钮，将位置定位到第一天
def beforeButton():
    before = browser.find_element(by=By.XPATH, value="""//*[@id="rc-tabs-0-panel-1"]/div/div[4]/div/div/i[1]""")
    for k in range(10):
        before.click()


# 寻找“按资源看”按钮
def resourceButton():
    button_resource = browser.find_element(by=By.ID, value="rc-tabs-0-tab-2")
    button_resource.click()

    # 寻找"总体"
    # button_total = browser.find_element(by=By.XPATH,
    #                                   value="""//*[@id="rc-tabs-0-panel-2"]/div/div[1]/div[1]/label[1]""")
    # button_total.click()


# 获取每一天的完成度数据
def getData(type=1):  # 1为按学生看，2为按资源看
    DataList = []
    div_index = 0
    if type == 1:
        XPATH1 = """//*[@id="rc-tabs-0-panel-1"]/div/div[4]/div/div/div/div/div["""
    elif type == 2:
        XPATH1 = """//*[@id="rc-tabs-0-panel-2"]/div/div[3]/div/div/div/div/div["""
    while True:
        div_index += 1
        XPATH = XPATH1 + str(div_index) + """]/div[1]"""
        try:
            data = browser.find_element(by=By.XPATH,
                                        value=XPATH)
            if float(data.get_attribute("innerText")[5:-1]) >= limit:
                DataList.append(data.get_attribute("innerText")[5:-1])
        except Exception:
            return DataList


# 数据统计，如平均，最大值，最小值
def statistic(data):
    length = len(data)
    avr = 0.00
    _max, _min = max(data), min(data)
    for k in data:
        avr += float(k) / length
    return avr, _max, _min


# 选择学科
def selectSubject(subject=1):
    selectBox = browser.find_element(by=By.XPATH, value="""//*[@id="selectSubject"]/div""")
    sub = """/html/body/div[3]/div/div/div/div[2]/div/div/div[""" + str(subject) + "]"
    selectBox.click()
    Sub = browser.find_element(by=By.XPATH, value=sub)
    nowSubject = Sub.get_attribute("innerText")
    Sub.click()
    print(nowSubject)
    return nowSubject


def saveMeta(Data,Subject, fileName="meta.md"):
    formatted_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())  # 获取当前系统时间
    DIR = OUTPUT_DIR + "/" + fileName
    with open(DIR, "w") as file:
        index_ = 0
        file.write("# " + "生成时间：" + formatted_time + "\n\n")
        file.write("-------" + "\n\n")
        for a in Data:
            file.write("### 当前科目："+Subject[index_]+"\n\n")
            file.write("> ")
            for k in a:
                file.write(k)
            index_ +=1



    print("统计完成，结果已输出至：" + DIR)


def saveStatistic(Data, fileName="统计数据.md"):  # 保存数据到项目文件夹下
    formatted_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())  # 获取当前系统时间
    DIR = OUTPUT_DIR + "/" + fileName
    with open(DIR, "w") as file:
        index_ = 0
        file.write("# " + "生成时间：" + formatted_time + "\n\n")
        file.write("-------" + "\n\n")
        file.write("|科目|完成率|最低|最高|" + "\n")
        file.write("|:-----:|:-----:|:-----:|:-----:|" + "\n")
        for a in Data:
            file.write("|" + str(a[0]))
            file.write("|" + str(a[1])[:5])
            file.write("|" + str(a[2]))
            file.write("|" + str(a[3]) + "|" + "\n")
            index_ += 1
    print("统计完成，结果已输出至：" + DIR)

def debugger():
    while True:
        try:
            re = exec(input("Console>"))
            print(re)
        except Exception as error:
            print(error)


#########################
##########主程序部分#######
#########################

# 登陆
if input("自动登录(y/n)") == "y":
    login(userName, userPassword)
else:
    input("请手动登陆后回车")

resourceButton()
# 循环查看各学科完成度

totalData = []
metaData = []
subjectName=[]
for k in range(3, 9):
    sub = selectSubject(k)
    subjectName.append(sub)
    Data = getData(2)
    print(Data)
    avr, Max, Min = statistic(Data)
    print("平均:", avr, "\n[", Min, ",", Max, "]")
    metaData.extend([[Data]])
    totalData.extend([[sub, avr, Min, Max]])
saveStatistic(totalData)
browser.quit()
