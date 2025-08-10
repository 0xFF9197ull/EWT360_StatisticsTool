from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from os import getcwd

# 填写关键内容
# 账号，密码，你要统计的地址
userName = "yourName"# 如果不填写，后面请选择手动登陆
userPassword = "yourPassword"
url = "yourURL"#必填
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


# 自动登录ewt
def login(userName, userPassword):
    login_userName = browser.find_element(by=By.ID, value="login__password_userName")
    login_Password = browser.find_element(by=By.ID, value="login__password_password")
    loginButton = browser.find_element(by=By.CLASS_NAME, value="ant-btn")
    login_userName.send_keys(userName)
    login_Password.send_keys(userPassword)
    loginButton.click()

    closeButton = WebDriverWait(browser, timeout=10).until(EC.presence_of_element_located(
        (By.XPATH, """//*[@id="driver-popover-item"]/div[4]/button""")))
    closeButton.click()


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
    button_total = browser.find_element(by=By.XPATH,
                                        value="""//*[@id="rc-tabs-0-panel-2"]/div/div[1]/div[1]/label[1]""")
    button_total.click()

#获取每一天的完成度数据
def getData():
    DataList=[]
    div_index = 0
    while True:
        div_index += 1
        XPATH = """//*[@id="rc-tabs-0-panel-1"]/div/div[4]/div/div/div/div/div[""" + str(div_index) + """]/div[1]"""
        try :
            data = browser.find_element(by=By.XPATH,
                    value=XPATH)
            DataList.append(data.get_attribute("innerText"))
        except Exception:
            return DataList

if input("自动登录(y/n)") == "y":
    login(userName, userPassword)
else:
    input("请手动登陆后回车")

## resourceButton()


Data=getData()
print(Data)


browser.quit()
