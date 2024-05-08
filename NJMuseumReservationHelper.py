# 抢票自动化脚本，用于在指定时间抢购南京博物院的门票
# 注意：需要提前安装好Edge浏览器和对应的webdriver

import datetime
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import win32com.client
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 请根据需要修改以下信息：
# 1. 抢票时间
# 2. 日期对应的XPath
# 3. 注意最后一个按键（结算按钮）的代码被注释掉了，需要根据实际情况进行修改

# 初始化语音合成引擎，用于最后的语音提示
speaker = win32com.client.Dispatch("SAPI.SpVoice")
# 设置秒杀时间，格式为'YYYY-MM-DD HH:MM:SS'
times = '2024-04-13 18:25:00'

# 配置Edge浏览器选项，去除自动化痕迹
options = webdriver.EdgeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")

# 创建Edge浏览器驱动实例
browser = webdriver.Edge(options=options)

# 打开南京博物院预约登录页面
browser.maximize_window()
browser.get("https://ticket.wisdommuseum.cn/reservation/userOut/outSingle/toSingleIndex.do")
time.sleep(10)

# 用于标记是否已经成功买到票
is_ticket_bought = False

while True:
    # 获取当前时间，并格式化
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
    print(now)

    # 判断当前时间是否已经超过设定的秒杀时间且尚未购票成功
    if now > times and not is_ticket_bought:
        # 不断尝试直到购票成功或无法操作
        while True:
            try:
                browser.refresh()
                # 切换到iframe中进行操作
                browser.switch_to.frame(browser.find_element(By.ID, "Conframe"))

                # 使用WebDriverWait等待页面元素加载
                wait = WebDriverWait(browser, 1)
                button = wait.until(EC.presence_of_element_located((By.XPATH,
                                                                    '//*[@id="layui-laydate1"]/div/div[2]/table/tbody/tr[3]/td[7]/span')))
                button.click()  # 选择日期
                print("找到按钮啦~")
                now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
                print(now)

                # 添加参观人员信息
                wait2 = WebDriverWait(browser, 1)
                button2 = wait2.until(EC.presence_of_element_located((By.XPATH, '//*[@id="tr24559787"]/td[1]/input')))
                button2.click()
                button3 = wait2.until(EC.presence_of_element_located((By.XPATH, '//*[@id="tr24578066"]/td[1]/input')))
                button3.click()
                print("添加完成人员信息啦~")
                now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
                print(now)

                # 点击结算按钮
                wait3 = WebDriverWait(browser, 1)
                button4 = wait2.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ss"]')))
                button4.click()
                browser.find_element(By.XPATH, '//*[@id="layui-layer1"]/div[3]/a[1]').click()  # 确认订单
                print(f"抢到票啦啦啦！")
                now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
                print(now)

                # 使用语音合成引擎播报抢票结果
                speaker.Speak(f"抢到票啦")
                is_ticket_bought = True
                break
            except:
                print("找不到！")
                pass

        if is_ticket_bought:
            print("轻松拿下！")
            time.sleep(20000)  # 保持页面停留一段时间，以便查看订单

    time.sleep(0.000001)  # 循环间隔，调整此值可改变轮询频率
