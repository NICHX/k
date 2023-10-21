from time import sleep

from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
import xlwt
import time
import os
import subprocess
import re


def replace_letter(text, replace_letter):
    pattern = re.compile(r'\b{}\b'.format(replace_letter))
    return pattern.sub(replace_letter, text)

# 先切换到chrome可执行文件的路径
os.chdir(r"C:\Program Files\Google\Chrome\Application")
# user-data-dir为路径
subprocess.Popen('chrome.exe --remote-debugging-port=9527 --user-data-dir="D:\project\kaoshibao\AutomationProfile"')

chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9527")
driver = webdriver.Chrome(options=chrome_options)

# driver = webdriver.Chrome()  # 谷歌浏览器
driver.get(
    'https://www.zaixiankaoshi.com/online/?paperId=11522202&practice=&modal=1&is_recite=&qtype=&text=%E9%A1%BA%E5%BA'
    '%8F%E7%BB%83%E4%B9%A0&sequence=0&is_collect=1&is_vip_paper=0')
driver.implicitly_wait(1)
driver.find_element(By.XPATH, '//*[@id="body"]/div[2]/div[1]/div[2]/div[2]/div[2]/div[1]/p[2]/span[2]/div').click()

for i in range(1259):
    time.sleep(1)
    all = []
    # 定位元素并提取内容
    answer = driver.find_element(By.XPATH,
                                 '//*[@id="body"]/div[2]/div[1]/div[2]/div[1]/div/div[3]/div[1]/div/div[1]/div['
                                 '1]/b/span').text
    title = driver.find_element(By.XPATH,
                                '// *[@id ="body"]/div[2]/div[1]/div[2]/div[1]/div/div[1]/div/div[1]/div/div').text
    part = driver.find_element(By.XPATH,
                                '//*[@id="body"]/div[2]/div[1]/div[2]/div[1]/div/div[1]/div/div[2]/div').text.replace("\n", ".")
    replace_letter(part, 'A.',)
    replace_letter(part, '\nB.')
    replace_letter(part, '\nC.')
    replace_letter(part, '\nD.')
    replace_letter(part, '\nE.')
    replace_letter(part, '\nF.')
    part = part.replace('.B.', '\nB.', 1)
    part = part.replace('.C.', '\nC.', 1)
    part = part.replace('.D.', '\nD.', 1)
    part = part.replace('.E.', '\nE.', 1)
    part = part.replace('.F.', '\nF.', 1)
    analysis = driver.find_element(By.XPATH,
                                   '//*[@id="body"]/div[2]/div[1]/div[2]/div[1]/div/div[3]/div[2]/div/div[1]').text
    # 对内容个性化处理
    title = str(i + 1) + "." + title
    answer = "参考答案：" + answer
    analysis = "解析：" + analysis
    all.append(title)
    all.append(part)
    all.append(answer)
    all.append(analysis)
    ques = title + ' \n' + part + '\n' + answer + ' \n' + analysis + '\n '
    print(ques)
    with open(r"D:\project\kaoshibao\2.txt", "a") as f:
        f.write(ques)  # 自带文件关闭功能，不需要再写f.close()
    # 第1条数据 最大化窗口
    if i == 0:
        driver.maximize_window()
        time.sleep(1)
    # 点击下一条
    driver.find_element(By.CLASS_NAME, 'el-button--primary').click()

# 存储表格
# 退出浏览器
driver.quit()
