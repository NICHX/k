#!/usr/bin/python
# -*- coding: utf-8 -*-
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import os
import subprocess
import re
import io
import sys
from gooey import Gooey, GooeyParser

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='gb18030')  # 改变标准输出的默认编码


@Gooey(language='chinese', program_name=u'kaoshibao', required_cols=2, optional_cols=2,
       advanced=True, clear_before_run=True, sidebar_title='工具列表', terminal_font_family='Courier New',
       menu=[{
           'name': '关于',
           'items': [{
               'type': 'AboutDialog',
               'menuTitle': '关于',
               'name': 'kaoshibao',
               'description': 'Created by NICHX !',
               'version': '0.0.1',
           }]
       }])
def main_window():
    parser = GooeyParser(description="Created by NICHX !")
    subs = parser.add_subparsers(help='commands', dest='command')
    ticket_parser = subs.add_parser('kaoshibao', help='kaoshibao题库')
    subgroup = ticket_parser.add_argument_group('配置')
    subgroup.add_argument('谷歌浏览器安装位置', default="C:\Program Files\Google\Chrome\Application",
                          help="谷歌浏览器安装位置")
    subgroup.add_argument('题库地址', help="请收藏题库后打开顺序练习复制地址", widget='TextField')
    subgroup.add_argument('题目数量', help="输入题库题目数量")
    subgroup.add_argument('保存目录', help="请选择想要保存到的目录", widget='DirChooser')
    subgroup.add_argument('保存文件名', help="保存文件名,无需后缀", widget='TextField')

    args = parser.parse_args()
    if args.command == 'kaoshibao':
        download_ques(args.谷歌浏览器安装位置, args.题目数量, args.题库地址, args.保存目录, args.保存文件名)


def replace_letter(text, replace_letter):
    pattern = re.compile(r'\b{}\b'.format(replace_letter))
    return pattern.sub(replace_letter, text)


def download_ques(谷歌浏览器安装位置, 题目数量, 题库地址, 保存目录, 保存文件名):
    # 先切换到chrome可执行文件的路径
    os.chdir(谷歌浏览器安装位置)
    # user-data-dir为路径
    subprocess.Popen('chrome.exe --remote-debugging-port=9222 --user-data-dir="D:\project\kaoshibao\AutomationProfile"')
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver = webdriver.Chrome(options=chrome_options)
    driver.get(str(题库地址))
    driver.implicitly_wait(3)
    driver.find_element(By.XPATH, '//*[@id="body"]/div[2]/div[1]/div[2]/div[2]/div[2]/div[1]/p[2]/span[2]/div').click()
    for i in range(int(题目数量)):
        time.sleep(1)
        all = []
        # 定位元素并提取内容
        answer = driver.find_element(By.XPATH,
                                     '//*[@id="body"]/div[2]/div[1]/div[2]/div[1]/div/div[3]/div[1]/div/div[1]/div[1]/b/span').text
        title = driver.find_element(By.XPATH,
                                    '// *[@id ="body"]/div[2]/div[1]/div[2]/div[1]/div/div[1]/div/div[1]/div/div').text
        part = driver.find_element(By.XPATH,
                                   '//*[@id="body"]/div[2]/div[1]/div[2]/div[1]/div/div[1]/div/div[2]/div').text.replace(
            "\n", ".")
        replace_letter(part, ' A.', )
        replace_letter(part, '\nB.')
        replace_letter(part, '\nC.')
        replace_letter(part, '\nD.')
        replace_letter(part, '\nE.')
        replace_letter(part, '\nF.')
        part = part.replace('.B.', ' B.', 1)
        part = part.replace('.C.', ' C.', 1)
        part = part.replace('.D.', ' D.', 1)
        part = part.replace('.E.', ' E.', 1)
        part = part.replace('.F.', ' F.', 1)
        analysis = driver.find_element(By.XPATH,
                                       '//*[@id="body"]/div[2]/div[1]/div[2]/div[1]/div/div[3]/div[2]/div/div[1]').text
        # 对内容个性化处理
        title = str(i + 1).lstrip() + "." + title
        answer = "参考答案：" + answer
        analysis = "解析：" + analysis
        all.append(title)
        all.append(part)
        all.append(answer)
        all.append(analysis)
        ques = title.replace("\n", "") + ' ' + part.replace("\n", " ") + ' ' + answer.replace("\n",
                                                                                              "") + ' ' + analysis.replace(
            "\n", " ") + '\n'
        ques = ques.encode('gb18030')
        ques1 = ques.decode('gb18030')
        print(ques1, flush=True)
        with open(保存目录 + '/' + 保存文件名 + '.txt', "a", encoding='utf8') as f:
            f.write(ques1)  # 自带文件关闭功能，不需要再写f.close()
        # 第1条数据 最大化窗口
        if i == 0:
            driver.maximize_window()
            time.sleep(1)
        # 点击下一条
        driver.find_element(By.XPATH,
                            '//*[@id ="body"]/div[2]/div[1]/div[2]/div[1]/div/div[1]/div/div[3]/button[2]').click()
    # 存储表格
    # 退出浏览器
    driver.quit()
    os.startfile(保存目录 + '/' + 保存文件名 + '.txt')


if __name__ == '__main__':
    main_window()
