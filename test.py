#!/usr/bin/python
# -*- coding: utf-8 -*-
from DrissionPage import ChromiumPage
from gooey import Gooey, GooeyParser
import sys, codecs, os

if sys.stdout.encoding != 'UTF-8':
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
if sys.stderr.encoding != 'UTF- 8':
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')


@Gooey(language='chinese', program_name=u'kaoshibao', required_cols=2, optional_cols=2,
       advanced=True, clear_before_run=True, sidebar_title='工具列表', terminal_font_family='Courier New',
       menu=[{
           'name': '关于',
           'items': [{
               'type': 'AboutDialog',
               'menuTitle': '关于',
               'name': 'kaoshibao',
               'description': 'Created by NICHX !',
               'version': '0.2.0',
           }]
       }])
def main_window():
    parser = GooeyParser(description="Created by NICHX !  该程序免费共享，请勿付费购买！\n作者邮箱：nichx@nichx.cn")
    subs = parser.add_subparsers(help='commands', dest='command')
    ticket_parser = subs.add_parser('kaoshibao', help='kaoshibao题库')
    subgroup = ticket_parser.add_argument_group('配置')
    subgroup.add_argument('--考试宝帐号', help="可选项，公开题库不需要登录")
    subgroup.add_argument('--考试宝密码', widget='PasswordField', help="可选项，公开题库不需要登录")
    subgroup.add_argument('题库ID', help="请输入题库ID", widget='TextField')
    subgroup.add_argument('题目数量', help="输入要爬取的题目数量")
    subgroup.add_argument('保存目录', help="请选择想要保存到的目录", widget='DirChooser')
    subgroup.add_argument('保存文件名', help="保存文件名,无需后缀", widget='TextField')

    args = parser.parse_args()
    if args.command == 'kaoshibao':
        download_ques(args.考试宝帐号, args.考试宝密码, args.题目数量, args.题库ID, args.保存目录, args.保存文件名)


def download_ques(考试宝帐号, 考试宝密码, 题目数量, 题库ID, 保存目录, 保存文件名):
    page = ChromiumPage()
    if 考试宝帐号 and 考试宝密码:
        # 跳转到登录页面
        page.get('https://www.kaoshibao.com/login/')
        # 定位到账号文本框，获取文本框元素
        ele = page.ele('@placeholder=请输入您的11位手机号码')
        # 输入对文本框输入账号
        ele.input(考试宝帐号)
        # 定位到密码文本框并输入密码
        page.ele('@placeholder=请输入您的密码').input(考试宝密码)
        # 点击登录按钮
        page.ele('立即登录').click()
        page.wait.load_start()


    url = 'https://www.kaoshibao.com/online/?paperId=' + 题库ID
    page.get(url)
    # 打开背题模式
    button = page.ele('背题模式').after('@@role=switch@@class=el-switch')
    if button:
        button.click()
    for i in range(int(题目数量)):
        题型 = page.ele('@class=topic-type').text
        if 题型 == '单选题':
            option = page.s_ele('@class=select-left pull-left options-w').text
            answer = page.s_ele('正确答案').text.replace('\u2003', ':')
        elif 题型 == '多选题':
            option = page.s_ele('@class=select-left pull-left options-w check-box').text
            answer = page.s_ele('正确答案').text.replace('\u2003', ':')
        elif 题型 == '判断题':
            option = page.s_ele('@class=select-left pull-left options-w').text
            answer = page.s_ele('正确答案').text.replace('\u2003', ':')
        elif 题型 == '填空题':
            option = ''
            answer = '正确答案:' + page.s_ele('@class=mt20').text.replace('\u2003', ':')
        elif 题型 == '简答题':
            option = ''
            answer = '正确答案:' + page.s_ele('@class=mt20').text.replace('\u2003', ':')

        title = str(i + 1).lstrip() + "." + page.ele('@class=qusetion-box').text
        formatted_option = "\n".join(
            f"{line[0]}. {line[1:]}" if line[0].isupper() else line for line in option.splitlines())
        '''answer = page.ele('正确答案').text.replace('\u2003', ':')'''
        analysis = page.s_ele('@class=answer-analysis')
        if analysis:
            analysis = analysis.text
        if not analysis:
            analysis = '暂无解析'
        ques = f'{title} \n {formatted_option}\n {answer} \n解析： {analysis} \n'.replace('\n\n', '\n')
        next_ques = page.ele('@class=el-button el-button--primary el-button--small')
        if next_ques:
            next_ques.click()
            page.wait.load_start()
        info = f'第{i + 1}题已完成'
        print(info.encode('gb18030').decode('gb18030'), flush=True)
        filepath = 保存目录 + '/' + 保存文件名 + '.txt'
        with open(filepath, "a", encoding='utf8') as f:
            f.write(ques)  # 自带文件关闭功能，不需要再写f.close()

    os.startfile(filepath)


if __name__ == '__main__':
    main_window()
