# -*- coding: utf-8 -*-
import codecs
import os
import sys

from DrissionPage import ChromiumPage
from DrissionPage.common import Settings
from DrissionPage.errors import ElementNotFoundError
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.shared import Pt
from gooey import Gooey, GooeyParser

Settings.raise_when_ele_not_found = True

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
               'version': '1.0.0',
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
    subgroup.add_argument('保存目录', help="请选择想要保存到的目录", widget='DirChooser')

    args = parser.parse_args()
    if args.command == 'kaoshibao':
        download_ques(args.考试宝帐号, args.考试宝密码, args.题库ID, args.保存目录)


def download_ques(telephone, password, ID, path):
    page = ChromiumPage()
    if telephone and password:
        # 跳转到登录页面
        page.get('https://www.kaoshibao.com/login/')
        # 定位到账号文本框，获取文本框元素
        ele = page.ele('@placeholder=请输入您的11位手机号码')
        # 输入对文本框输入账号
        ele.input(telephone)
        # 定位到密码文本框并输入密码
        page.ele('@placeholder=请输入您的密码').input(password)
        # 点击登录按钮
        page.ele('立即登录').click()
        page.wait.load_start()

    url = f'https://www.kaoshibao.com/online/?paperId={ID}'
    page.get(url)

    doc = Document()
    doc.styles['Normal'].font.name = u'宋体'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    doc.styles['Normal'].font.size = Pt(11)

    number = page.ele('@style=float: left; font-weight: 700;').text[2:-1]
    # 打开背题模式
    try:
        page.s_ele('背题模式').after('@@role=switch@@class=el-switch').click()
    except ElementNotFoundError:
        print('背题模式已打开')
    for i in range(int(number)):
        title = f"{i + 1}. {page.ele('@class=qusetion-box').text}"
        doc.add_paragraph(title)
        try:
            ques_img = page.s_ele('@class^fr-fic ')
            ques_img_url = ques_img.attr('src')
            ques_img_url = f'{ques_img_url}'
            page.download(ques_img_url, rf'.\imgs\{ID}\ques', rename=f'ques{i + 1}-title.png')
            doc.add_picture(rf'.\imgs\{ID}\ques\ques{i + 1}-title.png')
        except ElementNotFoundError:
            pass
        topic = page.ele('@class=topic-type').text
        option = ''
        if topic == '单选题':
            options = page.s_eles('@class^option')
            for j in options:
                try:
                    # 定位选项内的类名以'fr-fic '开头的元素，假设它代表选项图片
                    option_img = j.s_ele('@class^fr-fic ')
                    # 定位当前选项内的类名以'before-icon'开头的元素
                    x = j.s_ele('@class^before-icon')  # 更改此处，确保使用当前选项的'before-icon'元素
                    option_img_url = option_img.attr('src')
                    option_img_url = f'{option_img_url}'
                    # 下载选项图片到指定目录，并重命名
                    page.download(option_img_url, rf'.\imgs\{ID}\option', rename=f'ques{i + 1}-option-{x.text}.png')
                    para = doc.add_paragraph()
                    run = para.add_run(j.text)  # 添加选项文本
                    img_path = rf'.\imgs\{ID}\option\ques{i + 1}-option-{x.text}.png'
                    run.add_picture(img_path, width=Inches(2.5))
                except Exception as e:
                    text = j.text
                    doc.add_paragraph(text)
                    option += text + "\n"
            answer = page.s_ele('正确答案').text.replace('\u2003', ':')
        elif topic == '判断题':
            options = page.ele('@class^select-left').children('@class^option')
            for j in options:
                text = j.text
                doc.add_paragraph(text)
                option += text + "\n"
            answer = page.s_ele('正确答案').text.replace('\u2003', ':')
        elif topic == '多选题':
            options = page.s_eles('@class^option')
            for j in options:
                try:
                    # 定位选项内的类名以'fr-fic '开头的元素，假设它代表选项图片
                    option_img = j.s_ele('@class^fr-fic ')
                    # 定位当前选项内的类名以'before-icon'开头的元素
                    x = j.s_ele('@class^before-icon')  # 更改此处，确保使用当前选项的'before-icon'元素
                    option_img_url = option_img.attr('src')
                    option_img_url = f'{option_img_url}'
                    # 下载选项图片到指定目录，并重命名
                    page.download(option_img_url, rf'.\imgs\{ID}\option', rename=f'ques{i + 1}-option-{x.text}.png')
                    para = doc.add_paragraph()
                    run = para.add_run(j.text)  # 添加选项文本
                    img_path = rf'.\imgs\{ID}\option\ques{i + 1}-option-{x.text}.png'
                    run.add_picture(img_path, width=Inches(2.5))
                except Exception as e:
                    text = j.text
                    doc.add_paragraph(text)
                    option += text + "\n"
            answer = page.s_ele('正确答案').text.replace('\u2003', ':')
        elif topic == '填空题':
            answer = '正确答案:' + page.s_ele('@class=mt20').text.replace('\u2003', ':')
        elif topic == '简答题':
            answer = '正确答案:' + page.s_ele('@class=mt20').text.replace('\u2003', ':')

        formatted_option = "\n".join(
            f"{line[0]}. {line[1:]}" if line[0].isupper() else line for line in option.splitlines())

        try:
            analysis = page.s_ele('@class=answer-analysis-row').text.replace('\n', '')
        except ElementNotFoundError:
            analysis = page.s_ele('@class=answer-analysis').text.replace('\n', '')
        if option != '':
            ques = f'{title}\n{formatted_option}\n{answer}\n解析：{analysis}\n'
        else:
            ques = f'{title}\n{answer}\n解析：{analysis}\n'
        try:
            page.ele('@class=el-button el-button--primary el-button--small', timeout=0.5).click()
        except ElementNotFoundError:
            pass
        # 添加答案段落
        doc.add_paragraph(answer)

        # 添加解析段落
        doc.add_paragraph(f'解析：{analysis}')
        info = f'第{i + 1}题已完成'
        print(info, flush=True)
        filepath = f'{path}/{ID}.txt'
        with open(filepath, "a", encoding='utf8') as f:
            f.write(ques)  # 自带文件关闭功能，不需要再写f.close()
    doc.save(f'{path}/{ID}.docx')
    os.startfile(f'{path}/{ID}.docx')


if __name__ == '__main__':
    main_window()
