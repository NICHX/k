# -*- coding: utf-8 -*-
import codecs
import os
import sys

from DrissionPage import ChromiumPage, SessionPage
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
if sys.stderr.encoding != 'UTF-8':
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

version = '2.1.6'


@Gooey(language='chinese', program_name=u'考试宝下载工具', required_cols=2, optional_cols=2,
       advanced=True, clear_before_run=True, sidebar_title='工具列表', terminal_font_family='Courier New',
       menu=[{
           'name': '关于',
           'items': [{
               'type': 'AboutDialog',
               'menuTitle': '关于',
               'name': '考试宝下载工具\n',
               'description': 'Created by NICHX !\n 1、修改了登陆逻辑 \n 2、修复部分报错 \n 3、增加自定义时间间隔，避免因网络波动导致题目重复',
               'version': version,
           }]
       }])
def main_window():
    parser = GooeyParser(
        description="Created by NICHX !  该程序免费共享，请勿付费购买！\n安装谷歌Chrome浏览器！")
    subs = parser.add_subparsers(help='考试宝下载工具', dest='command')
    normal_parser = subs.add_parser('考试宝', help='kaoshibao题库')
    subgroup = normal_parser.add_argument_group('考试宝')
    '''subgroup.add_argument('考试宝帐号', help="必填")
    subgroup.add_argument('考试宝密码', widget='PasswordField', help="必填")'''
    subgroup.add_argument('题库ID', help="请输入题库ID", widget='TextField')
    subgroup.add_argument('保存目录', help="请选择想要保存到的目录", widget='DirChooser')
    subgroup.add_argument('延迟时间', help="爬取延迟时间，默认0.4，若有题目重复手动调高", widget='TextField', default='0.4')

    args = parser.parse_args()

    if args.command == '考试宝':
        download_ques(args.题库ID, args.保存目录, args.延迟时间)


def main():
    page = SessionPage()
    # 访问网页
    page.get('https://space.nichx.cn/Version.txt')
    remote_version = page.ele('text:version').text[10:]
    if remote_version == version:
        print(f'当前版本为{version} , 是最新版本')
        main_window()
    else:
        print(f'当前版本为{version} , 最新版本为{remote_version} , 请到 https://share.nichx.cn//s/kaoshibao 下载最新版本')
        input('Press Enter to exit...')


def download_ques(ID, path, time):
    page = ChromiumPage()
    '''    page.get('https://www.zaixiankaoshi.com/login/')
    # 定位到账号文本框，获取文本框元素
    ele = page.ele('@placeholder=请输入您的11位手机号码')
    # 输入对文本框输入账号
    ele.input(telephone)
    # 定位到密码文本框并输入密码
    page.ele('@placeholder=请输入您的密码').input(password)
    # 点击登录按钮
    page.ele('立即登录').click()
    page.wait.load_start()'''

    url = f'https://www.zaixiankaoshi.com/online/?paperId={ID}'
    page.get(url)

    doc = Document()
    doc.styles['Normal'].font.name = u'宋体'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    doc.styles['Normal'].font.size = Pt(11)
    page.wait.eles_loaded('xpath://*[@id="body"]/div[2]/div[1]/div[2]/div[1]/div/div[1]/div/div[1]/div/span[2]')
    number = page.ele('xpath://*[@id="body"]/div[2]/div[1]/div[2]/div[1]/div/div[1]/div/div[1]/div/span[2]').text[2:-1]
    # 打开背题模式
    try:
        button_off = page.s_ele('@@role=switch@@class=el-switch')
        if button_off:
            page.ele('xpath://*[@id="body"]/div[2]/div[1]/div[2]/div[2]/div[2]/div[1]/p[2]/span[2]/div/input').click()
            print('点击背题模式按钮')
            page.wait(0.3, 1.0)
    except ElementNotFoundError:
        print('背题模式已打开')
        page.wait(0.3, 0.9)
    for i in range(int(number)):
        title = f"{i + 1}. {page.ele('@class=qusetion-box').text}"
        doc.add_paragraph(title)
        try:
            ques_img = page.s_ele('@class=qusetion-box').ele('tag:img')
            if ques_img.link:
                ques_img_url = ques_img.attr('src')
                ques_img_url = f'{ques_img_url}'
                page.download(ques_img_url, rf'.\imgs\{ID}\ques', rename=f'ques{i + 1}-title.png')
                page.wait(0.3, 0.6)
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
                    option_img_url = j.s_ele('tag:img').link
                    # 定位当前选项内的类名以'before-icon'开头的元素
                    x = j.s_ele('@class^before-icon')  # 更改此处，确保使用当前选项的'before-icon'元素
                    # 下载选项图片到指定目录，并重命名
                    page.download(option_img_url, rf'.\imgs\{ID}\option', rename=f'ques{i + 1}-option-{x.text}.png')
                    page.wait(0.3, 0.6)
                    para = doc.add_paragraph()
                    list_j = list(j.text)
                    list_j.insert(1, '.')
                    str_j = ''.join(list_j)
                    run = para.add_run(str_j)  # 添加选项文本
                    img_path = rf'.\imgs\{ID}\option\ques{i + 1}-option-{x.text}.png'
                    run.add_picture(img_path, width=Inches(2.5))
                except Exception as e:
                    list_j = list(j.text)
                    list_j.insert(1, '.')
                    str_j = ''.join(list_j)
                    doc.add_paragraph(str_j)
                    option += str_j + "\n"
            answer = page.ele('@class=right-ans').text.replace('\u2003', ':')
        elif topic == '判断题':
            options = page.ele('@class^select-left').children('@class^option')
            for j in options:
                list_j = list(j.text)
                list_j.insert(1, '.')
                str_j = ''.join(list_j)
                doc.add_paragraph(str_j)
                option += str_j + "\n"
            answer = page.ele(
                'xpath://*[@id="body"]/div[2]/div[1]/div[2]/div[1]/div/div[3]/div[1]/div/div[1]/div/b').text.replace(
                '\u2003', ':')
        elif topic == '多选题':
            options = page.s_eles('@class^option')
            for j in options:
                try:
                    # 定位选项内的类名以'fr-fic '开头的元素，假设它代表选项图片
                    option_img_url = j.s_ele('tag:img').link
                    # 定位当前选项内的类名以'before-icon'开头的元素
                    x = j.s_ele('@class^before-icon')  # 更改此处，确保使用当前选项的'before-icon'元素
                    # 下载选项图片到指定目录，并重命名
                    page.download(option_img_url, rf'.\imgs\{ID}\option', rename=f'ques{i + 1}-option-{x.text}.png')
                    page.wait(0.3, 0.6)
                    para = doc.add_paragraph()
                    list_j = list(j.text)
                    list_j.insert(1, '.')
                    str_j = ''.join(list_j)
                    run = para.add_run(str_j)  # 添加选项文本
                    img_path = rf'.\imgs\{ID}\option\ques{i + 1}-option-{x.text}.png'
                    run.add_picture(img_path, width=Inches(2.5))
                except Exception as e:
                    list_j = list(j.text)
                    list_j.insert(1, '.')
                    str_j = ''.join(list_j)
                    doc.add_paragraph(str_j)
                    option += str_j + "\n"
            answer = page.ele(
                'xpath://*[@id="body"]/div[2]/div[1]/div[2]/div[1]/div/div[3]/div[1]/div/div[1]/div/b').text.replace(
                '\u2003', ':')
        elif topic == '填空题':
            answer = '正确答案:' + page.s_ele('@class=mt20').text.replace('\u2003', ':')
        elif topic == '简答题':
            answer = '正确答案:' + page.s_ele('@class=mt20').text.replace('\u2003', ':')

        '''formatted_option = "\n".join(
            f"{line[0]}. {line[1:]}" if line[0].isupper() else line for line in option.splitlines())'''

        try:
            analysis = page.s_ele('@class^answer-analysis').text.replace('\n', '')
            try:
                analysis_img = page.s_ele('@class^answer-analysis').ele('tag:img')
                if analysis_img.link:
                    analysis_img_url = analysis_img.attr('src')
                    if analysis_img_url == 'https://resource.zaixiankaoshi.com/mini/ai_tag.png':
                        pass
                    else:
                        analysis_img_url = f'{analysis_img_url}'
                        page.download(analysis_img_url, rf'.\imgs\{ID}\analysis', rename=f'ques{i + 1}-analysis.png')
                        page.wait(0.3, 0.6)
            except ElementNotFoundError:
                pass
        except Exception as e:
            print(e)
        if option != '':
            ques = f'{title}\n{option}\n{answer}\n解析：{analysis}\n'
        else:
            ques = f'{title}\n{answer}\n解析：{analysis}\n'
        try:
            page.ele('@@class:el-button el-button--primary el-button--small@@text():下一题', timeout=5).click()
            page.wait(float(time))
        except Exception as e:
            print(e)
        # 添加答案段落
        doc.add_paragraph(answer)
        doc.add_paragraph(f'解析：{analysis}')
        try:
            doc.add_picture(rf'.\imgs\{ID}\analysis\ques{i + 1}-analysis.png')
        except Exception as e:
            pass
        info = f'第{i + 1}题已完成'
        print(info, flush=True)
        filepath = f'{path}/{ID}.txt'
        with open(filepath, "a", encoding='utf8') as f:
            f.write(ques)  # 自带文件关闭功能，不需要再写f.close()
    doc.save(f'{path}/{ID}.docx')
    os.startfile(f'{path}/{ID}.docx')


if __name__ == '__main__':
    main()
