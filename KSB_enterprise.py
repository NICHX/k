# -*- coding: utf-8 -*-
import codecs
import os
import sys

import wmi
import xlwt
from DrissionPage import ChromiumPage, SessionPage
from DrissionPage.common import Settings
from DrissionPage.errors import ElementNotFoundError
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.shared import Pt
from gooey import Gooey, GooeyParser
import configparser


Settings.raise_when_ele_not_found = True

if sys.stdout.encoding != 'UTF-8':
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
if sys.stderr.encoding != 'UTF-8':
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')


def read_or_create_config(file_path='config.ini'):
    config = configparser.ConfigParser()
    # 尝试读取配置文件
    try:
        config.read(file_path)
    except Exception as e:
        print("Warning: File does not appear to be a valid .ini file.")
    # 如果需要写入默认配置，确保在读取后进行
    # 示例：添加一个默认section和option
    if not config.has_section('config'):
        config.add_section('config')
        config.set('config', 'qq', '')
        config.set('config', 'orderid', '')
        config.set('config', 'code', '')

    # 如果文件一开始不存在，下面的代码会在写入时创建文件
    with open(file_path, 'w') as configfile:
        config.write(configfile)
    _orderid = config.get("config", "orderid")
    _qq = config.get("config", "qq")
    _code = config.get("config", "code")
    return [_orderid, _qq, _code]


c = wmi.WMI()


def printMain_board():
    for board_id in c.Win32_BaseBoard():
        tmpmsg = {'SerialNumber': board_id.SerialNumber}
    return tmpmsg


def get_user(orderid, qq, code):
    page = SessionPage()
    page.get(f'https://read.nichx.cn/users/get_user_order/{orderid}?QQ={qq}&hashed_code={code}')
    user = page.json
    return user


def write_config(orderid, qq, code):
    cf = configparser.ConfigParser()
    cf.read('config.ini')  # 读取配置文件
    cf.set("config", "qq", qq)
    cf.set("config", "orderid", orderid)
    cf.set("config", "code", code)
    cf.write(open(r'config.ini', "w"))


def reg_device(orderid, qq, code):
    page = SessionPage()
    try:
        userid = get_user(orderid, qq, code)['userid']
        board_id = printMain_board()['SerialNumber']
        page.post(f'https://read.nichx.cn/users/reg_board/{userid}/?board_id={board_id}')
        print('设备注册成功', flush=True)
    except Exception as e:
        input('用户未注册或输入有误,请在qq群中注册并删除config.ini以重试（按Enter退出）')
        sys.exit(0)


def register_window():
    config = read_or_create_config(r'config.ini')

    if config[0] == '':
        try:
            print('KSB工具注册')
            orderid = input('请输入您的订单号：')
            qq = input('请输入您的qq号： ')
            口令 = input('请输入从@NICHX_bot获取的口令： ')
            remote_code = get_user(orderid, qq, 口令)['hashed_code']
            if 口令 == remote_code:
                print(f'正确口令为{remote_code} , 校验通过', flush=True)
                write_config(orderid, qq, 口令)
                reg_device(orderid, qq, 口令)
                KSB_window()
            else:
                input(f'口令错误,请重新获取或联系管理员(按Enter退出)')
                sys.exit(1)
        except Exception as e:
            input('用户未注册或配置有误(按Enter退出)')
            sys.exit(1)

    else:
        try:
            board_id = printMain_board()['SerialNumber']
            reg_device_id = get_user(config[0], config[1], config[2])['board_id']
            if board_id == reg_device_id:
                KSB_window()
            else:
                input('账号错误或已注册其他设备,更换设备请联系管理员（按Enter退出）')
                sys.exit(0)
        except Exception as e:
            input('用户未注册或配置有误,请在qq群中注册并删除config.ini以重试（按Enter退出）')
            sys.exit(1)


def check_suffix(filename, default_suffix='.png'):
    """
    检查文件名是否有扩展名，如果没有则添加默认扩展名，并在目标文件名不存在的情况下重命名文件。

    :param filename: 字符串，原始文件名（不含路径）
    :param default_suffix: 字符串，默认要添加的文件扩展名（带点，如'.png'）
    :return: 新文件名（已添加扩展名，且如果执行了重命名操作）
    """
    # 使用os.path.splitext分离文件名和扩展名
    name, ext = os.path.splitext(filename)

    # 如果没有扩展名，则添加默认扩展名
    if not ext:
        new_name = filename + default_suffix

        # 检查新文件名是否存在，如果不存在则尝试重命名
        if not os.path.exists(new_name):
            try:
                os.rename(filename, new_name)
                print(f"文件已成功重命名为: {new_name}")
            except OSError as e:
                print(f"重命名文件时发生错误: {e}")
                return None  # 或者你可以选择返回原文件名或其他逻辑
        else:
            print(f"目标文件 {new_name} 已存在，跳过重命名操作。")
    else:
        # 文件已有扩展名，直接返回原文件名
        new_name = filename

    return new_name


version = '1.1.2'


@Gooey(language='chinese', program_name=u'KSB工具(enterprise_version) beta', required_cols=2, optional_cols=2,
       enterprise=True, clear_before_run=True, sidebar_title='工具列表', terminal_font_family='Courier New',
       menu=[{
           'name': '关于',
           'items': [{
               'type': 'AboutDialog',
               'menuTitle': '关于',
               'name': 'KSB工具(enterprise_version) beta\n',
               'description': 'Created by NICHX !\n 1、可导出题库为TXT、Word、excel格式\n 2、新增支持论述题、排序题、不定项选择题。'
                              '\n 3、新增解析开关\n已知问题：部分题目的解析无法获取！',

               'version': version,
           }]
       }])
def KSB_window():
    config = read_or_create_config(r'config.ini')
    level = get_user(config[0], config[1], config[2])['level']
    if level == 'enterprise':
        print('你的使用权限为：企业版', flush=True)
    elif level == 'advanced':
        input('你的使用权限为：高级版。请使用高级版客户端（按Enter退出）')
        sys.exit(1)
    else:
        input('您没有企业版使用权限（按Enter退出）')
        sys.exit(1)

    parser = GooeyParser(
        description="安装谷歌Chrome浏览器！")
    subs = parser.add_subparsers(help='KSB', dest='command')
    normal_parser = subs.add_parser('KSB', help='KSB工具')
    subgroup = normal_parser.add_argument_group('配置信息')
    '''subgroup.add_argument('KSB帐号', help="必填")
    subgroup.add_argument('KSB密码', widget='PasswordField', help="必填")'''
    # subgroup.add_argument('口令', help="请输入qq群中获取的口令", widget='TextField', default=config[2] )
    subgroup.add_argument('题库ID', help="请输入题库ID", widget='TextField')
    subgroup.add_argument('解析开关', help="是否需要解析", widget='Dropdown'
                          , choices=['是', '否'], default='是')
    subgroup.add_argument('起始题号', help="从哪一题开始下载", widget='TextField'
                           , default='1')
    subgroup.add_argument('默认打开文件', help="请选择完成后想要打开的文件类型", widget='Dropdown'
                          , choices=['.txt', '.docx', '.xls', '不自动打开'], default='.docx')
    subgroup.add_argument('延迟时间', help="爬取延迟时间，默认0.4，若有题目重复手动调高", widget='TextField',
                          default='0.4')
    enterprise_parser = subs.add_parser('KSB企业版', help='KSB企业版')
    subgroup1 = enterprise_parser.add_argument_group('配置信息')
    subgroup1.add_argument('题库ID', help="请输入题库ID", widget='TextField')
    subgroup1.add_argument('解析开关', help="是否需要解析", widget='Dropdown'
                           , choices=['是', '否'], default='是')
    subgroup1.add_argument('起始题号', help="从哪一题开始下载", widget='TextField'
                           , default='1')
    subgroup1.add_argument('默认打开文件', help="请选择完成后想要打开的文件类型", widget='Dropdown'
                           , choices=['.txt', '.docx', '.xls', '不自动打开'], default='.docx')
    subgroup1.add_argument('延迟时间', help="爬取延迟时间，默认0.4，若有题目重复手动调高", widget='TextField',
                           default='0.4')

    args = parser.parse_args()

    if args.command == 'KSB':
        download_ques(args.题库ID, args.延迟时间, args.起始题号, args.默认打开文件, args.解析开关)
    if args.command == 'KSB企业版':
        download_ques_enterprise(args.题库ID, args.延迟时间, args.起始题号, args.默认打开文件, args.解析开关)


def start():
    page = SessionPage()
    # 访问网页
    page.get('https://space.nichx.cn/enterprise_version.txt')
    remote_version = page.ele('text:version').text[10:]
    if remote_version == version:
        print(f'当前版本为{version} , 是最新版本', flush=True)
        register_window()
    else:
        print(f'当前版本为{version} , 最新版本为{remote_version} , 请加入qq群：338283650 获取最新版本', flush=True)
        input('Press Enter to exit...')


def download_ques(ID, time, begin, file_format, anl_switch):
    try:
        os.mkdir(rf'.\{ID}')
        print('目录创建成功', flush=True)
    except FileExistsError:
        print('该题库已下载', flush=True)
        pass

    doc = Document()
    doc.styles['Normal'].font.name = u'宋体'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    doc.styles['Normal'].font.size = Pt(11)

    wb = xlwt.Workbook(encoding='utf-8')  # 新建一个excel文件
    ws1 = wb.add_sheet('sheet1')  # 添加一个新表，名字为begin
    ws1.write(0, 0, '序号')
    ws1.write(0, 1, '题目')
    ws1.write(0, 2, 'A')
    ws1.write(0, 3, 'B')
    ws1.write(0, 4, 'C')
    ws1.write(0, 5, 'D')
    ws1.write(0, 6, 'E')
    ws1.write(0, 7, 'F')
    ws1.write(0, 8, 'G')
    ws1.write(0, 9, 'H')
    ws1.write(0, 10, '正确答案')
    ws1.write(0, 11, '解析')

    url = f'https://www.zaixiankaoshi.com/online/?paperId={ID}'
    page = ChromiumPage()
    page.get(url)
    page.wait.eles_loaded('xpath://*[@id="body"]/div[2]/div[1]/div[2]/div[1]/div/div[1]/div/div[1]/div/span[2]')
    number = page.ele('xpath://*[@id="body"]/div[2]/div[1]/div[2]/div[1]/div/div[1]/div/div[1]/div/span[2]').text[2:-1]
    if int(begin) <= int(number):
        try:
            page.ele(f'tag:span@text():{begin}').click()
        except Exception as e:
            print(e)
    else:
        print('起始题号超出题库范围！')
        sys.exit(1)
    # 打开背题模式
    try:
        button_off = page.s_ele('@@role=switch@@class=el-switch')
        if button_off:
            page.ele('xpath://*[@id="body"]/div[2]/div[1]/div[2]/div[2]/div[2]/div[1]/p[2]/span[2]/div/input').click()
            print('点击背题模式按钮')
            page.wait(0.3, 0.6)
    except ElementNotFoundError:
        print('背题模式已打开')
        page.wait(0.3)
    for i in range(int(begin) - 1, int(number)):
        try:
            title = f"{page.ele('@class=qusetion-box').text}".replace('\n', '')
            doc.add_paragraph(f'{i + 1}.{title}')
            try:
                ques_img = page.s_ele('@class=qusetion-box').ele('tag:img')
                if ques_img.link:
                    ques_img_url = ques_img.attr('src')
                    # ques_img_url = f'{ques_img_url}'
                    ques_img = page.download(ques_img_url, rf'.\{ID}\imgs\ques', rename=f'ques{i + 1}-title',
                                             file_exists='skip')
                    page.wait(0.3)
                doc.add_picture(ques_img[1], width=Inches(3.5))
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
                        option_img = page.download(option_img_url, rf'.\{ID}\imgs\option',
                                                   rename=f'ques{i + 1}-option-{x.text}',
                                                   file_exists='skip')
                        page.wait(0.3)
                        para = doc.add_paragraph()
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        run = para.add_run(str_j)  # 添加选项文本
                        img_path = option_img[1]
                        run.add_picture(img_path, width=Inches(2.5))
                    except Exception as e:
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        doc.add_paragraph(str_j)
                        option += str_j + "\n"
                answer = page.s_ele('@class=right-ans').text.replace('\u2003', ':')
            elif topic == '判断题':
                options = page.ele('@class^select-left').children('@class^option')
                for j in options:
                    list_j = list(j.text)
                    list_j.insert(1, '.')
                    str_j = ''.join(list_j)
                    doc.add_paragraph(str_j)
                    option += str_j + "\n"
                answer = page.s_ele('@class=right-ans').text.replace('\u2003', ':')
            elif topic == '多选题':
                options = page.s_eles('@class^option')
                for j in options:
                    try:
                        # 定位选项内的类名以'fr-fic '开头的元素，假设它代表选项图片
                        option_img_url = j.s_ele('tag:img').link
                        # 定位当前选项内的类名以'before-icon'开头的元素
                        x = j.s_ele('@class^before-icon')  # 更改此处，确保使用当前选项的'before-icon'元素
                        # 下载选项图片到指定目录，并重命名
                        option_img = page.download(option_img_url, rf'.\{ID}\imgs\option',
                                                   rename=f'ques{i + 1}-option-{x.text}', file_exists='skip')
                        page.wait(0.3)
                        para = doc.add_paragraph()
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        run = para.add_run(str_j)  # 添加选项文本
                        img_path = option_img[1]
                        run.add_picture(img_path, width=Inches(2.5))
                    except Exception as e:
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        doc.add_paragraph(str_j)
                        option += str_j + "\n"
                answer = page.s_ele('@class=right-ans').text.replace('\u2003', ':')
            elif topic == '不定项选择题':
                options = page.s_eles('@class^option')
                for j in options:
                    try:
                        # 定位选项内的类名以'fr-fic '开头的元素，假设它代表选项图片
                        option_img_url = j.s_ele('tag:img').link
                        # 定位当前选项内的类名以'before-icon'开头的元素
                        x = j.s_ele('@class^before-icon')  # 更改此处，确保使用当前选项的'before-icon'元素
                        # 下载选项图片到指定目录，并重命名
                        option_img = page.download(option_img_url, rf'.\{ID}\imgs\option',
                                                   rename=f'ques{i + 1}-option-{x.text}', file_exists='skip')
                        page.wait(0.3)
                        para = doc.add_paragraph()
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        run = para.add_run(str_j)  # 添加选项文本
                        img_path = option_img[1]
                        run.add_picture(img_path, width=Inches(2.5))
                    except Exception as e:
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        doc.add_paragraph(str_j)
                        option += str_j + "\n"
                answer = page.s_ele('@class=right-ans').text.replace('\u2003', ':')
            elif topic == '排序题':
                options = page.s_eles('@class^option')
                for j in options:
                    try:
                        # 定位选项内的类名以'fr-fic '开头的元素，假设它代表选项图片
                        option_img_url = j.s_ele('tag:img').link
                        # 定位当前选项内的类名以'before-icon'开头的元素
                        x = j.s_ele('@class^before-icon')  # 更改此处，确保使用当前选项的'before-icon'元素
                        # 下载选项图片到指定目录，并重命名
                        option_img = page.download(option_img_url, rf'.\{ID}\imgs\option',
                                                   rename=f'ques{i + 1}-option-{x.text}', file_exists='skip')
                        page.wait(0.3)
                        para = doc.add_paragraph()
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        run = para.add_run(str_j)  # 添加选项文本
                        img_path = option_img[1]
                        run.add_picture(img_path, width=Inches(2.5))
                    except Exception as e:
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        doc.add_paragraph(str_j)
                        option += str_j + "\n"
                answer = page.s_ele('@class=right-ans').text.replace('\u2003', ':')
            elif topic == '填空题':
                answer = '正确答案:' + page.s_ele('@class=mt20').text.replace('\u2003', ':')
            elif topic == '简答题':
                answer = '正确答案:' + page.s_ele('@class=mt20').text.replace('\u2003', ':')
            elif topic == '论述题':
                answer = '正确答案:' + page.s_ele('@class=mt20').text.replace('\u2003', ':')
            analysis = ''
            if anl_switch == '是':
                try:
                    analysis = '解析：' + page.s_ele('@class^answer-analysis').text.replace('\n', '')
                    try:
                        analysis_img = page.s_ele('@class^answer-analysis').ele('tag:img')
                        if analysis_img.link:
                            analysis_img_url = analysis_img.attr('src')
                            if analysis_img_url == 'https://resource.zaixiankaoshi.com/mini/ai_tag.png':
                                pass
                            else:
                                analysis_img_url = f'{analysis_img_url}'
                                analysis_img = page.download(analysis_img_url, rf'.\{ID}\imgs\analysis',
                                                             rename=f'ques{i + 1}-analysis', file_exists='skip')
                                page.wait(0.3)
                                doc.add_picture(analysis_img[1], width=Inches(2.5))
                    except ElementNotFoundError:
                        pass
                except Exception as e:
                    print(e)
            if option != '':
                ques = f'{i + 1}.{title}\n{option}{answer}\n{analysis}\n\n'
                option1 = option.replace('\n', '&@')
                ques1 = f'{i + 1}&@{title}&@{option1}&@{answer[5:]}&@{analysis}\n'
            else:
                ques = f'{i + 1}.{title}\n{option}{answer}\n{analysis}\n\n'
                ques1 = f'{i + 1}&@{title}&@{answer[5:]}&@{analysis}\n'
            # 添加答案段落
            doc.add_paragraph(answer)
            try:
                answer_img = page.s_ele('@class^mt20', timeout=0.3).ele('tag:img', timeout=0.3)
                if answer_img.link:
                    answer_img_url = answer_img.attr('src')
                    # ques_img_url = f'{ques_img_url}'
                    answer_img = page.download(answer_img_url, rf'.\{ID}\imgs\answer', rename=f'ques{i + 1}-answer',
                                               file_exists='skip')
                doc.add_picture(check_suffix(answer_img[1]), width=Inches(2.5))
            except ElementNotFoundError:
                pass
            doc.add_paragraph(f'{analysis} \n')
            try:
                doc.add_picture(check_suffix(analysis_img[1]), width=Inches(2.5))
            except Exception as e:
                pass
            list_a = ques1.split('&@')
            while len(list_a) <= 4:
                list_a.insert(2, '')
            while 4 < len(list_a) < 12:
                list_a.insert(-3, '')
            try:
                ws1.write(i + 1, 0, int(list_a[0]))
                ws1.write(i + 1, 1, list_a[1])
                ws1.write(i + 1, 2, list_a[2])
                ws1.write(i + 1, 3, list_a[3])
                ws1.write(i + 1, 4, list_a[4])
                ws1.write(i + 1, 5, list_a[5])
                ws1.write(i + 1, 6, list_a[6])
                ws1.write(i + 1, 7, list_a[7])
                ws1.write(i + 1, 8, list_a[8])
                ws1.write(i + 1, 9, list_a[9])
                ws1.write(i + 1, 10, list_a[-2])
                ws1.write(i + 1, 11, list_a[-1][3:])
            except IndexError as e:
                pass
            wb.save(rf'.\{ID}\{ID}-第{begin}题开始.xls')
            info = f'第{i + 1}题已完成'
            try:
                print(ques, flush=True)
            except Exception as e:
                print(e)
                print(info, flush=True)
            filepath = rf'.\{ID}\{ID}-第{begin}题开始.txt'
            with open(filepath, "a", encoding='utf8') as f:
                f.write(ques)  # 自带文件关闭功能，不需要再写f.close()
            doc.save(rf'.\{ID}\{ID}-第{begin}题开始.docx')
            try:
                page.ele('@@class:el-button el-button--primary el-button--small@@text():下一题', timeout=5).click()
                page.wait(float(time))
            except Exception as e:
                print(e)
        except ElementNotFoundError:
            print(f'第{i + 1}题下载失败\n', flush=True)
            with open(f'{ID}_error_log.txt', "a", encoding='utf8') as f:
                f.write(f'第{i + 1}题下载失败\n')  # 自带文件关闭功能，不需要再写f.close()
            try:
                page.ele('@@class:el-button el-button--primary el-button--small@@text():下一题', timeout=5).click()
                page.wait(float(time))
            except Exception as e:
                print(e)
        continue
    if file_format == '.txt':
        os.startfile(rf'.\{ID}\{ID}-第{begin}题开始.txt')
    elif file_format == '.docx':
        os.startfile(rf'.\{ID}\{ID}-第{begin}题开始.docx')
    elif file_format == '.xls':
        os.startfile(rf'.\{ID}\{ID}-第{begin}题开始.xls')
    elif file_format == '不自动打开':
        print('不自动打开文件')
    try:
        os.startfile(f'{ID}_error_log.txt')
    except FileNotFoundError:
        print('全部完成,未生成错误日志')


def download_ques_enterprise(ID, delay, begin, file_format, anl_switch):
    try:
        os.mkdir(rf'.\{ID}')
        print('目录创建成功')
    except FileExistsError:
        print('该题库已下载', flush=True)
        pass

    doc = Document()
    doc.styles['Normal'].font.name = u'宋体'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    doc.styles['Normal'].font.size = Pt(11)

    wb = xlwt.Workbook(encoding='utf-8')  # 新建一个excel文件
    ws1 = wb.add_sheet('sheet1')  # 添加一个新表，名字为begin
    ws1.write(0, 0, '序号')
    ws1.write(0, 1, '题目')
    ws1.write(0, 2, 'A')
    ws1.write(0, 3, 'B')
    ws1.write(0, 4, 'C')
    ws1.write(0, 5, 'D')
    ws1.write(0, 6, 'E')
    ws1.write(0, 7, 'F')
    ws1.write(0, 8, 'G')
    ws1.write(0, 9, 'H')
    ws1.write(0, 10, '正确答案')
    ws1.write(0, 11, '解析')
    login_url = f'https://s.kaoshibao.com/student'
    url = f'https://s.kaoshibao.com/online/?paperId={ID}'
    page = ChromiumPage()

    tab = page.get_tab(url='https://s.kaoshibao.com/sctk/')
    if tab is None:
        page.get(login_url)
        page.wait.url_change('https://s.kaoshibao.com/sctk/', timeout=30)
        page.wait(1)
    tab = page.new_tab()
    tab.get(url)
    tab.set.activate()
    tab.wait.eles_loaded('xpath://*[@id="body"]/div/div[1]/div[2]/div[1]/div[1]/div/div[1]/div/span[2]')
    number = tab.ele('xpath://*[@id="body"]/div/div[1]/div[2]/div[1]/div[1]/div/div[1]/div/span[2]').text[2:-1]
    try:
        tab.ele(f'tag:span@text():{begin}').click()
    except Exception as e:
        print(e)
    # 打开背题模式
    try:
        tab.wait(3)
        button_off = tab.s_ele('xpath://*[@id="body"]/div/div[1]/div[2]/div[2]/div[3]/p[2]/span[2]/div')
        auto_bext_button = tab.s_ele('xpath://*[@id="body"]/div/div[1]/div[2]/div[2]/div[3]/p[1]/span[2]/div')
        try:
            a = auto_bext_button.attr('aria-checked')
            if a == 'true':
                tab.ele('xpath://*[@id="body"]/div/div[1]/div[2]/div[2]/div[3]/p[1]/span[2]/div/input').click()
                print('关闭答对自动下一题', flush=True)
            else:
                print('答对自动下一题已关闭', flush=True)
        except Exception as e:
            print(e)
            sys.exit(0)
        try:
            a = button_off.attr('aria-checked')
            if a is None:
                tab.ele('xpath://*[@id="body"]/div/div[1]/div[2]/div[2]/div[3]/p[2]/span[2]/div/input').click()
                print('点击背题模式按钮', flush=True)
            else:
                print('背题模式已打开或已禁用', flush=True)
        except Exception as e:
            raise Exception(ElementNotFoundError)
    except ElementNotFoundError:
        print('背题模式已打开或已禁用', flush=True)
        tab.wait(0.3)

    for i in range(int(begin)-1, int(number)):
        try:
            title = f"{tab.ele('@class=qusetion-box').text}".replace('\n', '')
            doc.add_paragraph(f'{i + 1}.{title}')
            try:
                ques_img = tab.s_ele('@class=qusetion-box').ele('tag:img')
                if ques_img.link:
                    ques_img_url = ques_img.attr('src')
                    # ques_img_url = f'{ques_img_url}'
                    ques_img = tab.download(ques_img_url, rf'.\{ID}\imgs\ques', rename=f'ques{i + 1}-title',
                                            file_exists='skip')
                    tab.wait(0.3)
                doc.add_picture(check_suffix(ques_img[1]))
            except ElementNotFoundError:
                pass
            topic = tab.ele('@class=topic-type').text
            try:
                c = tab.s_ele('@@class=topic-type@@text()=案例分析题')
                if c:
                    print('暂不支持该题型')
                    sys.exit(0)
                else:
                    pass
            except ElementNotFoundError:
                pass
            option = ''
            if topic == '单选题':
                try:
                    tab.ele('@@class^before-icon@@text()=A').click()
                except Exception as e:
                    print(e)
                options = tab.s_eles('@class^option')
                for j in options:
                    try:
                        # 定位选项内的类名以'fr-fic '开头的元素，假设它代表选项图片
                        option_img_url = j.s_ele('tag:img').link
                        # 定位当前选项内的类名以'before-icon'开头的元素
                        x = j.s_ele('@class^before-icon')  # 更改此处，确保使用当前选项的'before-icon'元素
                        # 下载选项图片到指定目录，并重命名
                        option_img = tab.download(option_img_url, rf'.\{ID}\imgs\option',
                                                  rename=f'ques{i + 1}-option-{x.text}',
                                                  file_exists='skip')
                        tab.wait(0.3)
                        para = doc.add_paragraph()
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        run = para.add_run(str_j)  # 添加选项文本
                        run.add_picture(check_suffix(option_img[1]), width=Inches(2.5))
                    except Exception as e:
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        doc.add_paragraph(str_j)
                        option += str_j + "\n"
                answer = tab.s_ele('@class=right-ans').text.replace('\u2003', ':')
            elif topic == '判断题':
                try:
                    tab.ele('@@class^before-icon@@text()=A').click()
                except Exception as e:
                    print(e)
                options = tab.ele('@class^select-left').children('@class^option')
                for j in options:
                    list_j = list(j.text)
                    list_j.insert(1, '.')
                    str_j = ''.join(list_j)
                    doc.add_paragraph(str_j)
                    option += str_j + "\n"
                answer = tab.s_ele('@class=right-ans').text.replace('\u2003', ':')
            elif topic == '多选题':
                if tab.s_ele('xpath://*[@id="body"]/div/div[1]/div[2]/div[2]/div[3]/p[2]/span[2]/div').attr('aria-checked') is None:
                    try:
                        tab.ele('@@class^before-icon@@text()=A').click()
                        tab.wait(0.1)
                        tab.ele('xpath://*[@id="body"]/div/div[1]/div[2]/div[1]/div[1]/div/div[3]/button').click()
                    except Exception as e:
                        print(e)
                else:
                    pass
                options = tab.s_eles('@class^option')
                for j in options:
                    try:
                        # 定位选项内的类名以'fr-fic '开头的元素，假设它代表选项图片
                        option_img_url = j.s_ele('tag:img').link
                        # 定位当前选项内的类名以'before-icon'开头的元素
                        x = j.s_ele('@class^before-icon')  # 更改此处，确保使用当前选项的'before-icon'元素
                        # 下载选项图片到指定目录，并重命名
                        option_img = tab.download(option_img_url, rf'.\{ID}\imgs\option',
                                                  rename=f'ques{i + 1}-option-{x.text}', file_exists='skip')
                        tab.wait(0.3)
                        para = doc.add_paragraph()
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        run = para.add_run(str_j)  # 添加选项文本
                        run.add_picture(check_suffix(option_img[1]), width=Inches(2.5))
                    except Exception as e:
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        doc.add_paragraph(str_j)
                        option += str_j + "\n"
                answer = tab.s_ele('@class=right-ans').text.replace('\u2003', ':')
            elif topic == '不定项选择题':
                if tab.s_ele('xpath://*[@id="body"]/div/div[1]/div[2]/div[2]/div[3]/p[2]/span[2]/div').attr('aria-checked') is None:
                    try:
                        tab.ele('@@class^before-icon@@text()=A').click()
                        tab.wait(0.1)
                        tab.ele('xpath://*[@id="body"]/div/div[1]/div[2]/div[1]/div[1]/div/div[3]/button').click()
                    except Exception as e:
                        print(e)
                else:
                    pass
                options = tab.s_eles('@class^option')
                for j in options:
                    try:
                        # 定位选项内的类名以'fr-fic '开头的元素，假设它代表选项图片
                        option_img_url = j.s_ele('tag:img').link
                        # 定位当前选项内的类名以'before-icon'开头的元素
                        x = j.s_ele('@class^before-icon')  # 更改此处，确保使用当前选项的'before-icon'元素
                        # 下载选项图片到指定目录，并重命名
                        option_img = tab.download(option_img_url, rf'.\{ID}\imgs\option',
                                                  rename=f'ques{i + 1}-option-{x.text}', file_exists='skip')
                        tab.wait(0.3)
                        para = doc.add_paragraph()
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        run = para.add_run(str_j)  # 添加选项文本
                        run.add_picture(check_suffix(option_img[1]), width=Inches(2.5))
                    except Exception as e:
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        doc.add_paragraph(str_j)
                        option += str_j + "\n"
                answer = tab.s_ele('@class=right-ans').text.replace('\u2003', ':')
            elif topic == '排序题':
                options = tab.s_eles('@class^option')
                for j in options:
                    try:
                        # 定位选项内的类名以'fr-fic '开头的元素，假设它代表选项图片
                        option_img_url = j.s_ele('tag:img').link
                        # 定位当前选项内的类名以'before-icon'开头的元素
                        x = j.s_ele('@class^before-icon')  # 更改此处，确保使用当前选项的'before-icon'元素
                        # 下载选项图片到指定目录，并重命名
                        option_img = tab.download(option_img_url, rf'.\{ID}\imgs\option',
                                                  rename=f'ques{i + 1}-option-{x.text}', file_exists='skip')
                        tab.wait(0.3)
                        para = doc.add_paragraph()
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        run = para.add_run(str_j)  # 添加选项文本
                        run.add_picture(check_suffix(option_img[1]), width=Inches(2.5))
                    except Exception as e:
                        list_j = list(j.text)
                        list_j.insert(1, '.')
                        str_j = ''.join(list_j)
                        doc.add_paragraph(str_j)
                        option += str_j + "\n"
                answer = tab.s_ele('@class=right-ans').text.replace('\u2003', ':')
            elif topic == '填空题':
                try:
                    s = tab.s_ele('xpath://*[@id="body"]/div/div[1]/div[2]/div[2]/div[3]/p[2]/span[2]/div').attr(
                            'aria-checked')
                    if s == 'true':
                        pass
                    else:
                        try:
                            tab.ele('@class=el-input__inner').input('1')
                            tab.wait(0.1)
                            tab.ele('xpath://*[@id="body"]/div/div[1]/div[2]/div[1]/div[1]/div/div[2]/button').click()
                        except Exception as e:
                            print(e)
                except Exception as e:
                    print(e)
                answer = tab.s_ele('@class=right-ans').text.replace('\u2003', ':')
            elif topic == '简答题':
                try:
                    s = tab.s_ele('xpath://*[@id="body"]/div/div[1]/div[2]/div[2]/div[3]/p[2]/span[2]/div').attr(
                            'aria-checked')
                    if s == 'true':
                        pass
                    else:
                        try:
                            tab.ele('@class=el-textarea__inner').input('1')
                            tab.wait(0.1)
                            tab.ele('xpath://*[@id="body"]/div/div[1]/div[2]/div[1]/div[1]/div/div[2]/button[2]').click()
                        except Exception as e:
                            print(e)
                except Exception as e:
                    print(e)
                answer = tab.s_ele('@class=right-ans').text.replace('\u2003', ':')
            elif topic == '论述题':
                answer = tab.s_ele('@class=right-ans ').text.replace('\u2003', ':')
            else:
                print('暂不支持该题型')

            '''formatted_option = "\n".join(
                f"{line[0]}. {line[1:]}" if line[0].isupper() else line for line in option.splitlines())'''
            analysis = ''
            if anl_switch == '是':
                try:
                    analysis = tab.s_ele('@class^answer-analysis').text.replace('\n', '')
                    try:
                        analysis_img = tab.s_ele('@class^answer-analysis').ele('tag:img')
                        if analysis_img.link:
                            analysis_img_url = analysis_img.attr('src')
                            if analysis_img_url == 'https://resource.zaixiankaoshi.com/mini/ai_tag.png':
                                pass
                            else:
                                analysis_img_url = f'{analysis_img_url}'
                                analysis_img = tab.download(analysis_img_url, rf'.\{ID}\imgs\analysis',
                                                            rename=f'ques{i + 1}-analysis', file_exists='skip')
                                tab.wait(0.3)
                    except ElementNotFoundError:
                        pass
                except Exception as e:
                    print(e)
            if option != '':
                ques = f'{i + 1}.{title}\n{option}{answer}\n{analysis}\n\n'
                option1 = option.replace('\n', '&@')
                ques1 = f'{i + 1}&@{title}&@{option1}&@{answer[5:]}&@{analysis}\n'
            else:
                ques = f'{i + 1}.{title}\n{option}{answer}\n{analysis}\n\n'
                ques1 = f'{i + 1}&@{title}&@{answer[5:]}&@{analysis}\n'
            # 添加答案段落
            doc.add_paragraph(answer)
            try:
                answer_img = page.s_ele('@class^mt20', timeout=0.3).ele('tag:img', timeout=0.3)
                if answer_img.link:
                    answer_img_url = answer_img.attr('src')
                    # ques_img_url = f'{ques_img_url}'
                    answer_img = page.download(answer_img_url, rf'.\{ID}\imgs\answer', rename=f'ques{i + 1}-answer',
                                               file_exists='skip')
                doc.add_picture(check_suffix(answer_img[1]), width=Inches(2.5))
            except ElementNotFoundError:
                pass
            doc.add_paragraph(f'{analysis} \n')
            try:
                doc.add_picture(check_suffix(analysis_img[1]), width=Inches(2.5))
            except Exception as e:
                pass
            list_a = ques1.split('&@')
            while len(list_a) <= 4:
                list_a.insert(2, '')
            while 4 < len(list_a) < 12:
                list_a.insert(-3, '')
            try:
                ws1.write(i + 1, 0, int(list_a[0]))
                ws1.write(i + 1, 1, list_a[1])
                ws1.write(i + 1, 2, list_a[2])
                ws1.write(i + 1, 3, list_a[3])
                ws1.write(i + 1, 4, list_a[4])
                ws1.write(i + 1, 5, list_a[5])
                ws1.write(i + 1, 6, list_a[6])
                ws1.write(i + 1, 7, list_a[7])
                ws1.write(i + 1, 8, list_a[8])
                ws1.write(i + 1, 9, list_a[9])
                ws1.write(i + 1, 10, list_a[-2])
                ws1.write(i + 1, 11, list_a[-1][5:])
            except IndexError as e:
                pass
            wb.save(rf'.\{ID}\{ID}-第{begin}题开始.xls')
            info = f'第{i + 1}题已完成'
            try:
                print(ques, flush=True)
            except Exception as e:
                print(e)
                print(info, flush=True)
            filepath = rf'.\{ID}\{ID}-第{begin}题开始.txt'
            with open(filepath, "a", encoding='utf8') as f:
                f.write(ques)  # 自带文件关闭功能，不需要再写f.close()
            doc.save(rf'.\{ID}\{ID}-第{begin}题开始.docx')
            try:
                tab.ele('@@class:el-button el-button--primary el-button--small@@text():下一题', timeout=5).click()
                tab.wait(float(delay))
            except Exception as e:
                print(e)
        except ElementNotFoundError:
            print(f'第{i + 1}题下载失败\n', flush=True)
            with open(f'{ID}_error_log.txt', "a", encoding='utf8') as f:
                f.write(f'第{i + 1}题下载失败\n')  # 自带文件关闭功能，不需要再写f.close()
            try:
                tab.ele('@@class:el-button el-button--primary el-button--small@@text():下一题', timeout=5).click()
                tab.wait(float(delay))
            except Exception as e:
                print(e)
        continue
    if file_format == '.txt':
        os.startfile(rf'.\{ID}\{ID}-第{begin}题开始.txt')
    elif file_format == '.docx':
        os.startfile(rf'.\{ID}\{ID}-第{begin}题开始.docx')
    elif file_format == '.xls':
        os.startfile(rf'.\{ID}\{ID}-第{begin}题开始.xls')
    elif file_format == '不自动打开':
        print('不自动打开文件')
    try:
        os.startfile(f'{ID}_error_log.txt')
    except FileNotFoundError:
        print('全部完成,未生成错误日志')


if __name__ == '__main__':
    start()
