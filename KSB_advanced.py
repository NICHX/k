# -*- coding: utf-8 -*-
import codecs
import sys
from DrissionPage.common import Settings
from gooey import Gooey, GooeyParser

from func_advanced import download_ques_advanced


Settings.raise_when_ele_not_found = True

if sys.stdout.encoding != 'UTF-8':
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
if sys.stderr.encoding != 'UTF-8':
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')


@Gooey(language='chinese', program_name=u'KSB工具(advanced_version)', required_cols=3, optional_cols=3,
       advanced=True, clear_before_run=True, sidebar_title='工具列表', terminal_font_family='Courier New',
       menu=[{
           'name': '关于',
           'items': [{
               'type': 'AboutDialog',
               'menuTitle': '关于',
               'name': 'KSB工具(advanced_version)',
               'description': f'Created by NICHX !',
           }]
       }])
def KSB_window():
    parser = GooeyParser(
        description="安装谷歌Chrome浏览器！")
    subs = parser.add_subparsers(help='KSB', dest='command')
    normal_parser = subs.add_parser('KSB', help='KSB工具')
    subgroup = normal_parser.add_argument_group('配置信息')
    subgroup.add_argument('题库ID', help="请输入题库ID", widget='TextField')
    subgroup.add_argument('题库名称', help="请输入题库名称", widget='TextField')
    subgroup.add_argument('解析开关', help="是否需要解析", widget='Dropdown'
                          , choices=['是', '否'], default='是')
    subgroup.add_argument('起始题号', help="从哪一题开始下载", widget='TextField'
                          , default='1')
    subgroup.add_argument('默认打开文件', help="请选择完成后想要打开的文件类型", widget='Dropdown'
                          , choices=['.txt', '.docx', '.xlsx', '不自动打开'], default='不自动打开')
    subgroup.add_argument('延迟时间', help="爬取延迟时间，默认0.4，若有题目重复手动调高", widget='TextField',
                          default='0.4')
    subgroup.add_argument('超时时间', help="等待超时时间，默认0.05，若有异常手动调高", widget='TextField',
                          default='0.05')

    args = parser.parse_args()

    if args.command == 'KSB':
        download_ques_advanced(args.题库ID, args.题库名称 ,args.延迟时间, args.起始题号, args.默认打开文件, args.解析开关, args.超时时间)

if __name__ == '__main__':
    KSB_window()


