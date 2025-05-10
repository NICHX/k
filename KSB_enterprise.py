# -*- coding: utf-8 -*-
import codecs
import configparser
import os
import sys
import time
import requests
import wmi
import xlwt
from DrissionPage import ChromiumPage
from DrissionPage.common import Settings
from DrissionPage.errors import ElementNotFoundError
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.shared import Pt
from gooey import Gooey, GooeyParser
from func import download_ques_enterprise


Settings.raise_when_ele_not_found = True

if sys.stdout.encoding != 'UTF-8':
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
if sys.stderr.encoding != 'UTF-8':
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')


c = wmi.WMI()
version = '1.4.0'


@Gooey(language='chinese', program_name=u'KSB工具(enterprise_version) beta', required_cols=2, optional_cols=2,
       enterprise=True, clear_before_run=True, sidebar_title='工具列表', terminal_font_family='Courier New',
       menu=[{
           'name': '关于',
           'items': [{
               'type': 'AboutDialog',
               'menuTitle': '关于',
               'name': 'KSB工具(enterprise_version) beta\n',
               'description': 'Created by NICHX !\n 1、新增自动更新功能！',

               'version': version,
           }]
       }])
def KSB_window():
    parser = GooeyParser(
        description="安装谷歌Chrome浏览器！")
    subs = parser.add_subparsers(help='KSB', dest='command')

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


    if args.command == 'KSB企业版':
        download_ques_enterprise(args.题库ID, args.延迟时间, args.起始题号, args.默认打开文件, args.解析开关)


if __name__ == '__main__':
    KSB_window()
