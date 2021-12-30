# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog
import os
import glob
import shutil  # 对文件的高级操作
from tkinter.ttk import Separator  # 分割线模块
import tkinter.messagebox  # 弹窗库
from tkinter import ttk

import json
import sys
import time
from datetime import datetime
from time import sleep
import pandas as pd
import openpyxl

# Some Constants
input_dir = './input/'
output_dir = './output/'
csv_files_extension_name = '*.csv'
txt_files_extension_name = '*.txt'
xls_files_extension_name = '*.xls'
xlsx_files_extension_name = '*.xlsx'
input_dir_logs = './input/'

output_file = 'output.json'

csv_encoding = ['utf_16', 'utf_8', 'gb18030', 'gb2312']

# folder or directory functions.


def check_folder(dir_name):
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)
    print("Checking Dir: " + dir_name)
    return dir_name


def press_and_continue():
    a = input("Press ENTER key to continue.")


def press_and_exit(code):
    a = input("Press ENTER key to exit.")
    sys.exit(code)

# 用于遍历指定路径内要求文件数据


def get_filelist(path):
    Filelist = []
    for home, dirs, files in os.walk(path):
        for filename in files:
            (shotname, extension) = os.path.splitext(filename)
            Filelist.append(shotname)
    return Filelist


'''算法和界面分割栏'''

# 上传文件


def upload_data():
    # 新建input、output文件夹
    if not os.path.exists(input_dir):
        os.makedirs(input_dir)
    else:
        # 清空历史缓存文件（强制删除文件夹）
        shutil.rmtree(input_dir)
        os.makedirs(input_dir)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    else:
        # 清空历史缓存文件（强制删除文件夹）
        shutil.rmtree(output_dir)
        os.makedirs(output_dir)
    global file_path_dcs
    file_path = filedialog.askopenfilenames(
        title=u'请选择数据文件',
        filetypes=[
            ('CSV文件',
             '*.csv'),
            ('Excel',
             '*.xls'),
            ('Excel',
             '*.xlsx'),
            ('文本文件',
             '*.txt'),
            ('其他格式(暂不支持)',
             '*')],
        initialdir=(
            os.path.expanduser(desktop_path)))
    if len(file_path) == 0:
        text_preview.delete('0.0', 'end')
        text_preview.insert("insert", "上传失败!请重新上传")
    else:
        text_preview.delete('0.0', 'end')
        text_preview.insert("insert", "上传操作:成功!")

    # 后台拷贝dcs_csv到指定目录
    for i in range(len(file_path)):
        # 获取文件原始文件名称
        file_name = str(file_path[i]).split('/')[-1]
        print(file_name)
        target_file = input_dir + file_name
        shutil.copyfile(file_path[i], target_file)

# 一键处理


def start_processing():
    text_preview.delete('0.0', 'end')
    text_preview.insert("insert", "正在处理中……")
    # os.system('csv_format_conversion_tool.py')
    check_folder(input_dir_logs)
    check_folder(output_dir)
    # Sort and read input files.
    csv_files = sorted(glob.glob(input_dir_logs + csv_files_extension_name))
    txt_files = sorted(glob.glob(input_dir_logs + txt_files_extension_name))
    xls_files = sorted(glob.glob(input_dir_logs + xls_files_extension_name))
    xlsx_files = sorted(glob.glob(input_dir_logs + xlsx_files_extension_name))
    ori_name=get_filelist(input_dir_logs)
    if (len(csv_files) == 0)&(len(txt_files) == 0)&(len(xls_files) == 0)&(len(xlsx_files) == 0):
        print('ERROR: There are NOT Valid files in the directory %s.' % (input_dir_logs))
        press_and_exit(1)
    elif len(csv_files)!=0:
        print('There are [ %d ] csv files in the directory %s' % (len(csv_files), input_dir_logs))
        print(csv_files)
    elif len(txt_files) != 0:
        print('There are [ %d ] txt files in the directory %s' % (len(txt_files), input_dir_logs))
        print(txt_files)
    elif len(xls_files) != 0:
        print('There are [ %d ] xls files in the directory %s' % (len(xls_files), input_dir_logs))
        print(xls_files)
    elif len(xlsx_files) != 0:
        print('There are [ %d ] xls files in the directory %s' % (len(xlsx_files), input_dir_logs))
        print(xlsx_files)


    # CSV文件处理
    csv_sheets = pd.Series([])
    for (idx, sample_file) in enumerate(csv_files):
        try:
            for item in csv_encoding:
                try:
                    print('trying to read the file [%d] %s with the encoding %s,' % (idx + 1, sample_file, item),
                          end=' ')
                    # 日立设备 以及 ABB设备
                    csv = pd.read_csv(sample_file, sep=r'\s*\t|\s*,''', engine='python', encoding=item, header=None
                                      , squeeze=True, index_col=False)
                    print('\t[OK].')
                    csv = csv.replace("'", '', regex=True)
                    format_detection = csv.loc[0][0]
                    # 不同文件在此判断，abb系统与日立系统
                    if (format_detection == 'Time'):
                        # abb系统csv文件,如开封ABB
                        # remove the column with NAN
                        csv = csv.T.dropna()
                        csv = csv.T
                        title_list = csv.loc[0].values
                        csv.columns = title_list
                        csv = csv.drop(index=0)
                        csv['Time'] = pd.to_datetime(csv['Time'])
                        csv['Time'] = csv['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))  # 将时间戳转化为要求格式
                        csv.set_index('Time', inplace=True)
                        csv = csv.replace(" ", '', regex=True)
                        # 首阳山#1机组（日立设备）需删除“（B）-”字符
                        csv = csv.replace("B", '', regex=True)
                        # 批量删除括号
                        csv = csv.replace(r'[\(\)]+', '', regex=True)
                        # 批量浮点型转换
                        csv = csv.astype('float64')
                        print("转换后表格:")
                        csv.to_csv(output_dir + sample_file[8:-4] + ".csv", encoding='utf-8')
                        print(csv)
                        print(csv.dtypes)
                    elif (format_detection == '时间'):
                        # 日立系统csv文件，如阜阳dcs\洛热-日立--更新为utf-8格式
                        # 新需求 重新剔除所有中文释义
                        # title_list = csv.loc[1].tolist()
                        # title_list.insert(0, ' ')
                        # del title_list[-1]
                        # csv = csv.drop(index=1)
                        # csv = csv.replace("时间", 'Time', regex=True)
                        # # 由于新增中文索引，因此无法确保一列以统一时间转换
                        # csv.columns = title_list
                        # csv.set_index(title_list[0], inplace=True)
                        csv = csv.replace("时间", 'Time', regex=True)
                        title_list = csv.loc[0].tolist()
                        csv.columns = title_list
                        csv = csv.drop(index=0)
                        csv = csv.drop(index=1)
                        csv.set_index('Time', inplace=True)
                        # 批量浮点型转换
                        csv = csv.replace(" ", '', regex=True)
                        # 贺州#2机组（日立设备）需删除“（B）-”字符
                        csv = csv.replace("B", '', regex=True)
                        # 批量删除括号--贺州机组某组数据有五角星符号
                        csv = csv.replace(r'[\(\)]+', '', regex=True)
                        csv = csv.replace("☆", '', regex=True)
                        # print(csv.iloc[3002].values)
                        # 删除含空值的行
                        csv = csv.dropna()
                        csv = csv.astype('float64')
                        print("转换后表格:")
                        csv.to_csv(output_dir + sample_file[8:-4] + ".csv", encoding='utf-8')
                        print(csv)
                        print(csv.dtypes)
                    else:
                        # 其余暂认为未知类型(猜测为ovation导出为csv)如20210208.csv
                        # 标记时间戳起始标签
                        time_tag = csv.loc[csv.iloc[:, 0] == "Date Time"]
                        strat_tag = int(time_tag.index[0])  # + 1
                        # 根据检索标签跳过无用行
                        csv = pd.read_csv(sample_file, sep=r'\s*\t|\s*,', skiprows=strat_tag, engine='python',
                                          encoding=item, header=0
                                          , squeeze=True, index_col=False)
                        # 删除换行符及后缀
                        csv = csv.replace('\n', ' ', regex=True)
                        # 删除空列
                        csv = csv.T.dropna()
                        csv = csv.T
                        # xls = xls.drop(index=0)
                        csv.rename(columns={'Date Time': 'Time'}, inplace=True)
                        # 将时间列转化为datetime对象
                        csv['Time'] = pd.to_datetime(csv['Time'])
                        # 将时间戳转化为要求格式
                        csv['Time'] = csv['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))
                        csv.set_index('Time', inplace=True)
                        # 删除英文TRUE
                        csv = csv.replace("TRUE", '', regex=True)
                        csv = csv.replace("FALSE", '', regex=True)
                        # 批量浮点型转换
                        # print(csv.iloc[2].values)
                        # csv = csv.astype('float64')
                        print("转换后表格:")
                        # 按原名称另存为utf_8格式的csv文件
                        csv.to_csv(output_dir + sample_file[8:-4] + ".csv", encoding='utf-8')
                        print(csv)
                        print(csv.dtypes)
                    break
                except:
                    print("\t[Failed].")
                try:
                    print('trying to read the file [%d] %s with the encoding %s,' % (idx + 1, sample_file, item),
                          end=' ')
                    # 海丰#1机组 SIS数据
                    csv = pd.read_csv(sample_file, sep=r'\s*\t|\s*,''', engine='python', encoding=item
                                      , squeeze=True, index_col=False)  # , header=None
                    print('\t[OK].')
                    csv = csv.replace("'", '', regex=True)
                    # 带标题读空列名为默认为“Unnamed: 0”
                    # 不同文件在此判断，SIS数据首行首列为空
                    if (csv.columns.tolist()[0] == 'Unnamed: 0'):
                        # 海丰#1机组
                        # 重命名列，将第一列替换为指定名称
                        csv.rename(columns={'Unnamed: 0': 'Time'}, inplace=True)
                        # remove the column with NAN
                        csv = csv.T.dropna()
                        csv = csv.T
                        csv['Time'] = pd.to_datetime(csv['Time'])
                        csv['Time'] = csv['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))  # 将时间戳转化为要求格式
                        csv.set_index('Time', inplace=True)
                        # 批量浮点型转换
                        csv = csv.astype('float64')
                        print("转换后表格:")
                        csv.to_csv(output_dir + sample_file[8:-4] + ".csv", encoding='gb18030')
                        print(csv)
                        print(csv.dtypes)
                        break
                except:
                    print("\t[Failed].")
                try:
                    print('trying to read the file [%d] %s with the encoding %s,' % (idx + 1, sample_file, item),
                          end=' ')
                    csv = pd.read_csv(sample_file, sep=r',''', engine='python', encoding=item
                                      , squeeze=True, index_col=False)  # , header=None
                    # 寻找两张表的分割标签（Date Time）
                    for i in range(len(csv.index.tolist())):
                        if (csv.loc[i] == "Date Time"):
                            # 设置跳过行数
                            strat_tag = i + 3
                            break
                    # 广热#1机组
                    csv = pd.read_csv(sample_file, sep=r'\s*\t|\s*,', skiprows=strat_tag, engine='python',
                                      encoding=item,
                                      index_col=False)  # , squeeze=True header= None,
                    print('\t[OK].')
                    # 删除两行英文注释行（Auto Historian && Actual）
                    csv = csv.drop(index=0)
                    csv = csv.drop(index=1)
                    # remove the column with NAN
                    csv = csv.T.dropna()
                    csv = csv.T
                    # 删除全文所有双引号
                    csv.rename(columns=lambda x: x.replace('"', ''), inplace=True)
                    csv = csv.replace('"', ' ', regex=True)
                    # 新增处理曹妃甸#3机组 符号处理
                    csv = csv.replace("B", '', regex=True)
                    csv = csv.replace("F", '', regex=True)
                    csv = csv.replace("P", '', regex=True)
                    # 重命名Date Time
                    csv.rename(columns={'Date Time': 'Time'}, inplace=True)
                    csv['Time'] = pd.to_datetime(csv['Time'])
                    csv['Time'] = csv['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))  # 将时间戳转化为要求格式
                    csv.set_index('Time', inplace=True)
                    # 批量浮点型转换
                    csv = csv.astype('float64')
                    print("转换后表格:")
                    csv.to_csv(output_dir + sample_file[8:-4] + ".csv", encoding='utf-8')
                    print(csv)
                    print(csv.dtypes)
                    break
                except:
                    print("\t[Failed].")
                try:
                    print('trying to read the file [%d] %s with the encoding %s,' % (idx + 1, sample_file, item),
                          end=' ')
                    # 沧州#2机组数据（时间戳为24-八月-2018 12:00:00）
                    csv = pd.read_csv(sample_file, sep=r'\s*\t|\s*,''', engine='python', encoding=item, header=None
                                      , squeeze=True, index_col=False)
                    print('\t[OK].')
                    csv = csv.replace("'", '', regex=True)
                    format_detection = csv.loc[0][0]
                    if (format_detection == 'Time'):
                        # remove the column with NAN
                        csv = csv.T.dropna()
                        csv = csv.T
                        title_list = csv.loc[0].values
                        csv.columns = title_list
                        csv = csv.drop(index=0)
                        # 中文月份转换
                        csv['Time'] = csv['Time'].replace('一月', 'Jan', regex=True)
                        csv['Time'] = csv['Time'].replace('二月', 'Feb', regex=True)
                        csv['Time'] = csv['Time'].replace('三月', 'Mar', regex=True)
                        csv['Time'] = csv['Time'].replace('四月', 'Apr', regex=True)
                        csv['Time'] = csv['Time'].replace('五月', 'May', regex=True)
                        csv['Time'] = csv['Time'].replace('六月', 'Jun', regex=True)
                        csv['Time'] = csv['Time'].replace('七月', 'Jul', regex=True)
                        csv['Time'] = csv['Time'].replace('八月', 'Aug', regex=True)
                        csv['Time'] = csv['Time'].replace('九月', 'Sept', regex=True)
                        csv['Time'] = csv['Time'].replace('十月', 'Oct', regex=True)
                        csv['Time'] = csv['Time'].replace('十一月', 'Nov', regex=True)
                        csv['Time'] = csv['Time'].replace('十二月', 'Dec', regex=True)
                        csv['Time'] = pd.to_datetime(csv['Time'])
                        csv['Time'] = csv['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))  # 将时间戳转化为要求格式
                        csv.set_index('Time', inplace=True)
                        csv = csv.replace(" ", '', regex=True)
                        # print(csv.iloc[2].values)
                        # 批量浮点型转换
                        csv = csv.astype('float64')
                        print("转换后表格:")
                        csv.to_csv(output_dir + sample_file[8:-4] + ".csv", encoding='utf-8')
                        print(csv)
                        print(csv.dtypes)
                    break
                except:
                    print("\t[Failed].")
        except:
            print(', [Failed]')
            print("  - ERROR: Opening the file: %s, and Skip the file." % (sample_file))
            print(sys.exc_info())
            continue

    # TXT文件处理
    txt_sheets = pd.Series([])
    for (idx, sample_file) in enumerate(txt_files):
        try:
            for item in csv_encoding:
                try:
                    print(
                        'trying to read the file [%d] %s ,looking for specific flags "Data Time",with the encoding:utf-16' % (
                        idx + 1, sample_file),
                        end=' ')
                    # 以特定格式打开，例如赤壁机组DCS.txt文件
                    sign_character = 'Date Time'
                    file_txt = open(sample_file, 'r', encoding='utf-16').readlines()
                    for i in range(len(file_txt)):
                        if sign_character in file_txt[i]:
                            # 设置跳过行数
                            time_tag = i
                            break
                    txt = pd.read_csv(sample_file, sep=r'\t', skiprows=time_tag, engine='python', encoding=item)
                    print('\t[OK].')
                    # 删除空行、列
                    txt = txt.dropna()
                    txt = txt.T.dropna()
                    txt = txt.T
                    # 重命名列，将第一列替换为指定名称
                    ColNames_List = txt.columns.tolist()
                    ColNames_List[0] = 'Time'
                    list_space = []
                    # 寻找不可见字符列及特殊规则列
                    for i in range(len(ColNames_List)):
                        if str(ColNames_List[i]).isspace():
                            list_space.append(i - 1)
                        # xx-1文件及部分xx-4文件会产生不可见字符读为   .1
                        if str(ColNames_List[i]) == '  .1':
                            list_space.append(i - 1)
                    # print(list_space)
                    txt.columns = ColNames_List
                    # 替换空值为空
                    txt.dropna(axis=0, how='any')
                    txt.dropna(axis=1, how='any')
                    # 将时间列转化为datetime对象
                    txt['Time'] = pd.to_datetime(txt['Time'])
                    # 将时间戳转化为要求格式
                    txt['Time'] = txt['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))
                    txt.set_index('Time', inplace=True)
                    # 剔除不可见字符列
                    txt = txt.drop(txt.columns[list_space], axis=1)
                    # 剔除数据侧空格和不可见\x00字符以及特殊字符 “P”（湖北#1 250-4/330-4）
                    txt = txt.replace(r'\x00', '', regex=True)
                    txt = txt.replace(' ', '', regex=True)
                    txt = txt.replace('P', '', regex=True)
                    # 批量浮点型转换
                    txt = txt.astype('float64')
                    txt.dropna(axis=1, how='any')
                    txt.to_csv(output_dir + sample_file[8:-4] + ".csv", encoding='utf-8')
                    print(txt)
                    print(txt.dtypes)
                    break
                except:
                    print("\t[Failed].")
                try:
                    print(
                        'trying to read the file [%d] %s ,looking for specific flags "Data/Time",with the encoding:%s' % (
                        idx + 1, sample_file, item),
                        end=' ')
                    # 以特定格式首行(Date/Time)打开，例如赤壁机组#3号机组的DCS.txt文件
                    time_tag = 0
                    txt = pd.read_csv(sample_file, sep=r'\t', skiprows=time_tag, engine='python',
                                      encoding=item)  # sep=r'\s*\t|\s*,'
                    print('\t[OK].')
                    # 删除空行、列
                    txt = txt.dropna()
                    txt = txt.T.dropna()
                    txt = txt.T
                    # 重命名列，将第一列替换为指定名称
                    ColNames_List = txt.columns.tolist()
                    ColNames_List[0] = 'Time'
                    list_space = []
                    # 寻找不可见字符列及特殊规则列
                    for i in range(len(ColNames_List)):
                        if str(ColNames_List[i]).isspace():
                            list_space.append(i - 1)
                        # xx-1文件及部分xx-4文件会产生不可见字符读为   .1
                        if str(ColNames_List[i]) == '  .1':
                            list_space.append(i - 1)
                    txt.columns = ColNames_List
                    # 替换空值为空
                    txt.dropna(axis=0, how='any')
                    txt.dropna(axis=1, how='any')
                    # 将时间列转化为datetime对象
                    txt['Time'] = pd.to_datetime(txt['Time'])
                    # 将时间戳转化为要求格式
                    txt['Time'] = txt['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))
                    txt.set_index('Time', inplace=True)
                    # 剔除不可见字符列
                    txt = txt.drop(txt.columns[list_space], axis=1)
                    # 删除换行符及后缀
                    txt.dropna(axis=1, how='any')
                    # 剔除标题侧空格和不可见\x00字符
                    txt.rename(columns=lambda x: x.replace("\x00", ''), inplace=True)
                    # 剔除数据侧空格和不可见\x00字符
                    txt = txt.replace(r'\x00', '', regex=True)
                    txt = txt.replace(' ', '', regex=True)

                    list_cspace = []
                    for i in range(len(txt.iloc[2].values)):
                        if list(txt.iloc[2].values)[i]=='':
                            list_cspace.append(i)
                    txt = txt.drop(txt.columns[list_cspace], axis=1)
                    #print(txt.iloc[2].values)
                    print(txt)
                    # 批量浮点型转换
                    txt = txt.astype('float64')
                    txt.to_csv(output_dir + sample_file[8:-4] + ".csv", encoding='utf-8')
                    print(txt)
                    print(txt.dtypes)
                    break
                except:
                    print("\t[Failed].")
                try:
                    print(
                        'trying to read the file [%d] %s ,looking for specific flags "Period",with the encoding %s' % (
                        idx + 1, sample_file, item),
                        end=' ')
                    # 以特定格式打开，例如foxborolA-新密1期.txt文件 --需合并前两列时间--需用\s+分隔--需将前三行对应列拼接为编号
                    sign_character = 'Period'
                    file_txt = open(sample_file, 'r', encoding=item).readlines()
                    line_1_list = str(file_txt[0])[21:].split()
                    line_2_list = str(file_txt[1])[21:].split()
                    line_3_list = str(file_txt[2])[21:].split()
                    ColNames_List = []
                    for i in range(len(line_1_list)):
                        word = line_1_list[i] + '_' + line_2_list[i] + '.' + line_3_list[i]
                        ColNames_List.append(word)
                    for i in range(len(file_txt)):
                        if sign_character in file_txt[i]:
                            # 设置跳过行数
                            time_tag = i
                            break
                    ColNames_List.insert(0, 'Time')
                    txt = pd.read_csv(sample_file, sep=r'\s+', skiprows=(time_tag + 3), parse_dates=[[0, 1]],
                                      header=None, engine='python', index_col=False, encoding=item)
                    print('\t[OK].')
                    # 删除空行、列
                    txt = txt.dropna()
                    txt = txt.T.dropna()
                    txt = txt.T
                    txt.columns = ColNames_List
                    # 替换空值为空
                    txt.dropna(axis=0, how='any')
                    txt.dropna(axis=1, how='any')
                    # 将时间列转化为datetime对象
                    txt['Time'] = pd.to_datetime(txt['Time'])
                    # 将时间戳转化为要求格式
                    txt['Time'] = txt['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))
                    txt.set_index('Time', inplace=True)
                    # 批量浮点型转换
                    txt = txt.astype('float64')
                    txt.dropna(axis=1, how='any')
                    # format(txt, ".6f")
                    txt.to_csv(output_dir + sample_file[8:-4] + ".csv", encoding='utf-8')
                    print(txt)
                    print(txt.dtypes)
                    break
                except:
                    print("\t[Failed].")
        except:
            print(', [Failed]')
            print("  - ERROR: Opening the file: %s, and Skip the file." % (sample_file))
            print(sys.exc_info())
            continue

    # XLS文件处理
    xls_sheets = pd.Series([])
    for (idx, sample_file) in enumerate(xls_files):
        try:
            for item in csv_encoding:
                try:
                    print('trying to read the file [%d] %s with the function read_excel %s,' % (
                    idx + 1, sample_file, item),
                          end=' ')
                    xls = pd.read_excel(sample_file)
                    print('\t[OK].')
                    # ovation设备，如鹤淇机组及其低压轴封.xls
                    # 标记时间戳起始标签
                    time_tag = xls.loc[xls.iloc[:, 0] == "Date Time"]
                    strat_tag = int(time_tag.index[0]) + 1
                    # 根据检索标签跳过无用行
                    xls = pd.read_excel(sample_file, skiprows=strat_tag, index_col=False)
                    # 删除换行符及后缀
                    xls = xls.replace('\n', ' ', regex=True)
                    xls = xls.replace(' B', '', regex=True)
                    xls = xls.replace(' P', '', regex=True)
                    # xls = xls.replace(' Auto HistorianActual', ' ', regex=True)
                    xls.rename(columns=lambda x: x.replace('\n', ''), inplace=True)
                    xls.rename(columns=lambda x: x.replace('Auto HistorianActual', ''), inplace=True)
                    # 删除空列
                    xls = xls.T.dropna()
                    xls = xls.T
                    # xls = xls.drop(index=0)
                    xls.rename(columns={'Date Time': 'Time'}, inplace=True)
                    # 将时间列转化为datetime对象
                    xls['Time'] = pd.to_datetime(xls['Time'])
                    # 将时间戳转化为要求格式
                    xls['Time'] = xls['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))
                    xls.set_index('Time', inplace=True)
                    # 批量浮点型转换
                    xls = xls.astype('float64')
                    print("转换后表格:")
                    # 按原名称另存为utf_8格式的csv文件
                    xls.to_csv(output_dir + sample_file[8:-4] + ".csv", encoding='utf-8')
                    print(xls)
                    print(xls.dtypes)
                    break
                except:
                    print("\t[Failed].")
                try:
                    print('trying to read the file [%d] %s with the encoding %s,' % (idx + 1, sample_file, item),
                          end=' ')
                    # xls = pd.read_excel(sample_file)
                    # read_excel方法读取失败，尝试读取其他格式.xls
                    xls = pd.read_csv(sample_file, encoding=item)
                    print('\t[OK].')
                    # 特殊xls格式，如赤壁IMP#1、#3机组，由于空行和首次/t分隔失败，需寻找特定字符("采集计算机：")
                    # 标记时间戳起始标签(特殊格式)
                    time_tag = xls.loc[xls.iloc[:, 0] == "采集计算机："]
                    # strat_tag = int(time_tag.index[0]) + 1
                    strat_tag = 2
                    # 重命名中文释义行（补齐第1列空白名称）
                    ColNames_List_old = xls.loc[5].tolist()
                    ColNames_List = str(ColNames_List_old)[:-2].split(r"\t")  # -4
                    ColNames_List[0] = 'Time'
                    # 根据检索标签跳过无用行,中文需用gbk格式读取
                    xls = pd.read_csv(sample_file, skiprows=(strat_tag + 7), sep=r'\t', parse_dates=[[0, 1]],
                                      names=ColNames_List, index_col=False, encoding="gb18030")
                    # 删除空列
                    xls = xls.T.dropna()
                    xls = xls.T
                    # 将时间列转化为datetime对象
                    xls.rename(columns={'Time_名称': 'Time'}, inplace=True)
                    xls['Time'] = pd.to_datetime(xls['Time'])
                    # 将时间戳转化为要求格式
                    xls['Time'] = xls['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))
                    xls.set_index('Time', inplace=True)
                    # 批量浮点型转换
                    xls = xls.astype('float64')
                    print("转换后表格:")
                    # 按原名称另存为gb18030格式的csv文件（针对IMP带中文名称的需求，统一设为gb18030格式）
                    xls.to_csv(output_dir + sample_file[8:-4] + ".csv", encoding='gb18030')
                    print(xls)
                    print(xls.dtypes)
                    break
                except:
                    print("\t[Failed].")
                try:
                    print('trying to read the file [%d] %s with the function read_excel %s,' % (
                    idx + 1, sample_file, item),
                          end=' ')
                    xls = pd.read_excel(sample_file)
                    print('\t[OK].')
                    # 涟源2号机xls文件，开头为中英表 标志为“日期”（时间戳为日期和时间两列）
                    # 标记时间戳起始标签
                    time_tag = xls.loc[xls.iloc[:, 0] == "日期"]
                    strat_tag = int(time_tag.index[0]) + 1
                    # 根据检索标签跳过无用行且合并时间列（'日期''时间'列）
                    xls = pd.read_excel(sample_file, parse_dates=[[0, 1]], skiprows=strat_tag, index_col=False)
                    # 删除换行符及后缀
                    xls = xls.replace('\n', ' ', regex=True)
                    xls = xls.replace(' B', '', regex=True)
                    xls = xls.replace(' P', '', regex=True)
                    # xls = xls.replace(' Auto HistorianActual', ' ', regex=True)
                    xls.rename(columns=lambda x: x.replace('\n', ''), inplace=True)
                    # 删除空列
                    xls = xls.T.dropna()
                    xls = xls.T
                    # print(xls.columns.values)
                    xls.rename(columns={'日期_时间': 'Time'}, inplace=True)
                    xls['Time'] = xls['Time'].replace('. 0', '', regex=True)
                    # 将时间列转化为datetime对象
                    xls['Time'] = pd.to_datetime(xls['Time'])
                    # 将时间戳转化为要求格式
                    xls['Time'] = xls['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))
                    xls.set_index('Time', inplace=True)
                    # 批量浮点型转换
                    xls = xls.astype('float64')
                    print("转换后表格:")
                    # 按原名称另存为utf_8格式的csv文件
                    xls.to_csv(output_dir + sample_file[8:-4] + ".csv", encoding='utf-8')
                    print(xls)
                    print(xls.dtypes)
                    break
                except:
                    print("\t[Failed].")
                try:
                    print('trying to read the file [%d] %s with the function read_excel %s,' % (
                    idx + 1, sample_file, item),
                          end=' ')
                    xls = pd.read_excel(sample_file)
                    print('\t[OK].')
                    # 温州机组 内含两张表
                    # 标记时间戳起始标签
                    time_tag = xls.loc[xls.iloc[:, 0] == "Hour"]
                    strat_tag = int(time_tag.index[0]) + 1
                    # Block行作为列名
                    ColNames_List = xls.loc[3].tolist()
                    ColNames_List[1] = 'Time'
                    # 根据检索标签跳过无用行
                    xls = pd.read_excel(sample_file, skiprows=strat_tag, names=ColNames_List, index_col=False)
                    # 删除第一列空值列
                    xls = xls.drop(xls.columns[[0]], axis=1)
                    # 剔除列名中含指定字符的列
                    xls = xls[xls.columns.drop(list(xls.filter(regex='NaT')))]
                    # 删除换行符及后缀
                    xls = xls.replace('\n', ' ', regex=True)
                    xls = xls.replace(' B', '', regex=True)
                    xls = xls.replace(' P', '', regex=True)
                    xls = xls.replace(r'\?', '', regex=True)
                    xls.rename(columns=lambda x: x.replace('\n', ''), inplace=True)
                    xls.rename(columns=lambda x: x.replace('Auto HistorianActual', ''), inplace=True)
                    # 删除空列
                    xls = xls.T.dropna()
                    xls = xls.T
                    # 将时间列转化为datetime对象
                    xls['Time'] = pd.to_datetime(xls['Time'])
                    # 将时间戳转化为要求格式
                    xls['Time'] = xls['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))
                    xls.set_index('Time', inplace=True)
                    # 批量浮点型转换
                    xls = xls.astype('float64')
                    print("转换后表格:")
                    # 按原名称另存为utf_8格式的csv文件
                    xls.to_csv(output_dir + sample_file[8:-4] + ".csv", encoding='utf-8')
                    print(xls)
                    print(xls.dtypes)
                    break
                except:
                    print("\t[Failed].")

        except:
            print(', [Failed]')
            print("  - ERROR: Opening the file: %s, and Skip the file." % (sample_file))
            print(sys.exc_info())
            continue

    # XLSX文件处理
    xlsx_sheets = pd.Series([])
    for (idx, sample_file) in enumerate(xlsx_files):
        try:
            for item in csv_encoding:
                try:
                    print('trying to read the file [%d] %s with the engine "openpyxl",' % (idx + 1, sample_file,),
                          end=' ')
                    # pd rad_excel对xlsx不直接支持，需用特定openpyxl引擎打开
                    # IMP湖北#2机组/阜阳#1机组IMP/广热#1机组IMP
                    xlsx = pd.read_excel(sample_file, engine='openpyxl')
                    # xlsx = pd.read_csv(sample_file)
                    print('\t[OK].')
                    # 标记时间戳起始标签
                    time_tag = xlsx.loc[xlsx.iloc[:, 0] == "编号"]
                    strat_tag = int(time_tag.index[0]) + 2
                    # 重命名中文释义行（补齐第1列空白名称）
                    ColNames_List_old = xlsx.loc[0].tolist()
                    # 根据检索标签跳过无用行
                    xlsx = pd.read_excel(sample_file, skiprows=strat_tag, names=ColNames_List_old, index_col=False)
                    # 删除空列
                    xlsx = xlsx.dropna(how='all', axis=1)
                    xlsx = xlsx.dropna(how='any', axis=0)
                    xlsx.rename(columns={'名称': 'Time'}, inplace=True)
                    # 将时间列转化为datetime对象
                    xlsx['Time'] = pd.to_datetime(xlsx['Time'])
                    # 将时间戳转化为要求格式
                    xlsx['Time'] = xlsx['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))
                    xlsx.set_index('Time', inplace=True)
                    # 批量浮点型转换
                    xlsx = xlsx.astype('float64')
                    print("转换后表格:")
                    # 按原名称另存为gb18030格式的csv文件（针对IMP带中文名称的需求，统一设为gb18030格式）
                    xlsx.to_csv(output_dir + sample_file[8:-5] + ".csv", encoding='gb18030')
                    print(xlsx)
                    print(xlsx.dtypes)
                    break
                except:
                    print("\t[Failed].")
                try:
                    print('trying to read the file [%d] %s with the engine "openpyxl",' % (idx + 1, sample_file,),
                          end=' ')
                    # pd rad_excel对xlsx不直接支持，需用特定openpyxl引擎打开
                    # 赤壁#3机组IMP原始数据
                    xlsx = pd.read_excel(sample_file, engine='openpyxl')
                    print('\t[OK].')
                    # 重命名中文释义行（补齐第1列空白名称）
                    ColNames_List_old = xlsx.loc[6].tolist()
                    ColNames_List_old = str(ColNames_List_old).replace("'", "")
                    ColNames_List_old = str(ColNames_List_old).replace(" ", "")
                    ColNames_List = str(ColNames_List_old)[:-1].split(",")  # [:-2]
                    ColNames_List[0] = 'Time'
                    # 根据检索标签跳过无用行,中文需用gbk格式读取
                    # 此处合并两列时间列
                    xlsx = pd.read_excel(sample_file, skiprows=9, parse_dates=[[0, 1]], names=ColNames_List,
                                         index_col=False, engine='openpyxl')
                    # 删除空行
                    # xlsx = xlsx.dropna(how='any', axis=1)
                    xlsx = xlsx.dropna(how='any', axis=0)
                    # 注意名称前有个空格
                    xlsx.rename(columns={'Time_名称': 'Time'}, inplace=True)
                    # 寻找全是空格或者空值的序号
                    NONE_VIN = (xlsx['Time'].isnull()) | (xlsx['Time'].apply(lambda x: str(x).isspace()))
                    df_null = xlsx[NONE_VIN]
                    # 删除时间戳为空格（非空）的行
                    xlsx = xlsx.drop(df_null['Time'].index)
                    # 删除空行
                    # xlsx = xlsx.dropna(how='any', axis=0)
                    # 后四行时间项为多组空格（非空）转时间戳格式会出错
                    # 将时间列转化为datetime对象
                    xlsx['Time'] = pd.to_datetime(xlsx['Time'])
                    # 将时间戳转化为要求格式
                    xlsx['Time'] = xlsx['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))
                    xlsx.set_index('Time', inplace=True)
                    # 批量浮点型转换
                    xlsx = xlsx.astype('float64')
                    print("转换后表格:")
                    # 按原名称另存为utf-8格式的csv文件（针对IMP带中文名称的需求，统一设为gb18030格式）
                    xlsx.to_csv(output_dir + sample_file[8:-5] + ".csv", encoding='gb18030')
                    print(xlsx)
                    print(xlsx.dtypes)
                    break
                except:
                    print("\t[Failed].")
                try:
                    print('trying to read the file [%d] %s with the engine "openpyxl",' % (idx + 1, sample_file,),
                          end=' ')
                    # pd rad_excel对xlsx不直接支持，需用特定openpyxl引擎打开
                    # 宜昌#1机组
                    xlsx = pd.read_excel(sample_file, engine='openpyxl')
                    print('\t[OK].')
                    xlsx = pd.read_excel(sample_file,
                                         index_col=False, engine='openpyxl')
                    xlsx.rename(columns={'TAGNAME': 'Time'}, inplace=True)
                    # 删除指定行（第二行中文释义）
                    xlsx = xlsx.drop(index=0)
                    xlsx = xlsx.dropna(how='any', axis=0)
                    # 注意名称前有个空格
                    xlsx.rename(columns={'Time_ 名称': 'Time'}, inplace=True)
                    # 寻找全是空格或者空值的序号
                    NONE_VIN = (xlsx['Time'].isnull()) | (xlsx['Time'].apply(lambda x: str(x).isspace()))
                    df_null = xlsx[NONE_VIN]
                    # 删除时间戳为空格（非空）的行
                    xlsx = xlsx.drop(df_null['Time'].index)
                    # 删除空行
                    # xlsx = xlsx.dropna(how='any', axis=0)
                    # 后四行时间项为多组空格（非空）转时间戳格式会出错
                    # 将时间列转化为datetime对象
                    xlsx['Time'] = pd.to_datetime(xlsx['Time'])
                    # 将时间戳转化为要求格式
                    xlsx['Time'] = xlsx['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))
                    xlsx.set_index('Time', inplace=True)
                    # 批量浮点型转换
                    xlsx = xlsx.astype('float64')
                    print("转换后表格:")
                    # 按原名称另存为utf-8格式的csv文件（针对IMP带中文名称的需求，统一设为gb18030格式）
                    xlsx.to_csv(output_dir + sample_file[8:-5] + ".csv", encoding='gb18030')
                    print(xlsx)
                    print(xlsx.dtypes)
                    break
                except:
                    print("\t[Failed].")
                try:
                    print('trying to read the file [%d] %s with the engine "openpyxl",' % (idx + 1, sample_file,),
                          end=' ')
                    # pd rad_excel对xlsx不直接支持，需用特定openpyxl引擎打开
                    # 焦作自拷贝xlsx文件
                    xlsx = pd.read_excel(sample_file, engine='openpyxl')
                    print('\t[OK].')
                    xlsx = pd.read_excel(sample_file,
                                         index_col=False, engine='openpyxl')
                    xlsx.rename(columns={'Timestamp': 'Time'}, inplace=True)
                    xlsx = xlsx.dropna(how='any', axis=0)
                    # 寻找全是空格或者空值的序号
                    NONE_VIN = (xlsx['Time'].isnull()) | (xlsx['Time'].apply(lambda x: str(x).isspace()))
                    df_null = xlsx[NONE_VIN]
                    # 删除时间戳为空格（非空）的行
                    xlsx = xlsx.drop(df_null['Time'].index)
                    # 删除空行
                    # xlsx = xlsx.dropna(how='any', axis=0)
                    # 后四行时间项为多组空格（非空）转时间戳格式会出错
                    # 将时间列转化为datetime对象
                    xlsx['Time'] = pd.to_datetime(xlsx['Time'])
                    # 将时间戳转化为要求格式
                    xlsx['Time'] = xlsx['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))
                    xlsx.set_index('Time', inplace=True)
                    # 批量浮点型转换
                    xlsx = xlsx.astype('float64')
                    print("转换后表格:")
                    # 按原名称另存为utf-8格式的csv文件（针对IMP带中文名称的需求，统一设为gb18030格式）
                    xlsx.to_csv(output_dir + sample_file[8:-5] + ".csv", encoding='gb18030')
                    print(xlsx)
                    print(xlsx.dtypes)
                    break
                except:
                    print("\t[Failed].")
                try:
                    print('trying to read the file [%d] %s with the engine "openpyxl",' % (idx + 1, sample_file,),
                          end=' ')
                    # pd rad_excel对xlsx不直接支持，需用特定openpyxl引擎打开
                    # 板桥#4机组处理后数据（原始数据时间戳没有年月日期）
                    xlsx = pd.read_excel(sample_file, engine='openpyxl')
                    print('\t[OK].')
                    xlsx = pd.read_excel(sample_file,
                                         index_col=False, engine='openpyxl')
                    xlsx.rename(columns={'Timestamp': 'Time'}, inplace=True)
                    xlsx = xlsx.dropna(how='any', axis=0)
                    # 寻找全是空格或者空值的序号
                    NONE_VIN = (xlsx['Time'].isnull()) | (xlsx['Time'].apply(lambda x: str(x).isspace()))
                    df_null = xlsx[NONE_VIN]
                    # 删除时间戳为空格（非空）的行
                    xlsx = xlsx.drop(df_null['Time'].index)
                    # 删除空行
                    # xlsx = xlsx.dropna(how='any', axis=0)
                    # 后四行时间项为多组空格（非空）转时间戳格式会出错
                    # 将时间列转化为datetime对象
                    xlsx['Time'] = pd.to_datetime(xlsx['Time'])
                    # 将时间戳转化为要求格式
                    xlsx['Time'] = xlsx['Time'].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))
                    xlsx.set_index('Time', inplace=True)
                    # 批量浮点型转换
                    # print(xlsx.iloc[2].values)
                    xlsx = xlsx.replace('I/O Timeout', '1', regex=True)
                    xlsx = xlsx.astype('float64')
                    print("转换后表格:")
                    # 按原名称另存为utf-8格式的csv文件（针对IMP带中文名称的需求，统一设为gb18030格式）
                    xlsx.to_csv(output_dir + sample_file[8:-5] + ".csv", encoding='gb18030')
                    print(xlsx)
                    print(xlsx.dtypes)
                    break
                except:
                    print("\t[Failed].")
        except:
            print(', [Failed]')
            print("  - ERROR: Opening the file: %s, and Skip the file." % (sample_file))
            print(sys.exc_info())
            continue
    tkinter.messagebox.showinfo('提示', '转换完成！')

# 下拉框数据获取函数


def get_combobox_parameters(event):
    # 获取选中的值，保存
    combobox_value = parameter_value.get()
    print(combobox_value)

# 预览输出文件夹


def preview_file():
    global final_path
    os.system('start' + r'.\output')

# 导出结果


def download_file():
    user_path = filedialog.askdirectory()
    ouput_files = sorted(glob.glob(output_dir + '*.csv'))
    for (idx, sample_file) in enumerate(ouput_files):
        user_file = user_path + './' + sample_file[8:-4] + ".csv"
        shutil.copyfile(sample_file, user_file)

    text_preview.delete('0.0', 'end')
    text_preview.insert('0.end', "导出成功！")
    tkinter.messagebox.showinfo('提示', '导出成功！')


if __name__ == "__main__":
    # 实例化object，建立窗口window
    window = tk.Tk()
    # 给窗口的可视化起名字
    window.title('DCS数据格式转换工具-V1.12')
    # 设定窗口的大小(长 * 宽)
    window.geometry('1200x600')
    # 窗口默认最大化
    # window.state("zoomed")
    # 创建滑动条
    #scrollbar = tk.Scrollbar(window)
    # 主界面上设置主画布
    canvas_main = tk.Canvas(
        window,
        width=900,
        height=900,
        scrollregion=(
            0,
            0,
            520,
            520))
    # 放置主画布
    canvas_main.place(x=0, y=0, relwidth=1, relheight=1)
    # 创建主框架,把Frame放在canvas里
    frame_main = tk.Frame(canvas_main)
    frame_main.place(x=0, y=0, relwidth=1, relheight=1)
    canvas_main.create_window(0, 0, window=frame_main, anchor='nw')
    # 创建大标题
    label_Title = tk.Label(
        frame_main,
        text='DCS格式转换工具',
        font=(
            'microsoft yahei',
            22,
            "bold"),
        width=20,
        height=2,
        bg='#F5F5F5'). grid(
            row=0,
            column=3,
            padx=0,
        pady=10)
    # 设置分割线1-2
    sep_1 = Separator(
        frame_main,
        orient=tk.HORIZONTAL,
        style='blue.TSeparator')
    sep_1.place(x=0, y=90, relwidth=1)

    # 试验步骤一-试验步骤上传
    label_title1 = tk.Label(
        frame_main,
        text='试验数据上传',
        font=(
            'microsoft yahei',
            15,
            "bold roman"),
        width=20,
        height=2,
        bg='#F0F8FF'). grid(
            row=2,
            column=1,
            padx=0,
        pady=10)
    Label_DCS_1 = tk.Label(
        frame_main,
        text='*原始数据选择:',
        font=(8)).grid(
        row=3,
        column=2,
        padx=0,
        pady=10)

    button_upload = tk.Button(
        frame_main,
        text='上传',
        font=(7),
        command=upload_data,
        width=6).grid(
        row=3,
        column=4,
        padx=5,
        pady=10)
    text_preview = tk.Text(
        frame_main, width=20, height=1, font=(
            'microsoft yahei', 14))
    text_preview.grid(row=3, column=3, padx=5, pady=10)
    text_preview.insert("insert", "当前状态:")
    Label_IMP_1 = tk.Label(
        frame_main,
        text='*DCS设备选择:',
        font=(8)).grid(
        row=4,
        column=2,
        padx=0,
        pady=10)
    # 创建下拉菜单
    parameter_value = tk.StringVar()
    parameter_combobox = ttk.Combobox(
        frame_main, textvariable=parameter_value, font=("", 14))
    # .place(x=476, y=2788, width=300, height=30)
    parameter_combobox.grid(row=4, column=3, padx=0, pady=10)
    # 下拉菜单设定参数,可放于函数内
    value_set = (
        "Ovation",
        "ABB",
        "IMP",
        "Foxborol",
        "日立",
        "浙大中控",
        "国能智深",
        "和利时",
        "SIS数据",
        "其他")
    parameter_combobox["value"] = value_set
    parameter_combobox.current(0)  # 设置缺省参数
    parameter_combobox.bind(
        "<<ComboboxSelected>>",
        get_combobox_parameters)  # #给下拉菜单绑定事件

    button_start_processing = tk.Button(
        frame_main,
        text='开始转换',
        font=(7),
        command=start_processing,
        width=9).grid(
        row=4,
        column=4,
        padx=0,
        pady=10)

    # 增加空栏隔断,拓宽主页面页宽
    label_space1 = tk.Label(
        frame_main,
        text='\n',
        font=(
            'Arial',
            12,
            "bold roman"),
        width=50,
        height=1).grid(
            row=7,
            column=6,
            padx=0,
        pady=10)
    # 设置分割线1-2
    #sep_1 = Separator(frame_main, orient=tk.HORIZONTAL, style='blue.TSeparator')
    #sep_1.place(x=0, y=420, relwidth=1)

    # 试验步骤二-结果预览与导出
    label_title2 = tk.Label(
        frame_main,
        text='预览与导出',
        font=(
            'microsoft yahei',
            15,
            "bold roman"),
        width=20,
        height=2,
        bg='#F0F8FF'). grid(
            row=6,
            column=1,
            padx=0,
        pady=10)
    button_preview_file = tk.Button(
        frame_main,
        text='结果预览',
        font=(7),
        command=preview_file,
        width=9).grid(
        row=7,
        column=3,
        padx=0,
        pady=10)
    button_download_file = tk.Button(
        frame_main,
        text='导出',
        font=(7),
        command=download_file).grid(
        row=7,
        column=4,
        padx=0,
        pady=10)
    # 获取当前桌面地址
    global desktop_path
    desktop_path = os.path.join(os.path.expanduser("~"), 'Desktop')

    window.update()

    window.mainloop()
