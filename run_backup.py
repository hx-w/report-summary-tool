# -*- coding: utf-8 -*-
import os
import pandas as pd

WORKSPACE = '06_工作周报'


def check_workspace() -> bool:
    current_dir = os.getcwd()
    print(f'[INFO] 当前运行目录：{current_dir}')
    return current_dir[-len(WORKSPACE):] == WORKSPACE


def merge_exec(source_dir: str, dist_file: str):
    all_files = os.listdir(source_dir)
    filter_files = list(filter(lambda x: len(x.split('-')) == 5, all_files))
    df_list = []
    for eachfile in filter_files:
        name = eachfile.split('-')[-1].split('.')[0]
        full_path = os.path.join(source_dir, eachfile)
        df = pd.read_excel(full_path)
        df.insert(1, '姓名', name)
        df_list.append(df)
    new_df = pd.concat(df_list)
    new_df.reset_index(drop=True, inplace=True)
    new_df['序号'] = pd.Series(list(range(1, len(new_df) + 1)))
    new_df['日期（年/月/日）'] = new_df['日期（年/月/日）'].dt.strftime('%Y/%m/%d')
    dist_file = os.path.join('00_汇总', dist_file)
    new_df.to_excel(dist_file, '工作任务项')
    print('Done.')


def merge_one_group(group_name: str):
    print('===================')
    all_files = os.listdir(os.path.join('.', group_name))
    filter_files = list(filter(lambda x: x.startswith('20'), all_files))
    filter_files.sort(reverse=True)
    if len(filter_files) == 0:
        print(f'[WARN] {group_name} 中没有找到周报文件夹')
        input('运行结束')
        exit()
    newest = filter_files[0]
    print(f'[INFO] "{group_name}" 中最新的周报文件夹为："{newest}"，即将对其进行合并\n\n')
    
    source_dir = os.path.join(os.path.join('.', group_name), newest)
    all_files = os.listdir(source_dir)
    filter_files = list(filter(lambda x: len(x.split('-')) == 5, all_files))
    if len(filter_files) == 0:
        print(f'[WARN] {group_name}/{newest} 该目录下没有周报')
        input('运行结束')
        exit()
    file_prefix = filter_files[0].split('-')[0]
    dist_file_name = '-'.join([file_prefix, newest, group_name[3:]])
    merge_exec(os.path.join(os.path.join('.', group_name), newest), dist_file_name + '.xlsx')


def op_merge_single_group():
    print('===================')
    all_files = os.listdir('.')
    options = list(filter(lambda x: x.startswith('0')
                   and x[-1] == '组', all_files))
    if len(options) == 0:
        print('[ERROR] 当前文件夹没有找到小组文件夹')
        input('运行结束')
        exit()

    options.sort()
    message = '\n选择需要汇总的小组编号：\n'
    for idx in range(len(options)):
        message += f'{idx + 1}. {options[idx]}\n'

    message += f'（输入{1} ~ {len(options)}）\n>> '
    opt = input(message).strip()
    try:
        merge_one_group(options[int(opt) - 1])
    except Exception as ept:
        print(f'[ERROR] 非法输入 {ept}')


def op_merge_all_group():
    print('===================')
    
    pass


if __name__ == '__main__':
    if not check_workspace():
        print(f'[ERROR] 工作目录不正确，请将脚本放在"{WORKSPACE}"下运行')
        input('运行结束')
        exit()

    opt = input('\n1. 汇总单组表单\n2. 汇总所有组表单\n（输入1或2）\n>> ').strip()
    funcs = [op_merge_single_group, op_merge_all_group]
    try:
        intopt = int(opt)
        assert intopt >= 1 and intopt <= len(funcs)
        funcs[intopt - 1]()
    except Exception as ept:
        print(f'\n[ERROR] 非法输入 {ept}')
