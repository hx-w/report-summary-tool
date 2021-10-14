# -*- coding: utf-8 -*-

from typedef import *
import pandas as pd

def get_week_list(group_name: str) -> list:
    source_dir = os.path.join('.', group_name)
    pattern = re.compile(global_config.week_pattern)
    week_list = list(filter(
        lambda x: re.match(pattern, x),
        os.listdir(source_dir)
    ))
    week_list.sort(reverse=True)
    return week_list


def judge_sisg_result_exist(group_name: str, week: str) -> bool:
    source_dir = os.path.join('.', global_config.summary)
    result_file = os.path.join(
        source_dir,
        f'{global_config.prefix}小组工作周报-{week}-{group_name[3:]}.xlsx'
    )
    return os.path.exists(result_file)


def exec_sisg_merge(group_name: str, week: str):
    source_dir = os.path.join(
        '.', os.path.join(group_name, week)
    )
    pattern = re.compile(f'{global_config.prefix}个人工作周报-{week}-(.*).xlsx')
    filelist = list(filter(
        lambda x: re.match(pattern, x),
        os.listdir(source_dir)
    ))
    dist_path = os.path.join(
        global_config.summary,
        f'{global_config.prefix}小组工作周报-{week}-{group_name[3:]}.xlsx'
    )
    
    df_list = []
    for eachfile in filelist:
        name = re.findall(pattern, eachfile)[0]
        full_path = os.path.join(source_dir, eachfile)
        df = pd.read_excel(full_path)
        df.insert(1, '姓名', name)
        df_list.append(df)
    new_df = pd.concat(df_list)
    new_df.reset_index(drop=True, inplace=True)
    new_df['序号'] = pd.Series(list(range(1, len(new_df) + 1)))
    new_df['日期（年/月/日）'] = new_df['日期（年/月/日）'].dt.strftime('%Y/%m/%d')
    new_df.to_excel(dist_path, '工作任务项', index=False)
    logger.info(f'已经合并到文件 "{dist_path}"')


def judge_sisg_empty(group_name: str, week: str) -> bool:
    source_dir = os.path.join(
        '.', os.path.join(group_name, week)
    )
    pattern = re.compile(f'{global_config.prefix}个人工作周报-{week}-(.*).xlsx')
    filelist = list(filter(
        lambda x: re.match(pattern, x),
        os.listdir(source_dir)
    ))
    return len(filelist) == 0