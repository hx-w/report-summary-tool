# -*- coding: utf-8 -*-

from typedef import *
import pandas as pd


def get_swsg_week_list(group_name: str) -> list:
    source_dir = os.path.join('.', group_name)
    pattern = re.compile(global_config.week_pattern)
    week_list = list(filter(
        lambda x: re.match(pattern, x),
        os.listdir(source_dir)
    ))
    week_list.sort(reverse=True)
    return week_list


def get_swmg_week_list(reverse: bool = True, sw=False) -> list:
    source_dir = os.path.join('.', global_config.summary)
    pattern = re.compile(global_config.week_pattern)
    week_list = list(filter(
        lambda x: re.match(pattern, x),
        os.listdir(source_dir)
    ))
    week_list.sort(reverse=reverse)
    if not sw: return week_list
    week_list = list(filter(
        lambda x: os.path.exists(os.path.join(
            os.path.join(source_dir, x),
            f'{global_config.prefix}项目工作周报-{x}.xlsx'
        )),
        week_list
    ))
    return week_list

def get_week_count(start_week: str, end_week: str) -> int:
    week_list = get_swmg_week_list(False)
    return week_list.index(end_week) - week_list.index(start_week) + 1


def judge_swsg_result_exist(group_name: str, week: str) -> bool:
    source_dir = os.path.join('.', global_config.summary)
    result_file = os.path.join(
        source_dir,
        os.path.join(week, f'{global_config.prefix}小组工作周报-{week}-{group_name[3:]}.xlsx')
    )
    return os.path.exists(result_file)

def judge_swmg_result_exist(week: str) -> bool:
    source_dir = os.path.join('.', global_config.summary)
    result_file = os.path.join(
        source_dir,
        os.path.join(week, f'{global_config.prefix}项目工作周报-{week}.xlsx')
    )
    return os.path.exists(result_file)


def judge_swmg_group_complete(week: str) -> bool:
    trim_group = list(map(lambda x: x[3:], global_config.group))
    trim_group.sort()
    pattern = re.compile(f'{global_config.prefix}小组工作周报-{week}-(.*组).xlsx')
    source_dir = os.path.join('.', os.path.join(global_config.summary, week))
    group_list = list(filter(
        lambda x: re.match(pattern, x),
        os.listdir(source_dir)
    ))
    group_list = list(map(
        lambda x: re.findall(pattern, x)[0],
        group_list
    ))
    group_list.sort()
    return trim_group == group_list


def exec_swsg_merge(group_name: str, week: str) -> list:
    source_dir = os.path.join(
        '.', os.path.join(group_name, week)
    )
    pattern = re.compile(f'{global_config.prefix}个人工作周报-{week}-(.*).xlsx')
    filelist = list(filter(
        lambda x: re.match(pattern, x),
        os.listdir(source_dir)
    ))
    dist_dir = os.path.join(global_config.summary, week)
    if not os.path.exists(dist_dir):
        logger.debug(f'目录 "{dist_dir}" 不存在，已创建')
        os.makedirs(dist_dir)
    dist_path = os.path.join(
        dist_dir,
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
    return dist_path


def exec_swmg_merge(week: str) -> list:
    dist_dir = os.path.join(global_config.summary, week)
    # if not os.path.exists(dist_dir):
    #     logger.debug(f'目录 "{dist_dir}" 不存在，已创建')
    #     os.makedirs(dist_dir)
    dist_path = os.path.join(
        dist_dir,
        f'{global_config.prefix}项目工作周报-{week}.xlsx'
    )
    pattern = re.compile(f'{global_config.prefix}小组工作周报-{week}-(.*组).xlsx')
    filelist = list(filter(
        lambda x: re.match(pattern, x),
        os.listdir(dist_dir)
    ))
    df_list = []
    for eachfile in filelist:
        group_name = re.findall(pattern, eachfile)[0]
        full_path = os.path.join(dist_dir, eachfile)
        df = pd.read_excel(full_path)
        df.insert(1, '组名', group_name)
        df_list.append(df)
    new_df = pd.concat(df_list)
    new_df.reset_index(drop=True, inplace=True)
    new_df['序号'] = pd.Series(list(range(1, len(new_df) + 1)))
    # new_df['日期（年/月/日）'] = new_df['日期（年/月/日）'].dt.strftime('%Y/%m/%d')
    new_df.to_excel(dist_path, '工作任务项', index=False)
    logger.info(f'已经合并到文件 "{dist_path}"')

    return dist_path


def judge_swsg_empty(group_name: str, week: str) -> bool:
    source_dir = os.path.join(
        '.', os.path.join(group_name, week)
    )
    pattern = re.compile(f'{global_config.prefix}个人工作周报-{week}-(.*).xlsx')
    filelist = list(filter(
        lambda x: re.match(pattern, x),
        os.listdir(source_dir)
    ))
    return len(filelist) == 0

def judge_swmg_empty(week: str) -> bool:
    source_dir = os.path.join(
        os.path.join('.', global_config.summary), week
    )
    # print(os.listdir(source_dir))
    return not (os.path.exists(source_dir) and len(os.listdir(source_dir)) > 0)

def get_swsg_preview(group_name: str, week: str) -> str:
    source_dir = os.path.join(
        '.', os.path.join(group_name, week)
    )
    pattern = re.compile(f'{global_config.prefix}个人工作周报-{week}-(.*).xlsx')
    filelist = list(filter(
        lambda x: re.match(pattern, x),
        os.listdir(source_dir)
    ))
    filelist = list(map(
        lambda x: re.findall(pattern, x)[0],
        filelist
    ))
    info_list = []
    for idx in range(0, len(filelist), 3):
        info_list.append((' ' * 4).join(filelist[idx:idx + 3]))
    return '\n\n'.join(info_list)


def get_swmg_preview(week: str) -> str:
    source_dir = os.path.join(
        os.path.join('.', global_config.summary), week
    )
    pattern = re.compile(f'{global_config.prefix}小组工作周报-{week}-(.*组).xlsx')
    filelist = list(filter(
        lambda x: re.match(pattern, x),
        os.listdir(source_dir)
    ))
    filelist = list(map(
        lambda x: re.findall(pattern, x)[0],
        filelist
    ))
    info_list = []
    for idx in range(0, len(filelist), 2):
        info_list.append((' ' * 3).join(filelist[idx:idx + 2]))
    return '\n\n'.join(info_list)
