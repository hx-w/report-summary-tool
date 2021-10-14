# -*- coding: utf-8 -*-

import os
import pandas as pd
from typedef import *


def merge_single_group(opt: int):
    source_path = os.path.join(
        os.path.join('.', global_config.group[opt - 1]),
        global_config.newest_week
    )
    dist_filename = '-'.join([
        global_config.prefix,
        global_config.newest_week,
        global_config.group[opt - 1][3:]
    ]) + '.xlsx'
    dist_path = os.path.join(global_config.summary, dist_filename)

    allfiles = list(filter(
        lambda x: x.startswith(global_config.prefix),
        os.listdir(source_path)
    ))

    df_list = []
    for eachfile in allfiles:
        name = eachfile.split('-')[-1].split('.')[0]
        full_path = os.path.join(source_path, eachfile)
        df = pd.read_excel(full_path)
        df.insert(1, '姓名', name)
        df_list.append(df)
    
    new_df = pd.concat(df_list)
    new_df.reset_index(drop=True, inplace=True)
    new_df['序号'] = pd.Series(list(range(1, len(new_df) + 1)))
    new_df['日期（年/月/日）'] = new_df['日期（年/月/日）'].dt.strftime('%Y/%m/%d')
    new_df.to_excel(dist_path, '工作任务项', index=False)
    logger.info(f'已经合并到文件 "{dist_path}"')

def merge_all_group():
    print('')
    pattern = re.compile(f'{global_config.prefix}-{global_config.newest_week}-(.*组).xlsx')
    source_dir = os.path.join('.', global_config.summary)
    allfiles = list(filter(
        lambda x: re.match(pattern, x),
        os.listdir(source_dir)
    ))
    allgroup = list(map(
        lambda x: re.findall(pattern, x)[0],
        allfiles
    ))
    allgroup_off = list(map(
        lambda x: x[3:],
        global_config.group
    ))
    allgroup.sort()
    allgroup_off.sort()
    if allgroup == allgroup_off:
        logger.info('该周的单组的汇总文件完整，可以开始汇总')
    elif not global_config.safemode:
        logger.warn(f'汇总文件夹下该周的单组汇总文件包括：{allgroup}')
        logger.warn(f'所有小组为： {allgroup_off}')
        logger.warn('二者不相同，由于未开启safemode将继续汇总')
    else:
        logger.error(f'汇总文件夹下该周的单组汇总文件包括：{allgroup}')
        logger.error(f'所有小组为： {allgroup_off}')
        logger.error('二者不相同，由于开启safemode，程序终止')
        input('运行结束')
        exit()
        
    dist_path = os.path.join(source_dir, f'{global_config.prefix}-{global_config.newest_week}.xlsx')
    df_list = []
    for eachfile in allfiles:
        name = eachfile.split('-')[-1].split('.')[0]
        full_path = os.path.join(source_dir, eachfile)
        df = pd.read_excel(full_path)
        df.insert(1, '组名', name)
        df_list.append(df)
    
    new_df = pd.concat(df_list)
    new_df.reset_index(drop=True, inplace=True)
    new_df['序号'] = pd.Series(list(range(1, len(new_df) + 1)))
    # new_df['日期（年/月/日）'] = new_df['日期（年/月/日）'].dt.strftime('%Y/%m/%d')
    new_df.to_excel(dist_path, '工作任务项', index=False)
    logger.info(f'已经合并到文件 "{dist_path}"')