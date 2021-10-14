# -*- coding: utf-8 -*-

import os
from re import M
from typedef import *

def intopt_check(opt: str, imin: int, imax: int) -> int:
    try:
        intopt = int(opt)
        assert intopt >= imin and intopt <= imax
        return intopt
    except Exception as ept:
        logger.error(f'非法输入 {ept}')
        input('运行结束')
        exit()


def menu_level_1() -> int:
    print('\n======================')
    print('1. 汇总单组表单（最近周）')
    print('2. 汇总所有表单（最近周）')
    print('输入 1 ~ 2')
    opt = input('>> ').strip()
    return intopt_check(opt, 1, 2)


def menu_level_2() -> int:
    print('\n======================')
    print('选择需要汇总的小组编号：')
    for idx in range(len(global_config.group)):
        print(f'{idx + 1}. {global_config.group[idx]}')
    print(f'输入 1 ~ {len(global_config.group)}')
    opt = input('>> ').strip()
    print('')
    return intopt_check(opt, 1, len(global_config.group))


def fetch_newest_week(opt: int, dir: str):
    if global_config.week_number != 'newest': return
    logger.debug('正在获取最新周')
    if opt == 1:
        pattern = re.compile(global_config.week_pattern)
        fullpath = os.path.join('.', dir)
        allfiles = list(filter(
            lambda x: re.match(pattern, x),
            os.listdir(fullpath)
        ))
        if len(allfiles) == 0:
            logger.error(f'该目录下没有周报文件夹 "{fullpath}"')
            input('运行结束')
            exit()
        
        allfiles.sort(reverse=True)
        logger.info(f'最新周为 {allfiles[0]}')
        global_config.newest_week = allfiles[0]
    elif opt == 2:
        pattern = re.compile(f'{global_config.prefix}-({global_config.week_pattern})-.*组.xlsx')
        fullpath = os.path.join('.', global_config.summary)
        allfiles = list(filter(
            lambda x: re.match(pattern, x),
            os.listdir(fullpath)
        ))
        if len(allfiles) == 0:
            logger.error(f'汇总目录下没有小组的周报，无法汇总总周报')
            input('运行结束')
            exit()
        
        allfiles.sort(reverse=True)
        newweek = re.findall(pattern, allfiles[0])[0]
        logger.info(f'最新周为 {newweek}')
        global_config.newest_week = newweek