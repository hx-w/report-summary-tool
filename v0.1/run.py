from typedef import *
from interactor import *
from excel import *


def init_config(filename: str):
    global global_config
    global_config.load_config(filename)


if __name__ == '__main__':
    init_config('config.yml')
    opt1 = menu_level_1()
    if opt1 == 1:
        opt2 = menu_level_2()
        fetch_newest_week(opt1, global_config.group[opt2 - 1])
        merge_single_group(opt2)
    elif opt1 == 2:
        fetch_newest_week(opt1, '')
        merge_all_group()
    else:
        logger.error('未知错误')
        input('运行结束')
        exit()