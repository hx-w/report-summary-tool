import re
import os
import logging
import yaml

logger = logging.getLogger()


class Config(object):
    def __init__(self) -> None:
        self.safemode = True
        self.debug = False
        self.workspace = '06_工作周报'
        self.summary = '00_汇总'
        self.group = []
        self.prefix = '通用网格生成软件个人工作周报'
        self.group_pattern = '0.+_.+组'
        self.newest_week = ''
        self.week_pattern = '\d{4}-\d{4}-\d{4}'
        self.week_number = 'newest'
        logging.basicConfig(
            format='[%(levelname)s] %(message)s',
            level=logging.INFO
        )

    def __check(self) -> bool:
        logger.debug('正在检查工作目录')
        current_dir = os.getcwd()
        logger.info(f'当前运行目录 "{current_dir}"')
        return current_dir[-len(self.workspace):] == self.workspace

    def __gen_group(self):
        pattern = re.compile(self.group_pattern)
        self.group = list(filter(
            lambda x: re.match(pattern, x),
            os.listdir('.')
        ))
        logger.debug(f'小组名称获取结束 {self.group}')

    def load_config(self, filename: str) -> None:
        try:
            with open(filename, 'r', encoding='utf-8') as ifile:
                config = yaml.safe_load(ifile)
                self.safemode = config['strategy']['safemode']
                self.debug = config['strategy']['debug']
                self.workspace = config['file_structure']['workspace']
                self.summary = config['file_structure']['summary']
                self.group = config['file_structure']['group']
                self.prefix = config['file_prefix']
                self.group_pattern = config['group_pattern']
                self.week_pattern = config['week_pattern']
                self.week_number = config['week_number']
            logger.info('配置文件读取成功')
            
            if self.debug:
                logger.setLevel(logging.DEBUG)
            else:
                logger.setLevel(logging.INFO)
            
            if self.safemode:
                if not self.__check():
                    logger.error(f'已开启safemode，工作目录与"{self.workspace}"不同')
                    input('运行结束')
                    exit()
            else:
                logger.debug('未开启safemode，跳过工作目录检查，开始获取小组名称')
                self.__gen_group()
            
            if len(self.group) == 0:
                logger.error('配置错误？小组数量为0')
                input('运行结束')
                exit()
            else:
                self.group.sort()

            if self.week_number == 'newest':
                logger.debug('week_number = newest; 获取最新周')
            elif re.match(self.week_pattern, self.week_number):
                logger.debug(f'手动设置周号：{self.week_number}')
                self.newest_week = self.week_number
            else:
                logger.error(f'{self.week_number} 不满足pattern {self.week_pattern}')
                input('运行结束')
                exit()
        except Exception as ept:
            logger.error(f'读取配置文件失败 <{ept}>')
            input('运行结束')
            exit()

global_config = Config()