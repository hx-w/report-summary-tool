# -*- coding: utf-8 -*-

import pandas as pd

source_path = '01_几何组/2021-1014-1020/通用网格生成软件个人工作周报-2021-1014-1020-张三.xlsx'

df = pd.read_excel(source_path, sheet_name='工作任务项')

pivot_table = pd.pivot_table(df, values=['工作量（小时）'], columns=['日期（年/月/日）'], index=['工作任务项'], fill_value='', margins=True, margins_name='总计', aggfunc='sum')

pivot_table.to_excel(source_path, sheet_name='test')