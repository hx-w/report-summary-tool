# -*- coding: utf-8 -*-

from shutil import copyfile
import xlrd
import xlwt
from xlutils.copy import copy


# source_path = '01_几何组/2021-1014-1020/通用网格生成软件个人工作周报-2021-1014-1020-张三.xls'

# dist_path = 'tests/test.xls'

# copyfile(source_path, dist_path)

# rbook = xlrd.open_workbook(dist_path, formatting_info=True)

# wbook = copy(rbook)

# wbook.save(dist_path)


import openpyxl


excel_writer = "tests/test_xlsx_test.xlsx"
wb = openpyxl.load_workbook(excel_writer)  # 打开excel文件
pivot_sheet = wb[wb.sheetnames[1]]  # 打开指定Sheet
pivot = pivot_sheet._pivots[0]  # 任何一个都可以共享同一个缓存

pivot.cache.cacheSource.worksheetSource.ref='A1:G10'
pivot.cache.refreshOnLoad = True  # 刷新加载

wb.save(excel_writer)  # 保存