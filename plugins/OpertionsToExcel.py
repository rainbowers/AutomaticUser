# -*- coding:utf-8 -*-
__author__ = 'Rainbower'

import xlrd
from xlrd import xldate_as_tuple
import datetime
'''
xlrd中单元格的数据类型
数字一律按浮点型输出，日期输出成一串小数，布尔型输出0或1，所以我们必须在程序中做判断处理转换
成我们想要的数据类型
0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
'''
class ExcelData():
    # 初始化方法
    def __init__(self, data_path, sheetname):
        self.data_path = data_path
        self.sheetname = sheetname
        self.data = xlrd.open_workbook(self.data_path)
        self.table = self.data.sheet_by_name(self.sheetname)
        self.keys = self.table.row_values(0)
        self.rowNum = self.table.nrows
        self.colNum = self.table.ncols


    def readExcel(self):
        datas = []
        for i in range(1, self.rowNum):
            sheet_data = {}
            for j in range(self.colNum):
                c_type = self.table.cell(i,j).ctype
                c_cell = self.table.cell_value(i, j)
                if c_type == 2 and c_cell % 1 == 0:
                    c_cell = int(c_cell)
                elif c_type == 3:
                    date = datetime.datetime(*xldate_as_tuple(c_cell,0))
                    c_cell = date.strftime('%Y/%d/%m %H:%M:%S')
                elif c_type == 4:
                    c_cell = True if c_cell == 1 else False
                sheet_data[self.keys[j]] = c_cell
            datas.append(sheet_data)
        return datas


if __name__=='__main__':
    pass