#!/usr/env python
# coding: utf-8
import xlrd
from collections import Counter
import glob
import os
import sys


def get_week_day(date):
    weekday = date.weekday()
    return "星期" + str(weekday + 1)


def open_excel(file_name='1.xlsx'):
    # type: (file) -> object
    """

    :type file_name: file = 1.xlsx
    """
    try:
        data = xlrd.open_workbook(file_name)
        return data
    except Exception, e:
        print str(e)


# 根据表索引
def excel_table_byindex(file_name='1.xlsx', by_index=0):
    data = open_excel(file_name)
    table = data.sheets()[by_index]
    nrows = table.nrows  # rows
    lists = []
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        data1 = xlrd.xldate.xldate_as_datetime(table.cell(0, 0).value, 0)
        # if row:
        #     app = { }
        #     for i in range(len(colnames)):
        #         app[colnames[i]] = row[i]
        assert isinstance(data1, basestring)
        lists.append(data1)
    return lists


# 根据表名
def excel_table_byname(file='1.xlsx', by_name=u'Sheet1'):
    """

    :type by_name: sheet1
    """
    data = open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows
    for i in range(nrows):
        l1 = table.row_values(i)
        list.append(l1)
    return list


def summary_week(lists):
    res = dict(Counter(lists))
    return res


def main():
    # tables = excel_table_byindex()

    datas = []
    for fileName in read_all_xlsx():
        datas.append(open_excel(fileName))

    weeks = []
    for data in datas:
        table = data.sheet_by_index(0)
        for row in range(table.nrows):
            data1 = xlrd.xldate.xldate_as_datetime(table.cell(row, 0).value, 0)
            weeks.append(get_week_day(data1))
    return summary_week(weeks)
    # tables1 = excel_table_byname()
    # for row in tables1:
    #     print row


# 读取所有的xlsx文件
def read_all_xlsx():
    tmp_path = "{}/{}".format(cur_path(), "*.xlsx")

    all_file_name = glob.glob(tmp_path)
    return all_file_name


# 判断当前路径
def cur_path():
    tmp_path = sys.argv[0]
    if os.path.isdir(tmp_path):
        return tmp_path
    elif os.path.isfile(tmp_path):
        return os.path.dirname(tmp_path)


if __name__ == '__main__':
    # print get_week_day(datetime.datetime.now())
    dict1 = main()
    for _key in dict1:
        print _key, dict1[_key]
    read_all_xlsx()
