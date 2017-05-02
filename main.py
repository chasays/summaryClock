#!/usr/env python
#coding: utf-8
import xlrd
from collections import Counter
import bchart
def get_week_day(date):
    weekday = date.weekday()
    return "星期" + str(weekday+1)

def open_excel(file='1.xlsx'):
    try:
        data = xlrd.open_workbook(file)
        return  data
    except Exception,e:
        print str(e)
#根据表索引
def excel_table_byindex(file='1.xlsx', colnameindex=0, by_index=0):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #rows
    ncols = table.ncols #cols
    colnames = table.row_values(colnameindex) #某一行数据
    list = []
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        data1 = xlrd.xldate.xldate_as_datetime(table.cell(0,0).value,0)
        # if row:
        #     app = { }
        #     for i in range(len(colnames)):
        #         app[colnames[i]] = row[i]
        list.append(data1)
    return list

#根据表名
def excel_table_byname(file='1.xlsx', colnameindex =0, by_name=u'Sheet1'):
    data = open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows
    ncols = table.ncols
    colnames = table.row_values(colnameindex)
    list = []
    for i in range(nrows):
        l1 = table.row_values(i)
        list.append(l1)
    # for rownum in range(1, nrows):
    #     row = table.row_values(rownum)
    #     if row:
    #         app = { }
    #         for i in range(len(colnames)):
    #             app[colnames[i]] == row[i]
    #         list.append(app)
    return list


def summaryWeek(lists):
    res = dict(Counter(lists))
    return res


def showCharts(dict):
    options = {
        'chart': {'zoomType': 'xy'},
        'title': {'text': '统计哪天容易出现忘打卡'},
        'subtitle': {'text': '数据来源:每月月末考勤数据'},
        'xAxis': {'categories': ['week']},
        'yAxis': {'tilte': {'text': 'Times'}},
        'plotOptions': {'column': {'dataLabels': {'enbaled': True}}}
    }

def main():
    # tables = excel_table_byindex()
    data = open_excel()
    table = data.sheet_by_index(0)
    weeks = []
    for row in range(table.nrows):
        data1 = xlrd.xldate.xldate_as_datetime(table.cell(row, 0).value, 0)
        weeks.append(get_week_day(data1))
    return summaryWeek(weeks)
        # tables1 = excel_table_byname()
        # for row in tables1:
        #     print row

if __name__=='__main__':
    # print get_week_day(datetime.datetime.now())
    dict1 = main()
    from bchart import *

    options = {"width": 500, "height": 500}
    chart = AreaChart("#vis", options)
    chart.load([['group1', '34', '54', '33'], ['group2', '53', '44', '65']]).legend('display', 'none').background(
        'color', "#ffffff").colors(["#dd00dd", '#ffdd00']).option({"isStack": "true"})
    chart.to_json()
