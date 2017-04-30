#!/usr/env python
#coding: utf-8
import xlrd

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
        data1 = xlrd.xldate.xldate_as_datetime(table.cell(0,0).value,1)
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

def main():
    tables = excel_table_byindex()
    for row in tables:
        print row

    # tables1 = excel_table_byname()
    # for row in tables1:
    #     print row

if __name__=='__main__':
    # print get_week_day(datetime.datetime.now())
    main()