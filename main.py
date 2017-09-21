#!/usr/env python
# coding: utf-8
import xlrd
import sys

reload(sys)
sys.setdefaultencoding('utf-8')


def get_week_day(date):
    weekday = date.weekday()
    return "星期" + str(weekday + 1)


def open_excel(file='1.xlsx'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception, e:
        print str(e)


# 根据表索引
def excel_table_byindex(file='1.xlsx', colnameindex=0, by_index=0):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows  # rows
    ncols = table.ncols  # cols
    colnames = table.row_values(colnameindex)  # 某一行数据
    list = []
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        data1 = xlrd.xldate.xldate_as_datetime(table.cell(0, 0).value, 1)
        # if row:
        #     app = { }
        #     for i in range(len(colnames)):
        #         app[colnames[i]] = row[i]
        list.append(data1)
    return list


# 根据表名
def excel_table_byname(file='1.xlsx', colnameindex=0, by_name=u'Sheet1'):
    data = open_excel(file)
    table = data.sheet_by_name(by_name)
    # assert isinstance(table.nrows, basestring)
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


def PTGames():
    '''
     "img": "/res/images/ptgames/pt_rodz.jpg",
    "ename": "American Roulette",
    "name": "美式轮盘",
    "gamecode": "rodz"

    '''

    tables = excel_table_byname()
    for list in zip(tables):
        print list[0]

        with open(r'pt.js', 'a') as f:
            if list[0][1] != 'img':
                f.write('{\n')
                images = list[0][1]
                images = images.replace(' ', '-')
                print images
                f.write('"img": "/res/images/ptgames/{}.jpg",\n'.format(images))
                f.write('"ename": "{}",\n'.format(list[0][1]))
                f.write('"name": "{}",\n'.format(list[0][0]))
                f.write('"gamecode": "{}"\n'.format(list[0][6]))
                f.write('},\n')


def MgGames():
    '''
    {
        "img": "/res/images/mggames/BTN_LuckyFirecracker.png",
        "ename": "Lucky Firecracker",
        "name": "招财鞭炮",
        "pid": "1126",
        "appid": "1001",
        "gamecode": "LuckyFirecracker"
    },
    '''
    tables = excel_table_byname()
    for list in zip(tables):
        print list[0]

        with open(r'mg.js', 'a') as f:
            f.write('{\n')
            image = list[0][1]
            enname = list[0][2]
            code = list[0][2]
            code = code.replace(' ', '')
            print code
            f.write('"img": "/res/images/mg/{}.png",\n'.format(image))
            f.write('"ename": "{}",\n'.format(list[0][1]))
            f.write('"name": "{}",\n'.format(list[0][0]))
            f.write('"pid": "{}",\n'.format(str(list[0][4])[:-2]))
            f.write('"appid": "{}",\n'.format(str(1002)))
            f.write('"gamecode": "{}"\n'.format(code))
            f.write('},\n')


def getMGGamecodes(str):
    pass


def TTGGames():
    '''
     "img": "/images/ttg/hot_slot_game/1061.png",
        "gamename": "HotVolcanoH5",
        "name": "炽热火山",
        "gametype": "0",
        "gamecode": "1061"
    :return:
    '''
    tables = excel_table_byname()
    for list in zip(tables):
        print list[0]

        with open(r'ttg.js', 'a') as f:
            f.write('{\n')
            gamename = list[0][3]
            name = list[0][1]
            code = int(list[0][2])
            f.write('"img": "http://ams2-games.ttms.co/player/assets/images/games/{}.png",\n'.format(code))
            f.write('"gamename": "{}",\n'.format(gamename))
            f.write('"name": "{}",\n'.format(name))
            f.write('"gametype": "0",\n')
            f.write('"gamecode": "{}"\n'.format(code))
            f.write('},\n')


def batchDownloadTTGPNG(list):
    # http://ams2-games.ttms.co/player/assets/images/games/1020.png
    import requests
    for i in list:
        url = "http://ams2-games.ttms.co/player/assets/images/games/{}.png".format(i)
        print url
        try:
            pic = requests.get(url, timeout=10)
            print pic
        except requests.exceptions.ConnectionError:
            print '【错误】当前图片无法下载'
            continue
        string = 'ttg\\' + str(i) + '.jpg'
        with open(string, 'wb') as fp:
            fp.write(pic.content)
            fp.close()


def modifyMGGames(file):
    with open(file, 'r') as fs:
        lines = fs.readlines()
        for line in lines:
            with open(r'mg1.js', 'a') as f:
                # if line.rfind('{')!=-1:
                #     f.write(line)
                # if line.rfind('"img":')!=-1:
                #     f.write(line)
                if line.rfind('"ename') != -1:
                    # f.write(line)
                    lists = line.split('"')
                    name = lists[3]
                    name = name.replace(' ', '')
                    code = name
                # if line.rfind('"name"')!=-1:
                #     f.write(line)
                # if line.rfind('"pid"')!=-1:
                #     f.write(line)
                # if line.rfind('"appid":')!=-1:
                #     f.write(line)
                if line.rfind('"gamecode":') != -1:
                    # f.write(line)
                    line1 = line[:-2]
                    line2 = line1 + code + '"\n'
                    f.write(line2)
                else:
                    f.write(line)
                    # if line.rfind('"},')!=-1:
                    # f.write(line)


if __name__ == '__main__':
    # print get_week_day(datetime.datetime.now())
    # PTGames()
    # modifyMGGames('modifyMG.js')
    # TTGGames()
    picLists = [14472, 14477, 14446, 14471, 14481, 14478, 14409, 14404, 14476, 14402, 14271, 14269, 14270, 14256, 14257,
                14266, 14268, 14210, 14263, 14267]
    batchDownloadTTGPNG(picLists)
