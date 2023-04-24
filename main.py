# -*-coding:utf-8-*-
import pyodbc
import requests
import os
import sys
import json
import time
import pygame

# 连接数据库（不需要配置数据源）,connect()函数创建并返回一个 Connection 对象
cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\Program Files (x86)\DoorKq\data.mdb')
# cursor()使用该连接创建（并返回）一个游标或类游标的对象
crsr = cnxn.cursor()

# 打印数据库goods.mdb中的所有表的表名
print('`````````````` goods ``````````````')


# sys.exit()


## 存储用户签到信息
def save_sign_info(data):
    print('`````````````` 开始提交到OA服务端 ``````````````')
    url = 'http://oa.edushuyi.com/index/api/toutiaoapi_new'
    headers = {'Accept': '*/*',
               'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
               'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.87 Safari/537.36',
               'X-Requested-With': 'XMLHttpRequest'}
    r = requests.post(url, data=data, headers=headers)
    print(r.text)
    return r


def isdatasize():
    file_size = os.path.getsize('D:\Program Files (x86)\DoorKq\data.mdb')
    with open('D:/wwwroot/ace/size.txt', encoding='utf-8') as f2:
        contents = f2.read()
    if file_size != contents:
        print('`````````````` 考勤MDB文件发生变化重新读取 ``````````````')
        datarr = getData()
        if datarr:
            print(datarr)
            r = save_sign_info(str(datarr))
            data = r.json()
            print('`````````````` 开始标记ID写入情况 ``````````````')
            pygame.mixer.init()
            if data['error'] == 1:
                data['list'] = [1, 2, 3, 4, 5, 6]
                for index, element in enumerate(data['list']):
                    print(element)
                    open('D:/wwwroot/ace/' + str(element) + '.txt', 'w')
                    # pygame.mixer.music.load("D:/wwwroot/ace/audio/"+str(element)+".mp3")
                    pygame.mixer.music.load("http://qhqiqiu.shuchengrj.com/song/206bffe62bd149b82531d34775e2eeb3.mp3")
                    pygame.mixer.music.play()
                print('`````````````` 标记ID成功 ``````````````')
                f = open("D:/wwwroot/ace/size.txt", "w")
                f.write(str(file_size))
                f.close()
                print('---------------- 执行成功10S后重新查询 -------------')
            else:
                print('`````````````` error失败 ``````````````')
        else:
            print('`````````````` 数据没有值return 0 1 ``````````````')


def getData():
    print('`````````````` 读取考勤MDB文件 ``````````````')
    var_list = []
    # for table_info in crsr.tables(tableType='TABLE'):
    # print(table_info.table_name)
    SQL = 'SELECT * FROM mj_swipetime;'  # your query goes here
    rows = cnxn.execute(SQL).fetchall()
    # print(rows)
    for index, element in enumerate(rows):
        print('`````````````` 判断考勤是否重复提交过 ``````````````')
        if not os.path.exists('D:/wwwroot/ace/' + str(element[1]) + '.txt'):
            var_list.append(element)
    if var_list:
        print('`````````````` 数据有值整理POST ``````````````')
        return var_list
    else:
        print('`````````````` 数据没有值return 0 ``````````````')
        return 0


while True:
    isdatasize()
    time.sleep(10)



