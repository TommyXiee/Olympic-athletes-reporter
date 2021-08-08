# -*- coding: utf-8 -*-
import re
import requests
import time
import random
import csv
import json
import urllib
from lxml import etree, html

# 构造headers信息
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

def get_text(name):
    """
    从运动员界面爬取bio，返回列表
    """

    url = 'https://olympics.com/en/athletes/'+name  # 目标网址
    try:
        response = requests.get(url, headers=headers)  # 请求获取网页信息
    except:
        print('服务器请求失败')
    
    response.encoding = response.apparent_encoding  # 编码格式

    # fh = open('test.txt', 'w')
    # fh.write(response.content.decode("utf-8",'ignore'))
    # fh.close()

    Text = etree.HTML(response.text)

    a_x = Text.xpath('//*[@id="content"]/section[2]/div/div[2]/div/div/p/text()')

    return a_x


def get_name():    
    """
    通过API爬取运动员姓名,返回姓名
    """
    try:
        # time.sleep(random.randint(0, 3))  # 控制访问速度
        resp = urllib.request.urlopen('https://olympics.com/en/api/v1/search/default/athletes')
    except:
        print('API请求失败')
    
    ele_json = json.loads(resp.read())
    names = []
    for i in range(len(ele_json["modules"][0]["content"])):
        names.append(ele_json["modules"][0]["content"][i]["slug"])
    return names


def main():
    count = 0
    f = open('athletes.txt', 'w', encoding='utf-8')  # 创建一个csv文件
    print("ok")
    writer = csv.writer(f, delimiter='\t', quoting=csv.QUOTE_ALL)

    row = ['name','bios']  # 设置行首
    writer.writerow(row)
    names = get_name()
    print( len(names), "rows of data to be written",)
    for i in range(len(names)):
        bios = get_text(names[i])
        # print(type(bios))
        row = [names[i],bios]
        writer.writerow(row)
        print("a row written", i)

    f.close()


if __name__ == '__main__':
    main()

