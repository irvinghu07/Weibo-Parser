#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Sep 19 14:18:47 2020

@author: hanruo
"""

import copy
import aiohttp
import requests
import re
import asyncio
import json
import xlwt
import nest_asyncio
import os
import traceback


def get_page():
    """
    先用requests构造请求，解析出关键词搜索出来的微博总页数
    :return: 返回每次请求需要的data参数
    """
    data_list = []
    data = {
        'containerid': '100103type=1&q={}'.format(kw),
        'page_type': 'searchall'}
    try:
        resp = requests.get(url=url, headers=headers, params=data)
        total_page = resp.json()['data']['cardlistInfo']['total']  # 微博总数
        print(f"找到{total_page}条相关微博")
        # 一页有10条微博，用总数对10整除，余数为0则页码为总数/10，余数不为0则页码为（总数/10）+1
        if total_page % 10 == 0:
            page_num = int(total_page / 10)
        else:
            page_num = int(total_page / 10) + 1
        # 页码为1，data为当前data，页码不为1，通过for循环构建每一页的data参数
        if page_num == 1:
            data_list.append(data)
            return data_list
        else:
            for i in range(1, page_num + 1):
                data['page'] = i
                data_list.append(copy.deepcopy(data))
            return data_list
    except requests.RequestException:
        traceback.print_exc(file=open("error.txt", "a+"))


# async定义函数，返回一个协程对象
async def crawl(data):
    """
    多任务异步解析页面，存储数据
    :param data: 请求所需的data参数
    :return: None
    """
    async with aiohttp.ClientSession() as f:  # 实例化一个ClientSession
        async with await f.get(url=url, headers=headers, params=data) as resp:  # 携带参数发送请求
            text = await resp.text()  # await 等待知道获取完整数据
            text_dict = json.loads(text)['data']['cards']
            parse_dict = {}
            for card in text_dict:
                if card['card_type'] == 9:
                    scheme = card['scheme']
                    if card['mblog']['isLongText'] is False:
                        text = card['mblog']['text']
                        text = re.sub(r'<.*?>|\n+', '', text)
                    else:
                        text = card['mblog']['longText']['longTextContent']
                    user = card['mblog']['user']['profile_url']
                    created_at = card['mblog']['created_at']
                    comments_count = card['mblog']['comments_count']
                    attitudes_count = card['mblog']['attitudes_count']
                    parse_dict['url'] = scheme
                    parse_dict['created_at'] = created_at
                    parse_dict['text'] = text
                    parse_dict['author'] = user
                    parse_dict['comments_count'] = comments_count
                    parse_dict['attitudes_count'] = attitudes_count
                    parse_dict_list.append(copy.deepcopy(parse_dict))


def insert_data(file_name):
    """
    将数据导出到excle中
    :param file_name: 文件名
    :return:
    """
    wr = xlwt.Workbook(encoding='utf8')
    table = wr.add_sheet(file_name)
    table.write(0, 0, '原链接')
    table.write(0, 1, '创建时间')
    table.write(0, 2, '正文')
    table.write(0, 3, '作者首页')
    table.write(0, 4, '评论数')
    table.write(0, 5, '点赞数')
    for index, data in enumerate(parse_dict_list):
        table.write(index + 1, 0, data['url'])
        table.write(index + 1, 1, data['created_at'])
        table.write(index + 1, 2, data['text'])
        table.write(index + 1, 3, data['author'])
        table.write(index + 1, 4, data['comments_count'])
        table.write(index + 1, 5, data['attitudes_count'])
    file_path = file_name + '.xls'
    print(f"saving output file at path: {os.path.join(os.path.abspath('.'), file_path)}")
    wr.save(file_path)


def main(file_name):
    """
    开启多任务循环
    :return: None
    """
    data_list = get_page()  # 接收data参数列表
    task_list = []  # 定义一个任务列表
    for data in data_list:
        c = crawl(data)  # 调用协程，传参
        task = asyncio.ensure_future(c)  # 创建任务对象
        task_list.append(task)  # 将任务添加到列表中
    nest_asyncio.apply()
    loop = asyncio.get_event_loop()  # 创建事件循环
    loop.run_until_complete(asyncio.wait(task_list))  # 开启循环，并将阻塞的任务挂起
    insert_data(file_name)


if __name__ == '__main__':
    kw = input('关键词:')
    headers = {
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.1.2 Safari/605.1.15'}
    url = 'https://m.weibo.cn/api/container/getIndex'
    parse_dict_list = []  # 临时存放爬取的数据
    main(kw)
