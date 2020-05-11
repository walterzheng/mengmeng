# -*- coding: utf-8 -*-
import collections

import itertools
import socket
import re

import requests
from bs4 import BeautifulSoup

# 禁用verify = False的告警
from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

date_format_ro = re.compile(r'\d\d\d\d-\d\d-\d\d')

# 证监会公示网页
post_url = 'https://neris.csrc.gov.cn/alappl/home/volunteerLift.do'

def get_children_number(tag):
    '''获取tag的子tag数目'''
    number = 0
    for child in tag.children:
        if not hasattr(child, "text"):
            continue
        number += 1
    return number

def get_deep_text(tag):
    '''获取tag最深一层的text'''
    cur_tag = tag
    cur_tag_number = get_children_number(cur_tag)
    while get_children_number(cur_tag) != 0:
        if cur_tag_number == 1:
            for child in cur_tag.children:
                if not hasattr(child, "text"):
                    continue
                else:
                    cur_tag = child
                    cur_tag_number = get_children_number(cur_tag)
        else:
            return "find_more_than_one_tag"
    return cur_tag.text.strip()

def get_gongshi_page(pn):
    '''从每个page中爬title和date'''
    paras_dict = {'edCde':'300009', 'pageNo':str(pn), 'pageSize':'10'}
    doc = requests.post(post_url, paras_dict, verify = False)

    soup = BeautifulSoup(doc.text, "lxml")
    titles = soup.find_all("div", {"class":"titleshow"})

    contents = []
    for title in titles:
        title_text = title.text.strip()
        title_date = ""
        table_content = []
        cur_tag = title
        for i in xrange(7): # 最多向后检查7个sibling，检查是否有时间串
            cur_tag = cur_tag.next_sibling # 这里有坑，'\n'也算一个sibling
            if not hasattr(cur_tag, "text"):
                continue
            date_search = date_format_ro.search(cur_tag.text)
            if date_search and title_date == "":
                title_date = date_search.group()
            
            if cur_tag.name == "table" and len(table_content) == 0:
                # table的第一行为“进度追踪”，第二行为标题行“任务名称 | 完成时间”，因此从第三行开始解析
                # table的每一行有两列
                table_cur_line = 0
                for table_tag in cur_tag.children:
                    if not hasattr(table_tag, "text"):
                        continue
                    table_cur_line += 1
                    if table_cur_line < 3:
                        continue
                    td_tags = table_tag.find_all("td")
                    if len(td_tags) != 2:
                        continue
                    table_content.append((get_deep_text(td_tags[0]), get_deep_text(td_tags[1])))
        
        contents.append((title_text, title_date, table_content))
    return contents

if __name__ == '__main__':
    start_page = 1 # 从第几页开始拉数据
    max_page = 50 # 共拉取多少页的数据 

    all_contents = []
    for pn in xrange(start_page, max_page+start_page):
        print "processing page %d, please wait~" % pn
        all_contents.append(get_gongshi_page(pn))

    valid_action = collections.OrderedDict()
    valid_action[u"接收材料"] = ""
    valid_action[u"补正通知"] = ""
    valid_action[u"接收补正材料"] = ""
    valid_action[u"受理通知"] = ""
    valid_action[u"一次书面反馈"] = ""
    valid_action[u"接收书面回复"] = ""
    valid_action[u"行政许可决定书"] = ""
    valid_action[u"二次书面反馈"] = ""
    valid_action[u"一次中止审查通知"] = ""
    valid_action[u"申请人主动撤销"] = ""
    valid_action[u"终止审查通知"] = ""

    with open(u'机构公示.txt', "w") as ofid:
        ofid.write('%s|%s|%s\n'%(u'页码'.encode('utf8'), u'标题'.encode('utf8'), u"|".join(valid_action.keys()).encode('utf8')))
        for idx in xrange(start_page, max_page+start_page):
            for title, date, table in all_contents[idx-start_page]:
                table_content_dict = valid_action.copy()
                for table_content in table:
                    table_content_dict[table_content[0]] = table_content[1]
                ofid.write('%d|%s|%s\n'%(idx+1, title.encode('utf8'), u"|".join(table_content_dict.values()).encode('utf8')))