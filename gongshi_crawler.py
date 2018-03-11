# -*- coding: utf-8 -*-

import itertools
import socket
import re

import urllib2
import urllib
import requests
from bs4 import BeautifulSoup

date_format_ro = re.compile(r'\d\d\d\d-\d\d-\d\d')

# 证监会公示网页
post_url = 'http://ndes.csrc.gov.cn/alappl/home/volunteerLift.do'

def get_gongshi_page(pn):
    '''从每个page中爬title和date'''
    paras_dict = {'edCde':'300009', 'pageNo':str(pn), 'pageSize':'10'}
    doc = requests.post(post_url, paras_dict)

    soup = BeautifulSoup(doc.text, "lxml")
    titles = soup.find_all("div", {"class":"titleshow"})

    contents = []
    for title in titles:
        title_text = title.text.strip()

        cur_tag = title
        for i in xrange(10): # 最多向后检查10个sibling，检查是否有时间串
            cur_tag = cur_tag.next_sibling # 这里有坑，'\n'也算一个sibling
            if not hasattr(cur_tag, "text"):
                continue
            date_search = date_format_ro.search(cur_tag.text)
            if date_search:
                title_date = date_search.group()
                break
        else:
            title_date = u'fail to crawl date'
        
        contents.append((title_text, title_date))
    return contents

if __name__ == '__main__':
    max_page = 10

    all_contents = []
    for pn in xrange(1, max_page+1):
        all_contents.append(get_gongshi_page(pn))

    with open(u'机构公示.txt', "w") as ofid:
        for idx in xrange(len(all_contents)):
            for title, date in all_contents[idx]:
                ofid.write('%d\t%s\t%s\n'%(idx+1, title.encode('utf8'), date.encode('utf8')))