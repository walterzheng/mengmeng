# -*- coding: utf-8 -*-

import json
import random
import datetime

import requests
import urllib

post_url = 'http://gs.amac.org.cn/amac-infodisc/api/fund/account'
post_header = {'content-type': 'application/json'}

def query_zhuanhu_data(start_date, end_date):
    query_paras = {'rand': random.random(), 'page': '0', 'size': '100'}
    payload_paras={"registerDateFrom": start_date, "registerDateTo": end_date}

    response = requests.post('%s?%s'%(post_url, urllib.urlencode(query_paras)), 
        json.dumps(payload_paras), headers=post_header)

    data = json.loads(response.text)
    total_elements = data['totalElements']
    total_pages = data['totalPages']

    elements_list = data['content'][:]

    for i in xrange(1, total_pages):
        query_paras['page'] = str(i)
        response = requests.post('%s?%s'%(post_url, urllib.urlencode(query_paras)), 
            json.dumps(payload_paras), headers=post_header)

        data = json.loads(response.text)
        elements_list.extend(data['content'])

    return total_elements, elements_list

def dump_zhuanhu_date(total_elements, elements_list):
    if total_elements != len(elements_list):
        print 'error in query'
        return

    ofid = open('zhuanhu.txt', "w")
    for ele in elements_list:
        name = ele['name'].encode('utf8')
        manager = ele['manager'].encode('utf8')
        zhuanhu_type = ele['type'].encode('utf8')
        register_code = ele['registerCode'].encode('utf8')
        register_date = datetime.datetime.fromtimestamp(ele['registerDate']/1000).strftime("%Y-%m-%d %H:%M:%S")

        ofid.write("%s\t%s\t%s\t%s\t%s\n"%(name, manager, zhuanhu_type, register_code, register_date))
    ofid.close()


if __name__ == '__main__':
    start_date = "2018-03-05"
    end_date = "2018-03-10"

    total_elements, elements_list = query_zhuanhu_data(start_date, end_date)
    dump_zhuanhu_date(total_elements, elements_list)