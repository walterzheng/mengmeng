# -*- coding: utf-8 -*-
# 使用前先执行 pip install aiohttp
import collections
import aiohttp
import asyncio
import logging
import time
import re
import json

from datetime import datetime

# 爬取前MAX_PAGE页
MAX_PAGE = 20

NUM_PRODUCERS = 3
NUM_HTML_CONSUMERS = NUM_PRODUCERS

# 证监会公示网页
BASE_URL = 'https://neris.csrc.gov.cn/alappl/home1/newOnlinealog'

HEADERS_STR = '''
Accept-Encoding:gzip, deflate, br
Accept-Language:zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7
Cookie:JSESSIONID=73509259A4CB5A152E4779E1803B4792
Referer:https://neris.csrc.gov.cn/alappl/home/gongshi
Sec-Ch-Ua:"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"
Sec-Ch-Ua-Mobile:?0
Sec-Ch-Ua-Platform:"macOS"
Sec-Fetch-Dest:empty
Sec-Fetch-Mode:cors
Sec-Fetch-Site:same-origin
User-Agent:Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36
X-Requested-With:XMLHttpRequest'''


DATE_FORMAT_RO = re.compile(r'\d\d\d\d-\d\d-\d\d')

async def featch_data(pn: int, session: aiohttp.ClientSession) -> str:
    paras_dict = {'pageNo':str(pn), 'pageSize':'10', 'appMatrCde': '', 'appMatrName': '', 'apptName': ''}

    HEADERS = dict([[h.partition(':')[0].strip(), h.partition(':')[2].strip()]
                for h in HEADERS_STR.format(pn).split('\n')])
    async with session.get(BASE_URL, params = paras_dict, headers = HEADERS) as resp:
        return await resp.text()

async def producer(name: str,
                   page_range: range,
                   html_q: asyncio.Queue,
                   session: aiohttp.ClientSession) -> None:
    logging.info(f"Producer {name} inited," +
                 f" fetch page {page_range.start} -> {page_range.stop-1}.")
    for pn in page_range:
        html = await featch_data(pn, session)
        logging.info(f"Producer {name} fetch page {pn}, with length {len(html)}")
        await html_q.put((pn, html))
    logging.info(f"Producer {name} finished.")

async def html_consumer(name: str,
                        html_q: asyncio.Queue,
                        data: dict) -> None:
    logging.info(f"Html consumer {name} inited.")
    while True:
        pn, html = await html_q.get()
        logging.info(f"Html consumer {name} get page {pn}, with length {len(html)}")
        contents = parse_data(html)
        logging.info(f"Html consumer {name} parse page {pn}, get {len(contents)} lines")
        data[pn] = contents
        html_q.task_done()

def write_out(data: dict) -> None:
    valid_action = collections.OrderedDict()
    valid_action["接收材料"] = ""
    valid_action["补正通知"] = ""
    valid_action["接收补正材料"] = ""
    valid_action["受理通知"] = ""
    valid_action["一次书面反馈"] = ""
    valid_action["接收书面回复"] = ""
    valid_action["行政许可决定书"] = ""
    valid_action["二次书面反馈"] = ""
    valid_action["一次中止审查通知"] = ""
    valid_action["申请人主动撤销"] = ""
    valid_action["终止审查通知"] = ""
    with open('机构公示.txt', "w", encoding='utf-8') as ofid:
        ofid.write('%s|%s|%s\n'%('页码', '标题', "|".join(valid_action.keys())))
        for pn in range(1, MAX_PAGE + 1):
            if pn not in data:
                logging.warning(f"Page {pn} not in write out data.")
                continue
            logging.info(f"Write out page {pn}.")
            for title, _, table in data[pn]:
                table_content_dict = valid_action.copy()
                for table_content in table:
                    table_content_dict[table_content[0]] = table_content[1]
                ofid.write('%d|%s|%s\n'%(pn, title, "|".join(table_content_dict.values())))

async def main(prod_num: int, html_con_num: int):
    data = dict()
    async with aiohttp.ClientSession(connector=aiohttp.TCPConnector(ssl=False)) as session:
        html_q = asyncio.Queue()
        prod_num = min(prod_num, MAX_PAGE)
        producer_list = [asyncio.create_task(producer(idx, page_range, html_q, session))
                         for idx, page_range in divide_page(prod_num)]
        html_consumer_list = [asyncio.create_task(html_consumer(idx, html_q, data))
                              for idx in range(html_con_num)]
        await asyncio.gather(*producer_list)
        await html_q.join()
        for con in html_consumer_list:
            con.cancel()
    write_out(data)

def divide_page(prod_num):
    size = MAX_PAGE // prod_num
    for i in range(prod_num):
        left = i * size + 1
        right = 1 + (MAX_PAGE if i + 1 == prod_num else (i + 1) * size)
        yield (i, range(left, right))

def parse_data(html: str):
    contents = []

    parsed_html = json.loads(html)
    for data in parsed_html['appltList']:
        title = data['appMatrName']
        table_content = []
        for table_line in data['aprvSchdPubFlowPOs']:
            dt_object = datetime.fromtimestamp(int(table_line['foundTime'])//1000)
            table_content.append((table_line['taskName'], dt_object.strftime('%Y-%m-%d')))

        contents.append((title, '', table_content))
    return contents

if __name__ == '__main__':
    logging.basicConfig(format='%(asctime)s - %(message)s', level=logging.DEBUG)
    
    logging.info(f"Program begin, crawl {MAX_PAGE} pages, " +
                 f"with {NUM_PRODUCERS} producers and {NUM_HTML_CONSUMERS} consumers")
    start = time.perf_counter()
    asyncio.run(main(NUM_PRODUCERS, NUM_HTML_CONSUMERS))
    elapsed = time.perf_counter() - start
    logging.info(f"Program completed in {elapsed:0.5f} seconds.")
