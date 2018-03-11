# -*- coding: utf-8 -*-

import xlrd
import math

def process(filename, sheetname, equity_type = 1):
    try:
        data = xlrd.open_workbook(filename)
        sheet = data.sheet_by_name(sheetname)
    except BaseException, e:
        print e.message

    rows = sheet.nrows
    cols = sheet.ncols

    rtn_dict = {}

    for j in xrange(0, cols, 6):
        equity_idx = j
        index_idx = j + 3

        equity_name = sheet.cell(0, j).value

        equity_history = {}
        index_history = {}
        for i in xrange(3, rows):
            if sheet.cell(i, equity_idx).value:
                if equity_type == 0:
                    equity_history[xlrd.xldate.xldate_as_datetime(sheet.cell(i, equity_idx).value, data.datemode)] \
                        = sheet.cell(i, equity_idx+1).value
                elif equity_type == 1 and sheet.cell(i-1, equity_idx).ctype == 3:
                    equity_history[xlrd.xldate.xldate_as_datetime(sheet.cell(i, equity_idx).value, data.datemode)] \
                        = (sheet.cell(i, equity_idx+1).value - sheet.cell(i-1, equity_idx+1).value) \
                            / sheet.cell(i-1, equity_idx+1).value * 100
            if sheet.cell(i, index_idx).value:        
                index_history[xlrd.xldate.xldate_as_datetime(sheet.cell(i, index_idx).value, data.datemode)] \
                    = sheet.cell(i, index_idx+1).value

        sum = 0.0
        counter = 0
        for key, value in equity_history.iteritems():
            if value != 0.0 and key in index_history:
                sum += (value - index_history[key])**2
                counter += 1
                # print key, value, index_history[key], sum, counter

        rtn_dict[equity_name] = math.sqrt(sum / counter) * 15.811388300841896
        print counter

    return rtn_dict

if __name__ == '__main__':
    filename = u'etf NEW.xlsx'
    sheet_name = [u'其他',u'恒生指数',u'沪深300',u'富时A50',u'恒生国企']
    # filename = u'恒生H股.xlsx'
    # sheet_name = [u'Sheet1']

    out_file_name = "result.txt"
    wfid = open(out_file_name, "w")
    for name in sheet_name:
        rtn = process(filename, name)
        for key in rtn:
            wfid.write("%s:%.2f%%\n"%(key.encode('utf8'), rtn[key]))
    wfid.close()
