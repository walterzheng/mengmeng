# coding: utf-8

import urllib2
import urllib
from bs4 import BeautifulSoup

import xlsxwriter

base_url = 'https://www.x-rates.com/historical/'

money_name_convert = {"Australian Dollar": "AUD",
                        "Bahraini Dinar": "BHD",
                        "Botswana Pula": "BWP",
                        "Brazilian Real": "BRL",
                        "Bruneian Dollar": "BND",
                        "Bulgarian Lev": "BGN",
                        "Canadian Dollar": "CAD",
                        "Chilean Peso": "CLP",
                        "Chinese Yuan Renminbi": "CNY",
                        "Colombian Peso": "COP",
                        "Croatian Kuna": "HRK",
                        "Czech Koruna": "CZK",
                        "Danish Krone": "DKK",
                        "Euro": "EUR",
                        "Hong Kong Dollar": "HKD",
                        "Hungarian Forint": "HUF",
                        "Icelandic Krona": "ISK",
                        "Indian Rupee": "INR",
                        "Indonesian Rupiah": "IDR",
                        "Iranian Rial": "IRR",
                        "Israeli Shekel": "ILS",
                        "Japanese Yen": "JPY",
                        "Kazakhstani Tenge": "KZT",
                        "South Korean Won": "KRW",
                        "Kuwaiti Dinar": "KWD",
                        "Libyan Dinar": "LYD",
                        "Malaysian Ringgit": "MYR",
                        "Mauritian Rupee": "MUR",
                        "Mexican Peso": "MXN",
                        "Nepalese Rupee": "NPR",
                        "New Zealand Dollar": "NZD",
                        "Norwegian Krone": "NOK",
                        "Omani Rial": "OMR",
                        "Pakistani Rupee": "PKR",
                        "Philippine Peso": "PHP",
                        "Polish Zloty": "PLN",
                        "Qatari Riyal": "QAR",
                        "Romanian New Leu": "RON",
                        "Russian Ruble": "RUB",
                        "Saudi Arabian Riyal": "SAR",
                        "Singapore Dollar": "SGD",
                        "South African Rand": "ZAR",
                        "Sri Lankan Rupee": "LKR",
                        "Swedish Krona": "SEK",
                        "Swiss Franc": "CHF",
                        "Taiwan New Dollar": "TWD",
                        "Thai Baht": "THB",
                        "Trinidadian Dollar": "TTD",
                        "Turkish Lira": "TRY",
                        "Emirati Dirham": "AED",
                        "British Pound": "GBP",
                        "US Dollar": "USD",
                        "Venezuelan Bolivar": "VEF"}

def getRate(cur_date = '2018-08-04', money = 'USD'):
    rtn = []
    
    paras = {'from': money, 'amount': '1', 'date': cur_date}

    url = "%s?%s"%(base_url, urllib.urlencode(paras))
    print url
    page = urllib2.urlopen(url)
    soup = BeautifulSoup(page, "lxml")

    table = soup.find_all('table', {"class":"tablesorter ratesTable"})[0]
    tr_iter = table.find_all('tr')
    for tr in tr_iter:
        td_iter = tr.find_all('td')
        rate_data = []
        for td in td_iter:
            rate_data.append(td.string)
        if len(rate_data) != 3:
            continue
        if rate_data[0] in money_name_convert:
            rate_data.append(money_name_convert[rate_data[0]])
            rtn.append(rate_data)
            
    return rtn

if __name__ == '__main__':
    # 希望获得日期，最早2008-01-01，注意日期的格式YYYY-MM-DD
    date_list = ["2008-01-01", 
                "2008-12-31", 
                '2009-12-31',
                '2010-12-31',
                '2011-12-30',
                '2012-12-31',
                '2013-12-31',
                '2014-12-31',
                '2015-12-31',
                '2016-12-30',
                '2017-12-29']  
    money_name = ["USD"]        # 希望转换的货币名字，money_name_convert中有所有合法的缩写
    out_filename = "exchange_rate.xlsx"   # 输出的文件

    workbook = xlsxwriter.Workbook(out_filename)
    worksheet = workbook.add_worksheet('Sheet1')
    worksheet.write(0, 0, "from")
    worksheet.write(0, 1, "to")
    worksheet.write(0, 2, "full name")
    worksheet.write(0, 3, "rate")
    worksheet.write(0, 4, "inverse rate")
    worksheet.write(0, 5, "date")


    counter = 0
    for date in date_list:
        for money in money_name:
            rtn = getRate(date, money)
            for rtn_data in rtn:
                counter += 1
                worksheet.write(counter, 0, money)
                worksheet.write(counter, 1, rtn_data[3])
                worksheet.write(counter, 2, rtn_data[0])
                worksheet.write(counter, 3, rtn_data[1])
                worksheet.write(counter, 4, rtn_data[2])
                worksheet.write(counter, 5, date)
    workbook.close()

