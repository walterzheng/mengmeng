import pandas as pd

data=pd.read_excel("a.xlsx", sheetname="sheet1")

def print_company_info(com_name):
    d1 = data[(data[u'管理人'] == com_name) \
        & (data[u'类型'] == u'宽基') \
        & (data[u'投资地域'] == u'内地')]\
        [[u'证券简称', u'剔除联接规模', u'指数类型', u'管理人']]

    d2 = data[(data[u'管理人'] == com_name) \
        & (data[u'类型'] == u'行业') \
        & (data[u'投资地域'] == u'内地')]\
        [[u'证券简称', u'剔除联接规模', u'指数类型', u'管理人']]

    d3 = data[(data[u'管理人'] == com_name) \
        & (data[u'类型'] == u'宽基') \
        & (data[u'投资地域'] == u'海外')]\
        [[u'证券简称', u'剔除联接规模', u'指数类型', u'管理人']]

    d4 = data[(data[u'管理人'] == com_name) \
        & (data[u'类型'] == u'行业') \
        & (data[u'投资地域'] == u'海外')]\
        [[u'证券简称', u'剔除联接规模', u'指数类型', u'管理人']]

    for _, d in d1.iterrows():
        fenjistr = u''
        start_idx = len(d[u'管理人']) - 2
        if d[u'指数类型'] == u'股票分级':
            fenjistr = u'分级'
        if d[u'证券简称'][-1] == u'A':
            jiancheng = d[u'证券简称'][start_idx:-1]
        else:
            jiancheng = d[u'证券简称'][start_idx:]
        guimo = u'（' + unicode(int(round(d[u'剔除联接规模']))) + u'）'
        print (jiancheng + fenjistr + guimo).encode('utf8')

    print '------------------'

    for _, d in d2.iterrows():
        fenjistr = u''
        start_idx = len(d[u'管理人']) - 2
        if d[u'指数类型'] == u'股票分级':
            fenjistr = u'分级'
        if d[u'证券简称'][-1] == u'A':
            jiancheng = d[u'证券简称'][start_idx:-1]
        else:
            jiancheng = d[u'证券简称'][start_idx:]
        guimo = u'（' + unicode(int(round(d[u'剔除联接规模']))) + u'）'
        print (jiancheng + fenjistr + guimo).encode('utf8')

    print '------------------'

    for _, d in d3.iterrows():
        fenjistr = u''
        start_idx = len(d[u'管理人']) - 2
        if d[u'指数类型'] == u'股票分级':
            fenjistr = u'分级'
        if d[u'证券简称'][-1] == u'A':
            jiancheng = d[u'证券简称'][start_idx:-1]
        else:
            jiancheng = d[u'证券简称'][start_idx:]
        guimo = u'（' + unicode(int(round(d[u'剔除联接规模']))) + u'）'
        print (jiancheng + fenjistr + guimo).encode('utf8')

    print '------------------'

    for _, d in d4.iterrows():
        fenjistr = u''
        start_idx = len(d[u'管理人']) - 2
        if d[u'指数类型'] == u'股票分级':
            fenjistr = u'分级'
        if d[u'证券简称'][-1] == u'A':
            jiancheng = d[u'证券简称'][start_idx:-1]
        else:
            jiancheng = d[u'证券简称'][start_idx:]
        guimo = u'（' + unicode(int(round(d[u'剔除联接规模']))) + u'）'
        print (jiancheng + fenjistr + guimo).encode('utf8')