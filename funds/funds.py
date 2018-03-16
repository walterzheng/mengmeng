# coding: utf-8

# In[1]:

import os
import datetime
import xlrd
import collections
import math


# In[22]:

def add_years(s, p):
    ''' 在日期s上加上p年，如果日期不存在，顺延到下个月1号 '''
    if s.month != 2 or s.day != 29:
        return datetime.datetime(year=s.year+p, month=s.month, day=s.day)
    else:
        try:
            rs = datetime.datetime(year=s.year+p, month=2, day=29)
        except ValueError, e:
            rs = datetime.datetime(year=s.year+p, month=3, day=1)
        return rs

def add_months(s, p):
    ''' 在日期s上加上p个月，如果日期不存在，顺延到下个月1号 '''
    p_years = s.year + p / 12
    p_months = s.month + p % 12 
    if p_months > 12:
        p_years += 1
        p_months = p_months % 12
    
    try:
        rs = datetime.datetime(p_years, p_months, s.day)
    except ValueError, e:
        if p_months == 12:
            rs = datetime.datetime(p_years + 1, 1, 1)
        else:
            rs = datetime.datetime(p_years, p_months+1, 1)
    return rs

def avg(list1):
    if len(list1) == 0:
        return float('nan')
    return sum(list1)/float(len(list1))

def var(list1):
    length = len(list1)
    if length <= 1:
        return float('nan')
    
    avg1 = avg(list1)
    return sum((foobar - avg1)**2 for foobar in list1) / float(length -1)

def cov(list1, list2):
    length = len(list1)
    if length <= 1 or length != len(list2):
        return float('nan')
    
    avg1 = avg(list1)
    avg2 = avg(list2)
    
    return sum((list1[i] - avg1)*(list2[i] - avg2) for i in xrange(length)) / float(length-1)


# In[3]:

Rp_data=collections.defaultdict(dict) #基金复权单位净值
Rm_data=collections.defaultdict(dict) #指数价格

RFKEY = "RFKEY" #为了格式保持一致，增加字典中不必要的一层
RMKEY = "RMKEY" #为了格式保持一致，增加字典中不必要的一层

Rp=None
Rf={RFKEY:{}} #为了格式保持一致，增加字典中不必要的一层
Rm=None

BEGIN_YEAR = 2004
END_YEAR = 2018

# 季度的分割日期，[datetime.datetime(2004, 1, 1, 0, 0), datetime.datetime(2004, 4, 1, 0, 0), datetime.datetime(2004, 7, 1, 0, 0) 。。。
quater_date = [datetime.datetime(v1,v2,1) for v1 in xrange(BEGIN_YEAR, END_YEAR) for v2 in xrange(1,12,3)]
quater_date.append(add_months(quater_date[-1], 3))


# In[4]:

def Rp_load_data(filename = 'Fund_NAV1.txt'):
    '''加载基金复权单位净值文件'''
    with open(filename, "r") as ifid:
        line_num = 0
        for line in ifid:
            if line_num > 0 and len(line) > 5:
                try:
                    FundClassID, TradingDate, Symbol, AccumulativeNAV = line.replace('\x00','').strip().split('\t')
                    TradingDate = datetime.datetime.strptime(TradingDate, "%Y-%m-%d")
                    if TradingDate >= datetime.datetime(BEGIN_YEAR, 1, 1) and TradingDate < datetime.datetime(END_YEAR, 1, 1):
                        Rp_data[Symbol][TradingDate] = float(AccumulativeNAV)
                except ValueError, e:
                    print filename, line_num, line.replace('\x00','').strip()
                                  
            line_num += 1
            
def Rm_load_data(filename="IDX_Idxtrd1.txt"):
    '''加载指数价格文件'''
    with open(filename, "r") as ifid:
        line_num = 0
        for line in ifid:
            if line_num > 0 and len(line) > 5:
                try:
                    Indexcd, Idxtrd01, Idxtrd05 = line.replace('\x00','').strip().split('\t')
                    Idxtrd01 = datetime.datetime.strptime(Idxtrd01, "%Y-%m-%d")
                    if Idxtrd01 >= datetime.datetime(BEGIN_YEAR, 1, 1) and Idxtrd01 < datetime.datetime(END_YEAR, 1, 1):
                        Rm_data[Indexcd][Idxtrd01] = float(Idxtrd05)
                except ValueError, e:
                    print filename, line_num, line.replace('\x00','').strip()
                                  
            line_num += 1
            
def Rf_load_data(filename="Fund_RiskFree.txt"):
    '''加载无风险利率文件，注意以上两个不同，这个直接得到的是收益率'''
    with open(filename, "r") as ifid:
        line_num = 0
        for line in ifid:
            if line_num > 0 and len(line) > 5:
                try:
                    BenchMarkID, BenchMark, TradingDate, InterestRate = line.replace('\x00','').strip().split('\t')
                    TradingDate = datetime.datetime.strptime(TradingDate, "%Y-%m-%d")
                    if TradingDate >= datetime.datetime(BEGIN_YEAR, 1, 1) and TradingDate < datetime.datetime(END_YEAR, 1, 1):
                        Rf[RFKEY][TradingDate] = float(InterestRate) / 100 / 250 # 无风险收益
                except ValueError, e:
                    print filename, line_num, line.replace('\x00','').strip()
                                  
            line_num += 1


# In[29]:

def calc_earning_ratio(data):
    '''计算基金或者市场基准收益率'''
    rtn = collections.defaultdict(dict)
    log = open("log_calc_earning_ratio.txt", "a")
    for identifier in data:
        id_data_num = 0
        pre_date = None
        pre_value = 0.0
        for cur_date in sorted(data[identifier].keys()): #按时间顺序遍历id下的所有数据
            value = data[identifier][cur_date]
            if id_data_num > 0:
                rtn[identifier][cur_date] = math.log(value) - math.log(pre_value)
                if cur_date - pre_date > datetime.timedelta(days=10):
                    log.write("%s %d %s %f 的前一条数据是 %s %f，疑似有数据丢失，请注意\n" % (identifier, (cur_date - pre_date).days, 
                          cur_date.strftime("%Y-%m-%d"), value, pre_date.strftime("%Y-%m-%d"), pre_value))
            id_data_num += 1
            pre_date = cur_date
            pre_value = value
    log.close()
    return rtn

def calc_Rm(data, log_file="calc_Rm.txt"):
    log = open(log_file, "w")
    
    rtn = {RMKEY: {}}
    
    set_000001 = set(data['000001'].iterkeys())
    set_399001 = set(data['399001'].iterkeys())
    set_000012 = set(data['000012'].iterkeys())
    
    union_set = (set_000001|set_399001|set_000012)
    inter_set = (set_000001&set_399001&set_000012)
    
    only_000001 = union_set - set_000001
    only_399001 = union_set - set_399001
    only_000012 = union_set - set_000012
    
    log.write("000001 lost %s\n"%'|'.join(foobar.strftime("%Y-%m-%d") for foobar in only_000001))
    log.write("399001 lost %s\n"%'|'.join(foobar.strftime("%Y-%m-%d") for foobar in only_399001))
    log.write("000012 lost %s\n"%'|'.join(foobar.strftime("%Y-%m-%d") for foobar in only_000012))
    
    for trade_date in inter_set:
        rtn[RMKEY][trade_date] = 0.4 * data['000001'][trade_date] + 0.4 * data['399001'][trade_date] + 0.2 * data['000012'][trade_date]
    
    log.close()
    return rtn

def get_quater_data(data):
    '''把数据按照季度重新整理'''
    quater_data = collections.defaultdict(lambda : collections.defaultdict(list))  # 保存季度收益率
    quater_stamp = collections.defaultdict(lambda : collections.defaultdict(list)) # 保存季度收益率对应的时间戳
    for identifier in data:
        cur_quater_idx = 0
        for cur_date in sorted(data[identifier].keys()): #按时间顺序遍历id下的所有数据
            if cur_date >= quater_date[cur_quater_idx]:
                while (cur_date >= quater_date[cur_quater_idx+1]):
                    cur_quater_idx += 1
                quater_data[identifier][quater_date[cur_quater_idx]].append(data[identifier][cur_date])
                quater_stamp[identifier][quater_date[cur_quater_idx]].append(cur_date)
    return quater_data, quater_stamp
    
def calc_quater_sigma(quater_data, dump_filename="quater_sigma.txt"):
    '''计算季度波动率'''
    sigma_data = collections.defaultdict(dict)
    n_data = collections.defaultdict(dict)
    mu_data = collections.defaultdict(dict)
    for identifier, values in quater_data.iteritems():
        for quater, p_list in values.iteritems():
            mu = avg(p_list)
            n = len(p_list)
            if n <= 1:
                sigma = float('nan')
            else:
                sigma = math.sqrt( sum((foobar - mu)**2 for foobar in p_list) / float(n-1) )
            sigma_data[identifier][quater] = sigma
            n_data[identifier][quater] = n
            mu_data[identifier][quater] = mu

    return sigma_data

def calc_sharpe(Rp_quater, Rf_quater, Rp_quater_sigma):
    sharpe_data = collections.defaultdict(dict)
    for identifier, values in Rp_quater.iteritems():
        for quater, p_list in values.iteritems():
            sharpe = float('nan')
            if Rp_quater_sigma[identifier][quater] != 0:
                sharpe = (avg(p_list) - avg(Rf_quater[RFKEY][quater])) / Rp_quater_sigma[identifier][quater]            
            sharpe_data[identifier][quater] = sharpe
            
    return sharpe_data

def join_quater(quater1, quater2, stamp1, stamp2):
    '''根据stamp1和stamp2中共同的日期，计算共有的序列，以便计算协方差'''
    join_set = (set(stamp1) & set(stamp2))
    
    list1 = [quater1[i] for i in xrange(len(quater1)) if stamp1[i] in join_set]
    list2 = [quater2[i] for i in xrange(len(quater2)) if stamp2[i] in join_set]
    
    return len(join_set), list1, list2
            

def calc_trevnor_jensen(Rp_quater, Rm_quater, Rf_quater, Rp_quater_stamp, Rm_quater_stamp, Rf_quater_stamp):
    trevnor_data = collections.defaultdict(dict)
    jensen_data = collections.defaultdict(dict)
    for identifier, values in Rp_quater.iteritems():
        for quater, rp_list in values.iteritems():
            rm_list = Rm_quater[RMKEY][quater]
            rf_list = Rf_quater[RFKEY][quater]
            
            rp_stamp_list = Rp_quater_stamp[identifier][quater]
            rm_stamp_list = Rm_quater_stamp[RMKEY][quater]
            rf_stamp_list = Rf_quater_stamp[RFKEY][quater]
            
            rm_var = var(rm_list)
            rp_rm_join_num, rp_join_list, rm_join_list = join_quater(rp_list, rm_list, rp_stamp_list, rm_stamp_list)
            rp_rm_cov = cov(rp_join_list, rm_join_list)
            
            rp_avg = avg(rp_list)
            rm_avg = avg(rm_list)
            rf_avg = avg(rf_list)
            
            trevnor = float('nan')
            if rp_rm_cov and rm_var and rp_rm_cov != 0 :
                trevnor = (rp_avg - rf_avg) * rm_var / rp_rm_cov
                
            jensen = float('nan')
            if rp_rm_cov and rm_var and rm_var != 0:
                jensen = (rp_avg - rf_avg) - (rm_avg - rf_avg) * rp_rm_cov / rm_var
                
            trevnor_data[identifier][quater] = trevnor
            jensen_data[identifier][quater] = jensen
            
    return trevnor_data, jensen_data
    
def calc_max_rollback(Rp_data_quater):
    max_rollback_data = collections.defaultdict(dict)
    for identifier, values in Rp_data_quater.iteritems():
        for quater, data_list in values.iteritems():
            max_rollback = 0
            if len(data_list) > 0:
                for i in xrange(len(data_list)-1):
                    for j in xrange(i+1, len(data_list)):
                        roll_back = (data_list[j] - data_list[i]) / data_list[i]
                        max_rollback = max((max_rollback, roll_back))
                        
            max_rollback_data[identifier][quater] = max_rollback
    return max_rollback_data
                


# In[34]:

def dump_result(data, filename):
    with open(filename, "w") as ofid:
        ofid.write('quater\t')
        for quater in  quater_date[:-1]:
            ofid.write("%dQ%d\t"% (quater.year, quater.month/3+1))
            ofid.write("%dQ%d\t"% (quater.year, quater.month/3+1))
            ofid.write("%dQ%d\t"% (quater.year, quater.month/3+1))
        ofid.write('\n')

        for identifier in data:
            ofid.write("%s\t"%identifier)
            for quater in  quater_date[:-1]:
                if quater in data[identifier]:
                    ofid.write("%s\t"% str(data[identifier][quater]))
                else:
                    ofid.write("-\t")
            ofid.write('\n')


# In[37]:

def load_data():
    Rm_load_data("IDX_Idxtrd1.txt")
    Rm_load_data("IDX_Idxtrd2.txt")
    Rm_load_data("IDX_Idxtrd3.txt")
    Rm_load_data("IDX_Idxtrd4.txt")

    Rp_load_data('Fund_NAV1.txt')
    Rp_load_data('Fund_NAV2.txt')
    Rp_load_data('Fund_NAV3.txt')
    Rp_load_data('Fund_NAV4.txt')
    Rp_load_data('Fund_NAV5.txt')
    Rp_load_data('Fund_NAV6.txt')
    
    Rf_load_data("Fund_RiskFree.txt")

    print "共加载%d只基金的%d条累计收益记录"%(len(Rp_data), sum(len(foobar) for foobar in Rp_data.itervalues()))
    print "共加载%d只指数的%d条价格记录"%(len(Rm_data), sum(len(foobar) for foobar in Rm_data.itervalues()))

def calc():
    if os.path.exists("log_calc_earning_ratio.txt"):
        os.remove("log_calc_earning_ratio.txt")
    Rp = calc_earning_ratio(Rp_data)
    Rm_tmp = calc_earning_ratio(Rm_data)
    Rm = calc_Rm(Rm_tmp)

    print "共加载%d只基金的%d天收益率"%(len(Rp), sum(len(foobar) for foobar in Rp.itervalues()))
    print "共加载%d只指数的%d天收益率"%(len(Rm), sum(len(foobar) for foobar in Rm.itervalues()))
    
    Rp_quater, Rp_quater_stamp = get_quater_data(Rp)
    Rm_quater, Rm_quater_stamp = get_quater_data(Rm)
    Rf_quater, Rf_quater_stamp = get_quater_data(Rf)
    
    Rp_quater_sigma = calc_quater_sigma(Rp_quater, "quater_sigma.txt")
    Rp_quater_sharpe = calc_sharpe(Rp_quater, Rf_quater, Rp_quater_sigma)
    Rp_trevnor_data, Rp_jensen_data = calc_trevnor_jensen(Rp_quater, Rm_quater, Rf_quater, Rp_quater_stamp, Rm_quater_stamp, Rf_quater_stamp)
    
    Rp_data_quater, Rp_data_quater_stamp = get_quater_data(Rp_data)
    max_roll_back_data = calc_max_rollback(Rp_data_quater)
    
    dump_result(Rp_quater_sigma, "sigma.txt")
    dump_result(Rp_quater_sharpe, "sharpe.txt")
    dump_result(Rp_trevnor_data, "trevnor.txt")
    
    dump_result(max_roll_back_data, "rollback.txt")
    dump_result(Rp_jensen_data, "jensen.txt")
    
    print "calc done"
    
def main():
    load_data()
    calc()

if __name__ == "__main__":
    main()



