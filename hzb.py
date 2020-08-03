import pandas as pd
import numpy as np
import xlsxwriter as writer
import re
import collections

def process():
    filename = '创业板压力测试模板-test-2-计算.xlsx'

    data = pd.read_excel(filename, "资产账户列表", engine="openpyxl")
    #zhengquan = pd.read_excel(filename, "创业板证券底表", engine="openpyxl")
    zhengquan_raw = pd.read_excel(filename, "创业板证券底表", engine="openpyxl")
    gongyunjiage = pd.read_excel(filename, "启用公允价格汇总表", engine="openpyxl")[['证券代码', '公允价格']]
    zhengquan_raw['证券代码'] = zhengquan_raw['证券代码'].apply(lambda s: "%06d"%s)
    gongyunjiage['证券代码'] = gongyunjiage['证券代码'].apply(lambda s: "%06d"%s)

    xinyongzhengquan = pd.read_excel(filename, "创业板信用证券底表", engine="openpyxl")
    xinyongzichan = pd.read_excel(filename, "信用资产底表", engine="openpyxl")[['资产账户', '警戒线比例', '平仓线比例', '负债总额', '担保资产', '个人维持担保比例值(不含场外资产)']].drop_duplicates()
    
    kehumingcheng = xinyongzhengquan[['资产账户', '客户姓名']].drop_duplicates()
    data1 = data.merge(kehumingcheng, on='资产账户', how='left') # lookup客户名称
    data2 = data1.merge(xinyongzichan, on='资产账户', how='left') # lookup'资产账户', '警戒线比例', '平仓线比例', '负债总额', '担保资产', '个人维持担保比例值(不含场外资产)'
    data2.rename(columns={'警戒线比例': '预警线', '平仓线比例': '平仓线', '个人维持担保比例值(不含场外资产)': '维保比例'},
                    inplace=True)

    zhengquan = zhengquan_raw.merge(gongyunjiage, on="证券代码", how="left")
    zhengquan['证券市值_modified'] = zhengquan.apply(lambda row: row['证券市值'] if np.isnan(row['公允价格']) else row['公允价格'] * row['当前数量'], axis=1)
    chicangshizhi = zhengquan.groupby('资产账户').agg({'证券市值_modified': 'sum'})
    data3 = data2.merge(chicangshizhi, on='资产账户', how='left').rename(columns={'证券市值_modified': '创业板持仓市值'})

    data3['创业板持仓集中度'] = data3['创业板持仓市值'] / data3['担保资产']

    jizhongdu = xinyongzhengquan.groupby('资产账户').agg({'证券集中度': 'max'})
    data4 = data3.merge(jizhongdu, on='资产账户', how='left').rename(columns={'证券集中度': '创业板持仓最大集中度'})

    def my_agg(group):
        df = group[['证券集中度', '证券名称']].groupby('证券集中度')['证券名称'].apply(','.join).reset_index()
        rtn = df[df['证券集中度'] == df['证券集中度'].max()]['证券名称']
        return rtn.iloc[0]

    zuidazhengquan = xinyongzhengquan[['资产账户', '证券集中度', '证券名称']].groupby('资产账户') \
                    .apply(my_agg).to_frame().reset_index()

    data5 = data4.merge(zuidazhengquan, on='资产账户', how='left').rename(columns={0: "创业板持仓最大集中度对应证券"})

    data5['跌20%后担保资产'] = data5['担保资产'] - data5['创业板持仓市值'] * 0.2
    data5['跌20%后维保比例'] = data5['跌20%后担保资产'] / data5['负债总额']
    data5['是否低于预警线1'] = data5.apply(lambda row: '' if not np.isfinite(row['跌20%后维保比例']) \
                                            else '是' if row['跌20%后维保比例'] < row['预警线'] else '否', axis = 1)
    data5['是否低于平仓线1'] = data5.apply(lambda row: '' if not np.isfinite(row['跌20%后维保比例']) \
                                            else '是' if row['跌20%后维保比例'] < row['平仓线'] else '否', axis = 1)

    data5['连续两次跌20%后担保资产'] = data5['担保资产'] - data5['创业板持仓市值'] * 0.2 - data5['创业板持仓市值'] * 0.8 * 0.2
    data5['连续两次跌20%后维保比例'] = data5['连续两次跌20%后担保资产'] / data5['负债总额']
    data5['是否低于预警线2'] = data5.apply(lambda row: '' if not np.isfinite(row['连续两次跌20%后维保比例']) \
                                            else '是' if row['连续两次跌20%后维保比例'] < row['预警线'] else '否', axis = 1)
    data5['是否低于平仓线2'] = data5.apply(lambda row: '' if not np.isfinite(row['连续两次跌20%后维保比例']) \
                                            else '是' if row['连续两次跌20%后维保比例'] < row['平仓线'] else '否', axis = 1)

    data5.replace([np.inf, -np.inf], np.nan, inplace=True)
    data5.fillna('', inplace=True)

    writer = pd.ExcelWriter('汇总表.xlsx')
    data5.to_excel(writer, index=False, encoding='utf-8',sheet_name='Sheet1')
    writer.save()

if __name__ == '__main__':
    process()