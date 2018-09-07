# -*- coding: gbk -*-

import datetime as dt
import pandas as pd
import os
import shutil
import utilities as ut
import numpy as np
# from WindPy import *



root_path = "//shiming/accounts/python/YZJ/accounts"
root_path_first = "//shiming/accounts"
all_products = [0, 1, 2, 4, 5, 7, 8, 10, 11, 12, 16, 17, 19, 101, 102, 107, 203, 208, 301, 309, 311, 312, 313] #每日要更的accounts


class TradingData:#这个类记录了各种common trading data，如期指价格，股票return等

    def __init__(self, date=None):
        self.trading_days = list(pd.read_csv('//SHIMING/accounts/python/trading_day_list.csv', names=None).iloc[:,0])
        print self.trading_days
        self.root_path = root_path

        #date格式：'2016-04-29'
        if date:
            self.date = date
        else:
            self.date = self.get_date()
        print 'today', self.date

        self.yesterday = self.trading_days[self.trading_days.index(self.date)-1]
        print 'yesterday', self.yesterday

        self.setting_accounts_folder()
        self.set_return_folder()
        self.copy_return()
        self.check_common_data()

        self.stock_return = self.get_stock_return()
        self.stock_pre_return = self.get_stock_pre_return()
        self.delivery_month = self.get_delivery_month()
        self.futpx = self.get_futpx()
        ut.update_yangyang_index(self.date)

    def get_date(self, date = dt.datetime.today().date()):
        #假如今天不是交易日，把日期调至最进的一个交易日

        date = str(date)
        while date not in self.trading_days:
            date = str(date-dt.timedelta(1))
        return date

    def setting_accounts_folder(self):
        for product_id in all_products:
            new_path = "{}/accounts-{}/{}".format(root_path, product_id, self.date.replace('-',''))
            if not os.path.exists(new_path):
                os.mkdir(new_path)



    def check_common_data(self):
        # sources = ["accounts.xlsx","futNum.xlsx","target_td.csv","holding.csv","trading_record.csv","index_hedge.csv"]
        # destinations = ["accounts.xlsx","delivery_month.csv","date_lock.csv","futpx.xlsx","futNum.xlsx","return_pre.csv", "target_pre.csv", "holding_pre.csv", "trading_record_pre.csv", "index_hedge_pre.csv"]
        common_file = ["delivery_month.csv", "future_wind.csv", "return.csv"]
        for fn in common_file:
            if not os.path.exists("{}/python/common_data/{}".format(root_path_first, fn)):
                print fn, " not exist! Program exit!"
                exit()

    def get_stock_return(self):
        return pd.read_csv("{}/python/common_data/return/{}/return.csv".format(root_path_first,self.date.replace('-',''))).fillna(0)

    def get_stock_pre_return(self):
        return pd.read_csv("{}/python/common_data/return/{}/return.csv".format(root_path_first,self.yesterday.replace('-',''))).fillna(0)

    def get_delivery_month(self):
        dm = pd.read_csv("//shiming/accounts/python/common_data/delivery_month.csv")
        dm.delivery = dm.delivery.astype(str)
        dm.pre_delivery = dm.pre_delivery.astype(str)
        return dm

    def get_futpx(self):
        futpx = pd.read_csv("//shiming/accounts/python/common_data/future_wind.csv")
        futpx['multiply'] = futpx['type'].apply(lambda x: 300 if (x == 'IF') or (x == 'IH') else 200)
        futpx['contract_name'] = futpx['id'].apply(lambda x: ut.set_contract_name(x))
        return futpx

    def copy_files(self):
        for product_id in all_products:
            from_path = "//shiming/trading/target_files/{}/".format(product_id)
            to_path = u"//shiming/accounts/Happy trading day/target/{}号/{}/".format(product_id, self.date)
            dirpath,dirnames,filenames = next(os.walk(from_path))
            for file in filenames:
                shutil.copy(from_path+file, (to_path+file).encode('gbk'))#只是做备份用，还是将trading下面的文件读进内存，这步在第一层里做

    def set_return_folder(self):
        return_folder_path = '//shiming/accounts/python/common_data/return/{}'.format(self.date.replace('-',''))
        if not os.path.exists(return_folder_path):
            os.mkdir(return_folder_path)

    def copy_return(self):
        from_path = "//shiming/accounts/return.csv"
        to_path = u"//shiming/accounts/python/common_data/return/{}/return.csv".format(self.date.replace('-',''))
        if not os.path.exists(to_path):
            shutil.copy(from_path, to_path.encode('gbk'))
        to_path = u"//shiming/accounts/python/common_data/return.csv".format(self.date.replace('-', ''))
        shutil.copy(from_path, to_path.encode('gbk'))


trading_data = TradingData('2017-02-20')

# print trading_data.__dict__
# products_trans = [1,2,12,13]
class Transform:
    def __init__(self, product_id):
        self.t = trading_data
        self.product_id = product_id
        # os.chdir("{}/{}hao".format(root_path, product_id))
        self.manual_path = "{}/{}hao".format(root_path_first, product_id)
        self.manual_transform_path = "{}/{}hao/transform".format(root_path_first, product_id)
        self.path = "{}/accounts-{}/{}".format(root_path, product_id, self.t.date.replace('-',''))

        # 这里是标准的六张表存储格式，包括列名与顺序
        self.column_name = {}
        self.column_name['asset'] = ['Frozen','Avaliable','MarketCap','TotalAsset']
        self.column_name['holding'] = ['Id','Holding','Price','TotalCost','FloatProfit']
        self.column_name['trade'] = ['Id', 'Price', 'Volume', 'Commission', 'Direction']
        self.column_name['asset_fut'] = ['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap']
        self.column_name['holding_fut'] = ['Id','Direction','Holding']#direction应该是buy, sell
        self.column_name['trade_fut'] = ['Id','Direction','Commission','Price','Volume']#direction应该是buy, sell

    # 默认跟迅投系统下下来的格式一样 (product 0,5,7,11)
    # 如果是excel文件，输入参数sep = '' or None
    # 函数T是一个转换模板，通过调usecols，所有asset, holding, trade都可以用这个来transform
    # 可能需要在一些列加上.astype(int)（如“成交数量”）
    def T(self, input_file='trade.csv', sep=',', usecols=None, dir_path=None, process = 1):
        if isinstance(input_file, (tuple, list)):
            for single_file in input_file:
                self.T(input_file=single_file, sep=sep, usecols=usecols, dir_path=dir_path, process = process)
            return
        if input_file.split('.')[0] == 'trade1':
            table_name = 'trade'
        else:
            table_name = input_file.split('.')[0]


        if dir_path:
            input_file = "{}/{}".format(dir_path, input_file)
        else:
            input_file = "{}/{}".format(self.manual_path, input_file)


        if sep:
            df = pd.read_csv(input_file, sep=sep, index_col=False, encoding='gbk')
        else:
            df = pd.read_excel(input_file, encoding='gbk')

        df = ut.check_dataframe(df)
        if u'成交状态' in df.columns:
            df = df[df[u'成交状态'] != u'撤单成交']

        if usecols:
            df = df.loc[:, usecols]

        df = df.fillna(0)  #将na默认为0

        if process == 1:
            if len(df.columns) != len(self.column_name[table_name]):
                df = df.iloc[:, range(len(self.column_name[table_name]))]
            df.columns = self.column_name[table_name]

            # 对于trade，trade_fut，holding_fut中的direction做买卖多空判断
            if 'Direction' in df.columns:
                df['Direction'] = df['Direction'].apply(ut.standard_csv_direction)
        else:
            if len(df.columns) != len(self.column_name[table_name]):
                df = df.reindex(columns=self.column_name[table_name])

        if 'Id' in df.columns:
            df = df[df['Id'] != 0]

        if input_file == 'holding.csv':
            df = df[0 < df['Id'] < 1000000]

        if 'Holding' in df.columns:
            # # st1 = u'股'
            # # st2 = u'份'
            # # df['Holding'] = df['Holding'].apply(lambda x: str(x).strip(st1))
            # # df['Holding'] = df['Holding'].apply(lambda x: str(x).strip(st2))
            # print df
            #
            # df['Holding'] = df['Holding'].apply(lambda x: float(x))
            df['Holding'] = df['Holding'].astype(int)
        # if 'Volume' in df.columns:
        #     print df['Volume']
        #     st1 = u'股'
        #     st2 = u'份'.encode(encoding='gbk')
        #     print st1
        #     df['Volume'] = df['Volume'].apply(lambda x: str(x).strip(st1))
        #     df['Volume'] = df['Volume'].apply(lambda x: str(x).strip(st2))
        #     df['Volume'] = df['Volume'].apply(lambda x: float(x))
        #     df['Volume'] = df['Volume'].astype(int)

        if input_file == 'holding.csv':
            df['int_type'] = df['Id'].apply(lambda x: ut.find_type(x))
            df = df[df['int_type'] == 1]
            del df['int_type']
            df['Id'] = df['Id'].apply(lambda x: int(x))


        if table_name.endswith('fut') and table_name != 'asset_fut':    # 排除商品期货,只留股指期货
            df['stock'] = df['Id'].apply(lambda x: 1 if (x[:2] == 'IF') or (x[:2] == 'IH') or (x[:2] == 'IC')else 0)
            df = df[df['stock'] == 1]

        df.to_csv("{}/{}.csv".format(self.path, table_name), encoding='gbk', index=False)


    def transform_cost(self, input_file):
        input_path = self.path+"/"+input_file
        df = pd.read_csv(input_path, encoding='gbk')
        print df
        df.TotalCost = df.TotalCost*df.Holding
        df.to_csv(input_path, encoding='gbk', index=False)




    #1号的holding比较特殊，要单独重写
    def holding_zhongxin_security_1(self, input_file='holding.xls', sep='\t', asset_columns=[u'冻结金额',u'可用',u'参考市值',u'资产'], holding_columns=[u'证券代码',u'参考持股',u'当前价',u'当前成本',u'浮动盈亏']):
        input_path = "{}/{}".format(self.manual_path, input_file)
        asset = pd.read_csv(input_path, sep=sep, encoding='gbk', nrows=1)#, usecols=asset_columns

        asset = ut.check_dataframe(asset).loc[:, asset_columns].fillna(0)
        asset.columns = self.column_name['asset']
        asset.to_csv("{}/asset.csv".format(self.path), encoding='gbk', index=False)

        holding = pd.read_csv(input_path, sep=sep, encoding='gbk', skiprows=3)#, usecols=holding_columns
        holding = ut.check_dataframe(holding)
        holding = holding.loc[:, holding_columns].fillna(0)
        holding.columns=self.column_name['holding']

        holding['int_type'] = holding['Id'].apply(lambda x: ut.find_type(x))
        holding = holding[holding['int_type'] == 1]
        del holding['int_type']
        holding['Id'] = holding['Id'].apply(lambda x: int(x))


        holding.to_csv("{}/holding.csv".format(self.path), encoding='gbk', index=False)

    def holding_zhongxin_security_17_low(self, input_file='holding_l.xls', sep='\t', asset_columns=[u'冻结金额', u'可用', u'参考市值', u'资产'], holding_columns=[u'证券代码',u'证券数量',u'当前价',u'摊簿成本价',u'实现盈亏']):
        input_path = "{}/{}".format(self.manual_path, input_file)
        asset = pd.read_csv(input_path, sep=sep, encoding='gbk', nrows=1)  # , usecols=asset_columns
        asset = ut.check_dataframe(asset).loc[:, asset_columns].fillna(0)
        asset.columns = self.column_name['asset']
        asset.to_csv("{}/asset_l.csv".format(self.manual_transform_path), encoding='gbk', index=False)

        holding = pd.read_csv(input_path, sep=sep, encoding='gbk', skiprows=3)  # , usecols=holding_columns
        holding = ut.check_dataframe(holding)
        holding = holding.loc[:, holding_columns].fillna(0)
        holding.columns = self.column_name['holding']

        holding['int_type'] = holding['Id'].apply(lambda x: ut.find_type(x))
        holding = holding[holding['int_type'] == 1]
        del holding['int_type']
        holding['Id'] = holding['Id'].apply(lambda x: int(x))


        holding.to_csv("{}/holding_l.csv".format(self.manual_transform_path), encoding='gbk', index=False)

    def transform_trade_17(self, input_file='trade_l.xls', sep='\t', columns=[u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖标志"]):
        input_path = "{}/{}".format(self.manual_path, input_file)
        df = pd.read_csv(input_path, sep=sep, encoding='gbk')
        df = ut.check_dataframe(df).loc[:, columns].fillna(0)
        df.to_csv("{}/trade_l.csv".format(self.manual_transform_path), encoding='gbk', index=False)

    def holding_zhongxin_security_normal(self, input_file='holding.xls', sep='\t', asset_columns=[u'冻结金额', u'可用', u'参考市值', u'资产'], holding_columns=[u'证券代码', u'证券数量', u'当前价', u'成本价', u'浮动盈亏']):
        input_path = "{}/{}".format(self.manual_path, input_file)
        asset = pd.read_csv(input_path, sep=sep, encoding='gbk', nrows=1)  # , usecols=asset_columns
        asset = ut.check_dataframe(asset).loc[:, asset_columns].fillna(0)
        asset.columns = self.column_name['asset']
        asset.to_csv("{}/asset.csv".format(self.manual_transform_path), encoding='gbk', index=False)

        holding = pd.read_csv(input_path, sep=sep, encoding='gbk', skiprows=3)  # , usecols=holding_columns
        holding = ut.check_dataframe(holding)
        holding = holding.loc[:, holding_columns].fillna(0)
        holding.columns = self.column_name['holding']

        holding['int_type'] = holding['Id'].apply(lambda x: ut.find_type(x))
        holding = holding[holding['int_type'] == 1]
        del holding['int_type']
        holding['Id'] = holding['Id'].apply(lambda x: int(x))


        holding.to_csv("{}/holding.csv".format(self.manual_transform_path), encoding='gbk', index=False)



    def transform_asset_fut_normal(self, columns):
        input_file = 'asset_fut.txt'
        input_path = "{}/{}".format("//shiming/accounts/{}hao".format(str(self.product_id)), input_file)
        f = open(input_path, 'r')
        lines = f.readlines()
        a = float(lines[14].strip().split(' ')[-1].replace(',', ''))
        b = float(lines[15].strip().split(' ')[-1].replace(',', ''))
        c = float(lines[22].strip().split(' ')[-1].replace(',', ''))
        df = pd.DataFrame(np.random.randn(1, 5), index=[0], columns=columns)
        df.ix[0, 0] = a
        df.ix[0, 1] = b
        df.ix[0, 2] = c
        df.ix[0, 3] = 0
        df.ix[0, 4] = 0
        df.to_csv("{}/asset_fut.csv".format(self.manual_transform_path), encoding='gbk', index=False)

    def transform_asset_fut_normal_l(self, columns):
        input_file = 'asset_fut_l.txt'
        input_path = "{}/{}".format("//shiming/accounts/{}hao".format(str(self.product_id)), input_file)
        f = open(input_path, 'r')
        lines = f.readlines()
        a = float(lines[14].strip().split(' ')[-1].replace(',', ''))
        b = float(lines[15].strip().split(' ')[-1].replace(',', ''))
        c = float(lines[22].strip().split(' ')[-1].replace(',', ''))
        df = pd.DataFrame(np.random.randn(1, 5), index=[0], columns=columns)
        df.ix[0, 0] = a
        df.ix[0, 1] = b
        df.ix[0, 2] = c
        df.ix[0, 3] = 0
        df.ix[0, 4] = 0
        df.to_csv("{}/asset_fut_l.csv".format(self.manual_transform_path), encoding='gbk', index=False)


    def holding_zhongxin_security_17(self, input_file='holding.xls', asset_columns=[u'冻结金额', u'可用', u'参考市值', u'总资产'], holding_columns=[u'证券代码', u'拥股数量', u'最新价', u'盈亏成本', u'浮动盈亏']):
        input_path = "{}/{}".format(self.manual_path, input_file)
        asset = pd.read_excel(input_path, skiprows=5).iloc[[0],:]  # , usecols=asset_columns
        asset = ut.check_dataframe(asset).loc[:, asset_columns].fillna(0)
        asset.columns = self.column_name['asset']
        asset.ix[0, 1] = str(asset.ix[0, 1]).replace(',','')
        asset.to_csv("{}/asset.csv".format(self.manual_transform_path), encoding='gbk', index=False)
        holding = pd.read_excel(input_path, skiprows=8)      #这里注释掉是因为17号下载holding.xls数据格式不稳定

        # holding = pd.read_excel(input_path)  # , usecols=holding_columns
        l = len(holding)
        holding = holding.drop(l-1, axis=0)
        holding = ut.check_dataframe(holding)
        holding = holding.loc[:, holding_columns].fillna(0)
        holding.columns = self.column_name['holding']

        holding['int_type'] = holding['Id'].apply(lambda x: ut.find_type(x))
        holding = holding[holding['int_type'] == 1]
        del holding['int_type']
        holding['Id'] = holding['Id'].apply(lambda x: int(x))


        holding.to_csv("{}/holding.csv".format(self.manual_transform_path), encoding='gbk', index=False)

    def transform_17(self, input_file='trade.xls', trade_columns=[u"证券代码", u"成交价格", u"成交数量", u"手续费", u"方向"]):
        input_path = "{}/{}".format(self.manual_path, input_file)
        trade = pd.read_excel(input_path, skiprows=5, encoding='gbk')  # , usecols=asset_columns
        trade = ut.check_dataframe(trade)
        trade = trade[trade[u'成交类型'] != u'撤单']

        trade = trade.loc[:, trade_columns].fillna(0)

        trade.columns = self.column_name['trade']

        if 'Direction' in trade.columns:
            trade['Direction'] = trade['Direction'].apply(ut.standard_csv_direction)
        trade.to_csv("{}/trade.csv".format(self.manual_transform_path), encoding='gbk', index=False)

    def transform_normal(self, input_file='trade.xls', trade_columns=[u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖标志"]):
        input_path = "{}/{}".format(self.manual_path, input_file)
        trade = pd.read_csv(input_path, sep='\t', encoding='gbk')  # , usecols=asset_columns
        trade = ut.check_dataframe(trade)
        trade = trade.loc[:, trade_columns].fillna(0)
        trade.columns = self.column_name['trade']

        if 'Direction' in trade.columns:
            trade['Direction'] = trade['Direction'].apply(ut.standard_csv_direction)
        trade.to_csv("{}/trade.csv".format(self.manual_transform_path), encoding='gbk', index=False)

    def transform_holding_fut(self, input_file, holding_columns, l=0):
        input_path = "{}/{}".format(self.manual_path, input_file)
        holding = pd.read_csv(input_path, encoding='gbk')  # , usecols=asset_columns
        # print input_path
        # print holding
        holding = ut.check_dataframe(holding).loc[:, holding_columns].fillna(0)
        holding.columns = self.column_name['holding_fut']
        # print holding
        if 'Direction' in holding.columns:
            holding['Direction'] = holding['Direction'].apply(ut.standard_csv_direction)

        holding['stock'] = holding['Id'].apply(lambda x: 1 if (x[:2] == 'IF') or (x[:2] == 'IH') or (x[:2] == 'IC')else 0)   # 排除商品期货,只留股指期货
        holding = holding[holding['stock'] == 1]
        if l == 0:
            holding.to_csv("{}/holding_fut.csv".format(self.manual_transform_path), encoding='gbk', index=False)
        else:
            holding.to_csv("{}/holding_fut_l.csv".format(self.manual_transform_path), encoding='gbk', index=False)

    def transform_trade_fut(self, input_file, trade_columns, l=0):
        input_path = "{}/{}".format(self.manual_path, input_file)
        trade = pd.read_csv(input_path, encoding='gbk')  # , usecols=asset_columns
        trade = ut.check_dataframe(trade).loc[:, trade_columns].fillna(0)
        trade.columns = self.column_name['trade_fut']
        if 'Direction' in trade.columns:
            trade['Direction'] = trade['Direction'].apply(ut.standard_csv_direction)
        trade['stock'] = trade['Id'].apply(lambda x: 1 if (x[:2] == 'IF') or (x[:2] == 'IH') or (x[:2] == 'IC')else 0)
        trade = trade[trade['stock'] == 1]
        if l == 0:
            trade.to_csv("{}/trade_fut.csv".format(self.manual_transform_path), encoding='gbk', index=False)
        else:
            trade.to_csv("{}/trade_fut_l.csv".format(self.manual_transform_path), encoding='gbk', index=False)

    def transform_asset_fut17(self, columns):
        input_file = 'asset_fut.txt'
        input_path = "{}/{}".format(self.manual_path, input_file)
        f = open(input_path, 'r')
        lines = f.readlines()
        a = float(lines[14].strip().split(' ')[-1].replace(',',''))
        b = float(lines[15].strip().split(' ')[-1].replace(',',''))
        c = float(lines[22].strip().split(' ')[-1].replace(',',''))
        df = pd.DataFrame(np.random.randn(1,5), index=[0],columns=columns)
        df.ix[0, 0] = a
        df.ix[0, 1] = b
        df.ix[0, 2] = c
        df.ix[0, 3] = 0
        df.ix[0, 4] = 0
        df.columns = self.column_name['asset_fut']
        df.to_csv("{}/asset_fut.csv".format(self.manual_transform_path), encoding='gbk', index=False)


    def concat_low_to_high (self, type, file, file_l, columns_l, columns, dir_path, transform_path):
        input_path = "{}/{}".format(dir_path, file_l)
        df1 = pd.read_csv(input_path, encoding='gbk')
        df1 = ut.check_dataframe(df1).loc[:, columns_l].fillna(0)
        df1.columns = columns
        input_path = "{}/{}".format(dir_path, file)
        df =  pd.read_csv(input_path, encoding='gbk')
        df = ut.check_dataframe(df).fillna(0)
        df = pd.concat([df, df1])
        if type == 1:
            df['stock'] = df['Id'].apply(lambda x: 1 if (x[:2] == 'IF') or (x[:2] == 'IH') or (x[:2] == 'IC')else 0)
            df = df[df['stock'] == 1]
        df.to_csv("{}/{}".format(transform_path, file), encoding='gbk', index=False)

    def plus_low_to_high(self, file, file_l, columns_l, columns, dir_path, transform_path):
        input_path = "{}/{}".format(dir_path, file_l)
        df1 = pd.read_csv(input_path, encoding='gbk')
        df1 = ut.check_dataframe(df1).loc[:, columns_l].fillna(0)
        df1.columns = columns
        l = len(columns)
        input_path = "{}/{}".format(dir_path, file)
        df = pd.read_csv(input_path, encoding='gbk')
        df = ut.check_dataframe(df).fillna(0)
        for i in range(l):
            df.ix[0, i] = df.ix[0, i] + df1.ix[0, i]
        df.to_csv("{}/{}".format(transform_path, file), encoding='gbk', index=False)

    # def holding_zhongxin_security_11(self, input_file='holding.xls', sep='\t', asset_columns=[u'冻结金额', u'可用金额', u'总市值', u'总资产'], holding_columns=[u"证券代码", u"当前拥股", u"最新价", u"成本价", u"盈亏"]):
    #     input_path = "{}/{}".format("//shiming/accounts/CSVDownloader/Download/XT11", input_file)
    #     asset = pd.read_csv(input_path, sep=sep, encoding='gbk', nrows=1)  # , usecols=asset_columns
    #     asset = ut.check_dataframe(asset).loc[:, [u'冻结金额', u'可用', u'参考市值', u'资产']].fillna(0)
    #     asset.columns = asset_columns
    #     asset.to_csv("{}/asset_l.csv".format("//shiming/accounts/CSVDownloader/Download/XT11"), encoding='gbk', index=False)

        # holding = pd.read_csv(input_path, sep=sep, encoding='gbk', skiprows=3)  # , usecols=holding_columns
        # holding = ut.check_dataframe(holding)
        # holding = holding.loc[:, [u"证券代码", u"参考持股", u"当前价", u"成本价", u"浮动盈亏"]].fillna(0)
        # holding.columns = holding_columns
        # holding[u"当前拥股"] = holding[u"当前拥股"].apply(lambda x: str(int(x)))
        # holding.to_csv("{}/holding_l.csv".format("//shiming/accounts/CSVDownloader/Download/XT11"), encoding='gbk', index=False)

    # def transform_trade_11(self, input_file='trade.xls', sep='\t', columns=[u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖"]):
    #     input_path = "{}/{}".format("//shiming/accounts/CSVDownloader/Download/XT11", input_file)
    #     df =pd.read_csv(input_path, sep=sep, encoding='gbk')
    #     df = ut.check_dataframe(df).loc[:, columns].fillna(0)
    #     df.to_csv("{}/trade_l.csv".format("//shiming/accounts/CSVDownloader/Download/XT11"), encoding='gbk', index=False)
    #
    # def transform_asset_fut11(self, columns):
    #     input_file = 'asset_fut_l.txt'
    #     input_path = "{}/{}".format("//shiming/accounts/CSVDownloader/Download/XT11", input_file)
    #     f = open(input_path, 'r')
    #     lines = f.readlines()
    #     a = float(lines[14].strip().split(' ')[-1].replace(',', ''))
    #     b = float(lines[15].strip().split(' ')[-1].replace(',', ''))
    #     c = float(lines[22].strip().split(' ')[-1].replace(',', ''))
    #     df = pd.DataFrame(np.random.randn(1, 5), index=[0], columns=columns)
    #     df.ix[0, 0] = a
    #     df.ix[0, 1] = b
    #     df.ix[0, 2] = c
    #     df.ix[0, 3] = 0
    #     df.ix[0, 4] = 0
    #     df.to_csv("{}/asset_fut_l.csv".format("//shiming/accounts/CSVDownloader/Download/XT11"), encoding='gbk', index=False)
    #
    # def plus_holding_11hao(self):
    #     holding = pd.read_csv("//shiming/accounts/CSVDownloader/Download/XT11/holding.csv", encoding='gbk')
    #     holding_l = pd.read_csv("//shiming/accounts/CSVDownloader/Download/XT11/holding_l.csv", encoding='gbk')
    #     c = [u"证券代码", u"当前拥股", u"最新价", u"成本价", u"盈亏"]
    #     columns = [1, 2, 3, 4, 5]
    #     holding.columns = columns
    #     holding_l.columns = columns
    #     l = list(holding[1])
    #     l1 = list(holding_l[1])
    #     for id in l1:
    #         if id in l:
    #             a = int(holding[holding[1] == id][2])
    #             b = int(holding_l[holding_l[1] == id][2])
    #             a = a + b
    #             i = holding[holding[1] == id].index
    #             holding.loc[i, 2] = a
    #     for id in l1:
    #         if id not in l:
    #             df = holding_l[holding_l[1] == id]
    #             holding = pd.concat([holding, df])
    #     holding.columns = c
    #     holding.to_csv("{}/holding.csv".format("//shiming/accounts/CSVDownloader/Download/XT11"), encoding='gbk', index=False)
    #     holding.to_csv("{}/holding_all.csv".format("//shiming/accounts/CSVDownloader/Download/XT11"), encoding='gbk', index=False)

    # def plus_holding_7hao(self):
    #     holding = pd.read_csv("//shiming/accounts/CSVDownloader/Download/XT7/holding.csv", encoding='gbk')
    #     holding_l = pd.read_csv("//shiming/accounts/CSVDownloader/Download/XT7/holding_l.csv", encoding='gbk')
    #     c = [u"证券代码", u"当前拥股", u"最新价", u"成本", u"盈亏"]
    #     columns = [1, 2, 3, 4, 5]
    #     columns_l = [u"证券代码", u"当前拥股", u"最新价", u"持仓成本", u"盈亏"]
    #     holding.columns = columns
    #     holding_l = holding_l.loc[:, columns_l].fillna(0)
    #     holding_l.columns = columns
    #     l = list(holding[1])
    #     l1 = list(holding_l[1])
    #     for id in l1:
    #         if id in l:
    #             a = int(holding[holding[1] == id][2])
    #             b = int(holding_l[holding_l[1] == id][2])
    #             a = a + b
    #             i = holding[holding[1] == id].index
    #             holding.loc[i, 2] = a
    #     for id in l1:
    #         if id not in l:
    #             df = holding_l[holding_l[1] == id]
    #             holding = pd.concat([holding, df])
    #     holding.columns = c
    #     holding.to_csv("{}/holding.csv".format("//shiming/accounts/CSVDownloader/Download/XT7"), encoding='gbk', index=False)
    #     holding.to_csv("{}/holding_all.csv".format("//shiming/accounts/CSVDownloader/Download/XT7"), encoding='gbk', index=False)

    def plus_holding_17hao(self):
        holding = pd.read_csv("//shiming/accounts/17hao/transform/holding.csv", encoding='gbk')
        holding_l = pd.read_csv("//shiming/accounts/17hao/transform/holding_l.csv", encoding='gbk')
        c = holding.columns
        columns = [1, 2, 3, 4, 5]
        holding.columns = columns
        holding_l.columns = columns
        l = list(holding[1])
        l1 = list(holding_l[1])
        for id in l1:
            if id in l:
                a = int(holding[holding[1] == id][2])
                b = int(holding_l[holding_l[1] == id][2])
                a = a + b
                i = holding[holding[1] == id].index
                holding.loc[i, 2] = a
        for id in l1:
            if id not in l:
                df = holding_l[holding_l[1] == id]
                holding = pd.concat([holding, df])
        holding.columns = c
        holding.to_csv("{}/holding.csv".format(self.manual_transform_path), encoding='gbk', index=False)
        holding.to_csv("{}/holding_all.csv".format(self.manual_transform_path), encoding='gbk', index=False)
